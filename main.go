package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strings"
	"sync"
	"time"
)

type RequestPayload struct {
	Filename       string               `json:"filename"`
	Template       string               `json:"template"`
	Replacements   map[string]string    `json:"replacements"`
	IncludedBlocks [][]any              `json:"includedBlocks"`
	Datatables     map[string]TableData `json:"datatables"`
}

type TableData map[string][]string

type docxResult struct {
	path string
	hash string
}

const (
	templateDir  = "./template"
	pythonScript = "./docx_processor.py"
	outputDir    = "./output"
)

// CHANGE IF LIBREOFFICE INSTALLED ELSEWHERE
// Windows
// var sofficePath = "C:/Program Files/LibreOffice/program/soffice.exe"
// Linux
// var sofficePath = "/usr/bin/soffice"
var sofficePath = os.Getenv("SOFFICE_PATH")
var pythonCmd = os.Getenv("PYTHON_CMD")

var libreOfficeSem = make(chan struct{}, 1)

func main() {
	os.MkdirAll(outputDir, 0755)
	http.HandleFunc("/documents/generate", cors(documentsHandler))
	log.Println("Listening on :8088")
	log.Fatal(http.ListenAndServe(":8088", nil))
}
func cors(fn http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {

		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusOK)
			return
		}

		fn(w, r)
	}
}
func generateDocx(payload *RequestPayload) (string, string, error) {

	finalFilename := sanitizeFilename(payload.Filename)
	hashPath := filepath.Join(outputDir, finalFilename+"-docx.hash")
	docxPath := filepath.Join(outputDir, finalFilename+".docx")

	if filepath.Ext(payload.Template) != ".docx" {
		return "", "", fmt.Errorf("error wrong template, only .docx accepted")
	}
	templatePath := filepath.Join(templateDir, payload.Template)

	jsonBytes, err := canonicalJSON(payload)
	if err != nil {
		return "", "", fmt.Errorf("error canonicalJSON : %v", err)
	}

	hash := sha256Hex(jsonBytes)

	same, err := isSameContent(hashPath, hash)
	if err != nil {
		return "", "", fmt.Errorf("error isSameContent : %v", err)
	}
	fmt.Println(mustJSON(payload.Datatables))
	if !same {
		cmd := exec.Command(
			pythonCmd,
			pythonScript,
			templatePath,
			docxPath,
			mustJSON(payload.Replacements),
			mustJSON(payload.IncludedBlocks),
			mustJSON(payload.Datatables),
		)

		if out, err := cmd.CombinedOutput(); err != nil {
			return "", "", fmt.Errorf("%v : %v", err, out)
		}

		if err := saveHash(hashPath, hash); err != nil {
			return "", "", fmt.Errorf("error saveHash : %v", err)
		}
	}

	return docxPath, hash, nil
}

func documentsHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "POST only", http.StatusMethodNotAllowed)
		return
	}

	var payloads []RequestPayload
	if err := json.NewDecoder(r.Body).Decode(&payloads); err != nil {
		http.Error(w, "invalid JSON body", http.StatusBadRequest)
		return
	}

	// For Loop ///////////////////////////////////////////////////////
	// zipFiles := make(map[string]string)
	// for _, payload := range payloads {
	// 	files, err := generateDocx(&payload)
	// 	maps.Copy(zipFiles, files)

	// 	if err != nil {
	// 		http.Error(w, err.Error(), http.StatusInternalServerError)
	// 		return
	// 	}
	// }

	/* -------- DOCX WORKER POOL -------- */

	docxJobs := make(chan RequestPayload)
	docxOut := make(chan docxResult)
	docxToPDF := make(chan docxResult)
	pdfOut := make(chan string)
	allOut := make(chan string)

	errCh := make(chan error, 1)

	var docxWG sync.WaitGroup
	var pdfWG sync.WaitGroup

	numWorkers := 4
	for i := 0; i < numWorkers; i++ {
		docxWG.Add(1)
		go docxWorker(docxJobs, docxOut, errCh, &docxWG)
	}

	go func() {
		docxWG.Wait()
		close(docxOut)
	}()

	/* -------- DOCX TEE (ZIP + PDF) -------- */

	go func() {
		for docx := range docxOut {
			allOut <- docx.path
			docxToPDF <- docx
		}
		close(docxToPDF)
	}()

	/* -------- PDF WORKER -------- */

	pdfWG.Add(1)
	go pdfWorker(docxToPDF, pdfOut, errCh, &pdfWG)

	go func() {
		pdfWG.Wait()
		close(pdfOut)
	}()

	/* -------- PDF COLLECTOR -------- */

	go func() {
		for pdf := range pdfOut {
			allOut <- pdf
		}
		close(allOut)
	}()

	/* -------- FEED JOBS -------- */

	go func() {
		for _, p := range payloads {
			docxJobs <- p
		}
		close(docxJobs)
	}()

	/* -------- COLLECT OR ERROR -------- */

	var allFiles []string

	for {
		select {
		case err := <-errCh:
			if err != nil {
				http.Error(w, err.Error(), http.StatusInternalServerError)
				return
			}

		case f, ok := <-allOut:
			if !ok {
				goto DONE
			}
			allFiles = append(allFiles, f)
		}
	}

DONE:
	w.Header().Set("Content-Type", "application/zip")
	w.Header().Set(
		"Content-Disposition",
		"attachment; filename=\"Surat Tugas "+time.Now().Format("[2006-01-02 15_04]")+".zip\"",
	)

	if err := writeZipToResponse(w, allFiles); err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
	}
}

func docxWorker(
	jobs <-chan RequestPayload,
	out chan<- docxResult,
	errCh chan<- error,
	wg *sync.WaitGroup,
) {
	defer wg.Done()

	for p := range jobs {
		docxFile, h, err := generateDocx(&p)
		if err != nil {
			select {
			case errCh <- err:
			default:
			}
			return
		}
		out <- docxResult{path: docxFile, hash: h}
	}
}

func pdfWorker(
	docxIn <-chan docxResult,
	pdfOut chan<- string,
	errCh chan<- error,
	wg *sync.WaitGroup,
) {
	defer wg.Done()

	for docxFile := range docxIn {
		pdfFile, err := convertToPDFWithHash(docxFile.path, docxFile.hash)
		if err != nil {
			select {
			case errCh <- err:
			default:
			}
			return
		}

		pdfOut <- pdfFile
	}
}

func convertToPDFWithHash(docxPath string, docxHash string) (string, error) {
	base := filepath.Base(docxPath)
	name := base[:len(base)-len(filepath.Ext(base))]

	pdfPath := filepath.Join(outputDir, name+".pdf")
	pdfHashPath := filepath.Join(outputDir, name+"-pdf.hash")

	// PDF hash is derived from DOCX hash
	pdfHash := sha256Hex([]byte(docxHash))

	same, err := isSameContent(pdfHashPath, pdfHash)
	if err != nil {
		return "", fmt.Errorf("isSameContent(pdf): %w", err)
	}

	if !same {
		cmd := exec.Command(
			sofficePath,
			"--headless",
			"--convert-to", "pdf",
			"--outdir", outputDir,
			docxPath,
		)

		if out, err := cmd.CombinedOutput(); err != nil {
			return "", fmt.Errorf("soffice: %v | %s", err, out)
		}

		if err := saveHash(pdfHashPath, pdfHash); err != nil {
			return "", fmt.Errorf("saveHash(pdf): %w", err)
		}
	}

	return pdfPath, nil
}

func convertToPDF(docxPath string) (string, error) {
	libreOfficeSem <- struct{}{}
	defer func() { <-libreOfficeSem }()

	cmd := exec.Command(
		sofficePath,
		"--headless",
		"--convert-to", "pdf",
		"--outdir", outputDir,
		docxPath,
	)
	output, err := cmd.CombinedOutput()
	if err != nil {
		return "", fmt.Errorf("pdf error: %w, %s", err, output)
	}
	return ChangeExtension(docxPath, "pdf"), nil
}

func writeZipToResponse(w http.ResponseWriter, files []string) error {
	zipWriter := zip.NewWriter(w)
	defer zipWriter.Close()

	for _, filePath := range files {
		file, err := os.Open(filePath)
		if err != nil {
			return err
		}
		fileRel, err := filepath.Rel(outputDir, filePath)
		if err != nil {
			return err
		}
		fw, err := zipWriter.Create(fileRel)
		if err != nil {
			file.Close()
			return err
		}
		if _, err := io.Copy(fw, file); err != nil {
			file.Close()
			return err
		}

		file.Close()
	}

	return nil
}
func mustJSON(v any) string {
	b, _ := json.Marshal(v)
	return string(b)
}

func sanitizeFilename(filename string) string {
	// Define a regex pattern for common forbidden characters in filenames across different OSes.
	// The pattern matches any character NOT in the allowed set:
	// a-z, A-Z, 0-9, period (.), underscore (_), and dash (-).
	// Note: this is a strict approach. You might want to allow spaces or other chars depending on needs.
	re := regexp.MustCompile(`[^a-zA-Z0-9._-]`)

	// Replace all matches with an underscore
	sanitized := re.ReplaceAllString(filename, "_")

	// Optional: further sanitization for specific edge cases or simply replace all forbidden ones
	// A more direct approach to target specific forbidden chars for Windows
	forbiddenChars := `\/:*?"<>|`
	for _, char := range forbiddenChars {
		sanitized = strings.ReplaceAll(sanitized, string(char), "_")
	}

	return sanitized
}

func canonicalJSON(v any) ([]byte, error) {
	buf := &bytes.Buffer{}
	enc := json.NewEncoder(buf)
	enc.SetEscapeHTML(false)
	enc.SetIndent("", "")
	if err := enc.Encode(v); err != nil {
		return nil, err
	}
	return bytes.TrimSpace(buf.Bytes()), nil
}

func sha256Hex(data []byte) string {
	sum := sha256.Sum256(data)
	return hex.EncodeToString(sum[:])
}

func isSameContent(hashPath string, newHash string) (bool, error) {
	oldHash, err := os.ReadFile(hashPath)
	if err != nil {
		if os.IsNotExist(err) {
			return false, nil
		}
		return false, err
	}
	return bytes.Equal(bytes.TrimSpace(oldHash), []byte(newHash)), nil
}

func saveHash(hashPath, hash string) error {
	return os.WriteFile(hashPath, []byte(hash), 0644)
}

func ChangeExtension(path, newExt string) string {
	// 1. Get the current extension
	currentExt := filepath.Ext(path)

	// 2. Remove the current extension from the path
	// strings.TrimSuffix is used to ensure only the trailing extension is removed.
	baseName := strings.TrimSuffix(path, currentExt)

	// 3. Append the new extension
	// Ensure the new extension starts with a dot if it doesn't already
	if !strings.HasPrefix(newExt, ".") && newExt != "" {
		newExt = "." + newExt
	}

	return baseName + newExt
}
