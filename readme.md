# Docx Batch Service

An HTTP service that generates official letters (_surat tugas_, SPK, BAST, etc.) in bulk from `.docx` templates. Send a batch of JSON jobs, get back a single ZIP containing the generated DOCX files and their PDF conversions.

The service is a Go HTTP server that orchestrates two things:

- **`docx_processor.py`** — a Python script (via [python-docx](https://python-docx.readthedocs.io/)) that fills the template: placeholder replacement, optional block inclusion/removal, and repeating table rows.
- **LibreOffice** (`soffice --headless`) — converts each generated DOCX into PDF.

Results are cached by content hash, so re-sending an identical job skips both the Python run and the PDF conversion.

---

## Features

- **Batch generation** — one request, many documents, processed by a pool of 4 concurrent DOCX workers.
- **Formatting-preserving placeholder replacement** — placeholders split across multiple Word runs are handled correctly, and character formatting (bold/italic/underline/font/size) is carried over to the replaced text.
- **Body, table, header and footer coverage** — placeholders are replaced everywhere, including first-page headers/footers.
- **Conditional blocks** — mark a region of the template with start/end markers and include or drop it per document.
- **Dynamic tables** — expand a marker row into as many rows as your data has.
- **Content-hash caching** — identical payloads reuse the previously generated DOCX and PDF.
- **ZIP streaming response** — files are written straight into the HTTP response.

---

## Requirements

Running with Docker (recommended) requires only Docker and Docker Compose. For a local (non-Docker) run you need:

| Dependency  | Notes                                                             |
| ----------- | ----------------------------------------------------------------- |
| Go          | 1.25+ (see `go.mod`)                                              |
| Python      | 3.x with `python-docx==1.2.0` (`pip install -r requirements.txt`) |
| LibreOffice | `soffice` binary, used for DOCX → PDF                             |

---

## Quick start (Docker)

```bash
docker compose up --build
```

The service listens on **http://localhost:8181** (container port 8080). The compose file mounts `./template` and `./output` into the container, so you can drop new templates in without rebuilding.

The image installs the fonts from `./fonts/` into the container font cache — this matters because LibreOffice needs the same fonts as your templates to produce faithful PDFs.

---

## Quick start (local)

Install the Python dependency and point the service at your `soffice` and Python executables:

```bash
pip install -r requirements.txt
```

Linux / macOS:

```bash
SOFFICE_PATH=/usr/bin/soffice PYTHON_CMD=python3 go run .
```

Windows (PowerShell):

```powershell
$env:SOFFICE_PATH = "C:/Program Files/LibreOffice/program/soffice.exe"; $env:PYTHON_CMD = "python"; go run .
```

The server listens on **:8080** and creates `./output` on startup.

### Environment variables

| Variable       | Description                                                                  |
| -------------- | ---------------------------------------------------------------------------- |
| `SOFFICE_PATH` | Full path to the LibreOffice `soffice` executable                            |
| `PYTHON_CMD`   | Python interpreter used to run `docx_processor.py` (e.g. `python3`)          |
| `TZ`           | Timezone, used for the ZIP filename timestamp (compose sets `Asia/Makassar`) |

---

## API

### `POST /documents/generate`

Accepts a **JSON array** of job objects and responds with `application/zip` containing every generated `.docx` and `.pdf`.

CORS is open (`Access-Control-Allow-Origin: *`) and `OPTIONS` preflight is handled.

#### Job object

| Field            | Type   | Description                                                                |
| ---------------- | ------ | -------------------------------------------------------------------------- |
| `filename`       | string | Output base name. Sanitized — anything outside `a-zA-Z0-9._-` becomes `_`. |
| `template`       | string | Template file name inside `./template`. Must end in `.docx`.               |
| `replacements`   | object | Map of placeholder → replacement text.                                     |
| `includedBlocks` | array  | List of `[startMarker, endMarker, included]` triples. Optional.            |
| `datatables`     | object | Map of table marker → column marker → array of cell values. Optional.      |

#### Example request

```bash
curl -X POST http://localhost:8181/documents/generate \
  -H "Content-Type: application/json" \
  -d @data.example.json \
  --output result.zip
```

#### Example payload

```json
[
  {
    "filename": "B-001",
    "template": "surat_tugas_dalam_kota_template.docx",
    "replacements": {
      "{{NoSurat}}": "B-0001/64713/VS.330/2025",
      "{{TanggalSurat}}": "3 Februari 2026",
      "{{Nama}}": "Rizqi Aul",
      "{{Tujuan}}": "Balikpapan",
      "{{TanggalTugas}}": "3 Februari - 10 Maret 2026"
    },
    "includedBlocks": [
      ["[[VISUM_KUNJUNGAN_START]]", "[[VISUM_KUNJUNGAN_END]]", true],
      ["[[LAPORAN_TRANSPORT_START]]", "[[LAPORAN_TRANSPORT_END]]", false],
      ["[[SURAT_PERNYATAAN_START]]", "[[SURAT_PERNYATAAN_END]]", true]
    ]
  }
]
```

See [data.example.json](data.example.json) for a complete multi-document example.

#### Response

A ZIP attachment named `Surat Tugas [YYYY-MM-DD HH_MM].zip`, containing one `.docx` and one `.pdf` per job.

Errors return plain text with `400` (invalid JSON body), `405` (non-POST), or `500` (template, Python, or LibreOffice failure).

---

## Writing templates

Templates live in `./template/`. Everything is driven by plain text markers you type directly into the Word document.

### 1. Placeholders

Write `{{Name}}` anywhere in the document — body, tables, headers, or footers — and map it in `replacements`:

```json
"replacements": { "{{Nama}}": "Nahar Ridlo" }
```

The processor merges the paragraph's runs before matching, so a placeholder that Word has split across several runs (a common problem when editing templates) still resolves. Formatting from the first character of the placeholder is applied to the whole replacement value.

### 2. Conditional blocks

Wrap an optional section between two marker paragraphs, each on its own line:

```
[[VISUM_KUNJUNGAN_START]]
   ... paragraphs and tables to include or drop ...
[[VISUM_KUNJUNGAN_END]]
```

Then control it per job:

```json
"includedBlocks": [
  ["[[VISUM_KUNJUNGAN_START]]", "[[VISUM_KUNJUNGAN_END]]", true]
]
```

- `true` — keep the block; the marker paragraphs themselves are blanked out.
- `false` — remove the block, its markers, and any tables inside it.

### 3. Dynamic tables

Add a single template row to a Word table containing a table marker plus one column marker per column:

| {{T1}}{{No}} | {{Kegiatan}} | {{Volume}} |
| ------------ | ------------ | ---------- |

Supply the data column-wise:

```json
"datatables": {
  "{{T1}}": {
    "{{No}}": ["1", "2", "3"],
    "{{Kegiatan}}": ["Susenas 2026", "Seruti TW I 2026", "HK Februari 2026"],
    "{{Volume}}": ["20", "10", "30"]
  }
}
```

The template row is duplicated once per data row (the row count is the longest column), filled in, and then removed.

### Footers

`fix_footer_first_page_only` runs on every document: it enables the different-first-page header/footer flag and clears footers for all sections after the first, so a multi-section letter keeps its footer only on the opening page.

---

## Caching and output

Generated files land in `./output`:

```
output/
  B-001.docx
  B-001.pdf
  B-001-docx.hash
  B-001-pdf.hash
```

The DOCX hash is the SHA-256 of the canonical JSON of the whole job; the PDF hash is derived from it. On each request the service compares the incoming hash to the stored one and regenerates only when they differ.

Because the cache key is the job payload but the file name is `filename`, **two different jobs sharing a `filename` will overwrite each other**. Keep `filename` unique per batch.

Nothing prunes `./output` automatically — clean it up yourself if disk usage matters.

---

## Project layout

```
.
├── main.go              # HTTP server, worker pipeline, ZIP + PDF, hash cache
├── docx_processor.py    # python-docx template filling (invoked as a subprocess)
├── template/            # .docx templates
├── fonts/               # fonts baked into the Docker image for LibreOffice
├── output/              # generated files + hash markers (created at runtime)
├── data.example.json    # sample request payload
├── Dockerfile
└── docker-compose.yml
```

### How a request flows

```
payloads ──► docxJobs ──► [4x docxWorker] ──► docxOut ──┬──► allOut (docx paths)
                              (python)                   │
                                                         └──► [pdfWorker] ──► allOut (pdf paths)
                                                              (soffice)                │
                                                                                       ▼
                                                                                 ZIP response
```

DOCX generation is parallel across 4 workers; PDF conversion runs on a single worker, since LibreOffice does not handle concurrent headless invocations well.

---

## Notes

- `.gitignore` excludes `data.example.json` from version control — it is present locally as a reference payload.
- There is no authentication on the endpoint. Do not expose it directly to the internet without putting a proxy and access control in front of it.
