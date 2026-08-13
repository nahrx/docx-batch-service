# --- Stage 1: Build Go Application ---
FROM golang:1.25 AS builder
WORKDIR /app
COPY go.mod go.sum ./
RUN go mod download
COPY . .
RUN CGO_ENABLED=0 GOOS=linux go build -o main .

# --- Stage 2: Final Runtime Environment ---
FROM debian:bookworm-slim

# Install LibreOffice and Font tools
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# 1. Create a custom font directory
RUN mkdir -p /usr/share/fonts/truetype/custom

# 2. Copy your local fonts into the container
COPY ./fonts/* /usr/share/fonts/truetype/custom/

# 3. Refresh the system font cache so LibreOffice sees them
RUN fc-cache -f -v

# Install Python 3, Pip, and LibreOffice
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    libreoffice-writer \
    libreoffice-java-common \
    && rm -rf /var/lib/apt/lists/*

# Install required Python package
RUN pip3 install --break-system-packages python-docx

WORKDIR /app
# Copy the Go binary and Python script from builder
COPY --from=builder /app/main .
COPY --from=builder /app/docx_processor.py .

# Create template directory
# RUN mkdir -p /app/template
# COPY --from=builder /app/template .

# Create output directory for generated files
RUN mkdir -p /app/output


# Set Linux paths for Go application to use
ENV SOFFICE_PATH=/usr/bin/soffice
ENV PYTHON_CMD=python3

EXPOSE 8080
CMD ["./main"]
