# Excel/Office → PDF API (Go + LibreOffice)

This repository exposes a minimal Go HTTP API that receives Excel/Office files and returns PDFs generated through LibreOffice (headless).  
It is the backend service that now powers the `pdf-to-excel-api` project and replaces the earlier Node implementation.  
Originally based on [wteja/pdf-converter](https://github.com/wteja/pdf-converter) – all credits to the original author – but heavily adapted (Swagger docs, single-page sheets, container build, etc.).

## Features

- Converts various office files to PDF format using LibreOffice.
- Handles file uploads via HTTP POST requests (multipart form data).
- Automatic cleanup of old files from the temporary directory after one hour.
- Minimal and efficient implementation using Go + LibreOffice headless.
- **Each spreadsheet sheet renders as a single PDF page** thanks to the `SinglePageSheets` filter.
- **Swagger/OpenAPI documentation** – interactive UI at `/docs` + raw spec at `/api/openapi.json`.

## Requirements

- **Go**: Ensure Go is installed on your system ([Download Go](https://golang.org/dl/)).
- **LibreOffice**: LibreOffice must be installed and accessible via the `soffice` command.

## Supported File Formats

This application relies on LibreOffice for file conversion, so it supports any file format LibreOffice can handle. Below are the formats that can be converted to PDF:

### Document Formats

- `.doc`, `.docx` (Microsoft Word)
- `.odt`, `.ott` (LibreOffice/OpenDocument Text)
- `.rtf` (Rich Text Format)
- `.txt` (Plain Text)

### Spreadsheet Formats

- `.xls`, `.xlsx` (Microsoft Excel)
- `.ods`, `.ots` (LibreOffice/OpenDocument Spreadsheet)
- `.csv` (Comma-Separated Values)

### Presentation Formats

- `.ppt`, `.pptx` (Microsoft PowerPoint)
- `.odp`, `.otp` (LibreOffice/OpenDocument Presentation)

### Other Formats

- `.svg` (Scalable Vector Graphics)
- `.html`, `.htm` (HTML Files)
- `.xml` (XML Files)
- `.pdf` (for PDF editing and re-exporting)

## Installation

> The service is typically consumed via Docker (see below).  
> If you prefer building from source:

1. Clone the repository:

   ```bash
   git clone git@github.com:iams18i/pdf-to-excel-api.git
   cd pdf-to-excel-api
   ```

2. Install dependencies:
   This application does not have external dependencies, but ensure you have LibreOffice installed.

3. Build the application:

   ```bash
   go build -o pdf-converter .
   ```

4. Run the application:
   ```bash
   ./pdf-converter
   ```

The server will start listening on `http://localhost:5000`.

## Quick start (Docker Compose)

```bash
docker compose up -d --build
```

This builds the Go binary, installs LibreOffice inside the runtime image, and exposes port `5000`.

---

## API Usage

### **API Documentation**

- **Swagger UI**: `http://localhost:5000/docs` - Interactive API documentation
- **OpenAPI Spec**: `http://localhost:5000/api/openapi.json` - OpenAPI 3.0 specification

### **Endpoints**

#### **Health Check**

- **Endpoint**: `GET /` or `GET /health`
- **Response**: JSON with service status, timestamp, and version

#### **Convert File to PDF**

- **Endpoint**: `POST /convert`
- **Method**: `POST`
- **Content-Type**: `multipart/form-data`
- **Field Name**: `file`
- **File Type**: Any supported format (e.g., `.xlsx`, `.docx`).

#### Request Example (Using `curl`):

```bash
curl -X POST -F "file=@example.xlsx" http://localhost:5000/convert --output output.pdf
```

#### Response:

- **Success (200)**: Returns the converted PDF file as a response with the `Content-Type` set to `application/pdf`.
- **Error (400)**: Bad request - invalid file or missing file
- **Error (405)**: Method not allowed
- **Error (500)**: Internal server error - conversion failed

## Public Docker Image

While this repository focuses on the self-hosted `pdf-to-excel-api`, the image is still compatible with the public `wteja/pdf-converter` image:

```bash
docker pull wteja/pdf-converter
```

Run the container:

```bash
docker run -p 5000:5000 wteja/pdf-converter
```

## Configuration

- Temporary files are stored in the `./tmp` directory. Ensure the application has write access to this directory.
- The application automatically removes files older than one hour from the `tmp` directory.

## Code Overview

### **Main Components**

1. **File Upload and Conversion**:

   - The `/convert` endpoint processes file uploads, saves them to the `tmp` directory, and invokes LibreOffice in headless mode to perform the conversion.

2. **Temporary Directory Management**:

   - All uploaded and converted files are stored in the `tmp` directory.
   - A background goroutine periodically checks and deletes files older than one hour.

3. **Error Handling**:
   - Comprehensive error handling for file uploads, conversions, and temporary file management.

### **Key Functions**

- `handleConvert`: Handles the HTTP requests, manages file upload and conversion, and returns the resulting PDF.
- `cleanupOldFiles`: Periodically deletes old files from the `tmp` directory.

## Future Improvements

- Support additional file formats (e.g., `.pptx`, `.odg`).
- Add configurable cleanup duration and temporary directory path.
- Implement better logging and monitoring.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Author

Adapted for `pdf-to-excel-api` by [iams18i](https://github.com/iams18i).  
Original implementation by [Weerayut Teja](https://github.com/wteja) – thank you!
