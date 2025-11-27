package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"

	"github.com/go-pdf/fpdf"
	"github.com/pdfcpu/pdfcpu/pkg/api"
	"github.com/phpdave11/gofpdi"
)

const tempDir = "./tmp" // Directory for temporary files

func main() {
	// Ensure the temporary directory exists
	if err := os.MkdirAll(tempDir, os.ModePerm); err != nil {
		fmt.Println("Failed to create temp directory:", err)
		return
	}

	// Start the file cleanup goroutine
	go cleanupOldFiles(tempDir, 1*time.Hour)

	apiToken := os.Getenv("API_TOKEN")
	if apiToken == "" {
		log.Fatal("API_TOKEN environment variable is required")
	}

	http.HandleFunc("/", handleHealthCheck)
	http.HandleFunc("/health", handleHealthCheck)
	http.HandleFunc("/docs", handleSwaggerUI)
	http.HandleFunc("/api/openapi.json", handleOpenAPISpec)
	http.HandleFunc("/convert", authMiddleware(apiToken, handleConvert))

	fmt.Println("Starting server on :5000")
	if err := http.ListenAndServe(":5000", nil); err != nil {
		fmt.Println("Failed to start server:", err)
	}
}

// @title PDF Converter API
// @version 1.0.0
// @description API for converting Excel files to PDF using LibreOffice
// @host localhost:5000
// @BasePath /
func handleHealthCheck(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{
		"status":    "ok",
		"timestamp": time.Now().Format(time.RFC3339),
		"service":   "PDF Converter",
		"version":   "1.0.0",
	})
}

// handleSwaggerUI serves the Swagger UI
func handleSwaggerUI(w http.ResponseWriter, r *http.Request) {
	swaggerHTML := `<!DOCTYPE html>
<html>
<head>
    <title>PDF Converter API - Swagger UI</title>
    <link rel="stylesheet" type="text/css" href="https://unpkg.com/swagger-ui-dist@5.10.0/swagger-ui.css" />
    <style>
        html { box-sizing: border-box; overflow: -moz-scrollbars-vertical; overflow-y: scroll; }
        *, *:before, *:after { box-sizing: inherit; }
        body { margin:0; background: #fafafa; }
    </style>
</head>
<body>
    <div id="swagger-ui"></div>
    <script src="https://unpkg.com/swagger-ui-dist@5.10.0/swagger-ui-bundle.js"></script>
    <script src="https://unpkg.com/swagger-ui-dist@5.10.0/swagger-ui-standalone-preset.js"></script>
    <script>
        window.onload = function() {
            const ui = SwaggerUIBundle({
                url: "/api/openapi.json",
                dom_id: '#swagger-ui',
                deepLinking: true,
                presets: [
                    SwaggerUIBundle.presets.apis,
                    SwaggerUIStandalonePreset
                ],
                plugins: [
                    SwaggerUIBundle.plugins.DownloadUrl
                ],
                layout: "StandaloneLayout"
            });
        };
    </script>
</body>
</html>`
	w.Header().Set("Content-Type", "text/html")
	w.Write([]byte(swaggerHTML))
}

// handleOpenAPISpec returns the OpenAPI specification
func handleOpenAPISpec(w http.ResponseWriter, r *http.Request) {
	spec := map[string]interface{}{
		"openapi": "3.0.0",
		"info": map[string]interface{}{
			"title":       "PDF Converter API",
			"version":     "1.0.0",
			"description": "API for converting Excel files (.xlsx, .xls) to PDF documents using LibreOffice",
		},
		"servers": []map[string]interface{}{
			{
				"url":         "http://localhost:5000",
				"description": "Development server",
			},
		},
		"components": map[string]interface{}{
			"securitySchemes": map[string]interface{}{
				"ApiTokenAuth": map[string]interface{}{
					"type": "apiKey",
					"in":   "header",
					"name": "x-auth-token",
				},
			},
		},
		"security": []map[string]interface{}{
			{"ApiTokenAuth": []interface{}{}},
		},
		"paths": map[string]interface{}{
			"/": map[string]interface{}{
				"get": map[string]interface{}{
					"summary":     "Health check",
					"description": "Returns the health status of the API",
					"operationId": "healthCheck",
					"responses": map[string]interface{}{
						"200": map[string]interface{}{
							"description": "Service is healthy",
							"content": map[string]interface{}{
								"application/json": map[string]interface{}{
									"schema": map[string]interface{}{
										"type": "object",
										"properties": map[string]interface{}{
											"status": map[string]interface{}{
												"type":        "string",
												"example":     "ok",
												"description": "Service status",
											},
											"timestamp": map[string]interface{}{
												"type":        "string",
												"format":      "date-time",
												"description": "Current timestamp",
											},
											"service": map[string]interface{}{
												"type":        "string",
												"example":     "PDF Converter",
												"description": "Service name",
											},
											"version": map[string]interface{}{
												"type":        "string",
												"example":     "1.0.0",
												"description": "API version",
											},
										},
									},
								},
							},
						},
					},
				},
			},
			"/health": map[string]interface{}{
				"get": map[string]interface{}{
					"summary":     "Health check endpoint",
					"description": "Returns the health status of the API",
					"operationId": "health",
					"responses": map[string]interface{}{
						"200": map[string]interface{}{
							"description": "Service is healthy",
							"content": map[string]interface{}{
								"application/json": map[string]interface{}{
									"schema": map[string]interface{}{
										"type": "object",
										"properties": map[string]interface{}{
											"status": map[string]interface{}{
												"type":    "string",
												"example": "ok",
											},
										},
									},
								},
							},
						},
					},
				},
			},
			"/convert": map[string]interface{}{
				"post": map[string]interface{}{
					"summary":     "Convert Excel to PDF",
					"description": "Upload an Excel file (.xlsx or .xls) and convert it to PDF using LibreOffice. Each sheet will be rendered in the PDF.",
					"operationId": "convertExcelToPdf",
					"security": []map[string]interface{}{
						{"ApiTokenAuth": []interface{}{}},
					},
					"requestBody": map[string]interface{}{
						"required": true,
						"content": map[string]interface{}{
							"multipart/form-data": map[string]interface{}{
								"schema": map[string]interface{}{
									"type": "object",
									"required": []string{"file"},
									"properties": map[string]interface{}{
										"file": map[string]interface{}{
											"type":        "string",
											"format":      "binary",
											"description": "Excel file (.xlsx or .xls)",
										},
									},
								},
							},
						},
					},
					"responses": map[string]interface{}{
						"200": map[string]interface{}{
							"description": "PDF file generated successfully",
							"content": map[string]interface{}{
								"application/pdf": map[string]interface{}{
									"schema": map[string]interface{}{
										"type":   "string",
										"format": "binary",
									},
								},
							},
							"headers": map[string]interface{}{
								"Content-Disposition": map[string]interface{}{
									"schema": map[string]interface{}{
										"type":        "string",
										"description": "Attachment filename",
										"example":     "attachment; filename=\"output.pdf\"",
									},
								},
							},
						},
						"400": map[string]interface{}{
							"description": "Bad request - invalid file or missing file",
							"content": map[string]interface{}{
								"text/plain": map[string]interface{}{
									"schema": map[string]interface{}{
										"type": "string",
									},
								},
							},
						},
						"405": map[string]interface{}{
							"description": "Method not allowed",
							"content": map[string]interface{}{
								"text/plain": map[string]interface{}{
									"schema": map[string]interface{}{
										"type": "string",
									},
								},
							},
						},
						"500": map[string]interface{}{
							"description": "Internal server error - conversion failed",
							"content": map[string]interface{}{
								"text/plain": map[string]interface{}{
									"schema": map[string]interface{}{
										"type": "string",
									},
								},
							},
						},
					},
				},
			},
		},
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(spec)
}

func handleConvert(w http.ResponseWriter, r *http.Request) {
	// Ensure the request method is POST
	if r.Method != http.MethodPost {
		http.Error(w, "Only POST method is allowed", http.StatusMethodNotAllowed)
		return
	}

	// Parse the uploaded file
	file, fileHeader, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "Failed to read uploaded file", http.StatusBadRequest)
		return
	}
	defer file.Close()

	// Detect file extension from uploaded filename
	originalFileName := fileHeader.Filename
	fileExt := filepath.Ext(originalFileName)
	if fileExt == "" {
		fileExt = ".xlsx" // Default to xlsx if no extension
	}

	// Save the Excel file to a temporary location
	baseName := time.Now().Format("20060102150405") // Timestamp format
	inputFilePath := filepath.Join(tempDir, baseName+fileExt)

	inputFile, err := os.Create(inputFilePath)
	if err != nil {
		http.Error(w, "Failed to create temporary file", http.StatusInternalServerError)
		return
	}

	_, err = io.Copy(inputFile, file)
	if err != nil {
		inputFile.Close()
		os.Remove(inputFilePath)
		http.Error(w, "Failed to save uploaded file", http.StatusInternalServerError)
		return
	}

	// Close and flush the file before conversion
	inputFile.Close()

	// Get absolute paths (LibreOffice works better with absolute paths)
	absInputPath, err := filepath.Abs(inputFilePath)
	if err != nil {
		http.Error(w, "Failed to get absolute path", http.StatusInternalServerError)
		return
	}
	absTempDir, err := filepath.Abs(tempDir)
	if err != nil {
		http.Error(w, "Failed to get absolute temp directory", http.StatusInternalServerError)
		return
	}

	// Convert the Excel file to PDF using LibreOffice
	// Use calc_pdf_Export filter with SinglePageSheets option to fit each sheet on one page
	// Add 50px (~13.2mm) padding on every side via margin properties (values in 1/100 mm)
	// Filter format: pdf:calc_pdf_Export:{JSON filter data}
	filterData := `pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":true},"LeftMargin":{"type":"long","value":1320},"RightMargin":{"type":"long","value":1320},"TopMargin":{"type":"long","value":1320},"BottomMargin":{"type":"long","value":1320}}`
	
	var stdout, stderr bytes.Buffer
	cmd := exec.Command("soffice", "--headless", "--nodefault", "--nolockcheck", "--convert-to", filterData, absInputPath, "--outdir", absTempDir)
	cmd.Env = os.Environ()
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	
	fmt.Printf("Running LibreOffice conversion with SinglePageSheets: soffice --headless --nodefault --nolockcheck --convert-to '%s' %s --outdir %s\n", filterData, absInputPath, absTempDir)
	
	convErr := cmd.Run()
	if convErr != nil {
		fmt.Printf("LibreOffice conversion error with SinglePageSheets: %v\n", convErr)
		fmt.Printf("stdout: %s\n", stdout.String())
		fmt.Printf("stderr: %s\n", stderr.String())
		
		// Fallback: Try without filter options (will have page breaks but at least works)
		fmt.Printf("Trying fallback conversion without filter options...\n")
		stdout.Reset()
		stderr.Reset()
		
		cmdFallback := exec.Command("soffice", "--headless", "--nodefault", "--nolockcheck", "--convert-to", "pdf", absInputPath, "--outdir", absTempDir)
		cmdFallback.Env = os.Environ()
		cmdFallback.Stdout = &stdout
		cmdFallback.Stderr = &stderr
		
		convErr = cmdFallback.Run()
		if convErr != nil {
			fmt.Printf("Fallback conversion error: %v\n", convErr)
			fmt.Printf("stdout: %s\n", stdout.String())
			fmt.Printf("stderr: %s\n", stderr.String())
			http.Error(w, fmt.Sprintf("Failed to convert file to PDF: %v. stderr: %s", convErr, stderr.String()), http.StatusInternalServerError)
			return
		}
		fmt.Printf("Fallback conversion succeeded (may have page breaks)\n")
	}
	
	fmt.Printf("LibreOffice stdout: %s\n", stdout.String())
	if stderr.Len() > 0 {
		fmt.Printf("LibreOffice stderr: %s\n", stderr.String())
	}
	
	// Wait a moment for file system to sync
	time.Sleep(100 * time.Millisecond)
	
	// LibreOffice creates PDF with the same base name as input file
	// So if input is "20251127002624.xlsx", output will be "20251127002624.pdf"
	inputBaseName := filepath.Base(absInputPath)
	inputBaseNameWithoutExt := inputBaseName[:len(inputBaseName)-len(fileExt)]
	expectedPdfName := inputBaseNameWithoutExt + ".pdf"
	pdfPath := filepath.Join(absTempDir, expectedPdfName)
	
	// Verify the output file was created
	if _, err := os.Stat(pdfPath); os.IsNotExist(err) {
		// Search for any PDF file in temp directory
		files, readErr := os.ReadDir(absTempDir)
		if readErr != nil {
			fmt.Printf("Failed to read temp directory: %v\n", readErr)
		}
		
		found := false
		for _, f := range files {
			if !f.IsDir() && filepath.Ext(f.Name()) == ".pdf" {
				pdfPath = filepath.Join(absTempDir, f.Name())
				fmt.Printf("Found PDF file: %s\n", pdfPath)
				found = true
				break
			}
		}
		
		if !found {
			fmt.Printf("PDF file was not created. Expected: %s\n", pdfPath)
			fmt.Printf("Files in temp directory:\n")
			for _, f := range files {
				fmt.Printf("  - %s (dir: %v)\n", f.Name(), f.IsDir())
			}
			http.Error(w, "PDF conversion completed but file was not found", http.StatusInternalServerError)
			return
		}
	} else {
		fmt.Printf("PDF file found at: %s\n", pdfPath)
	}

	// Add padding around every page (~50px ≈ 13.2mm)
	const marginMM = 13.2
	paddedPath, err := addPaddingToPDF(pdfPath, marginMM)
	if err != nil {
		fmt.Printf("Failed to add padding to PDF: %v\n", err)
		paddedPath = pdfPath
	} else {
		defer os.Remove(paddedPath)
		os.Remove(pdfPath)
		pdfPath = paddedPath
	}

	// Read the converted PDF file
	pdfFile, err := os.Open(pdfPath)
	if err != nil {
		fmt.Println(err)
		http.Error(w, "Failed to read converted PDF", http.StatusInternalServerError)
		return
	}
	defer pdfFile.Close()

	// Write the PDF file as response
	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", `attachment; filename="output.pdf"`)

	if _, err := io.Copy(w, pdfFile); err != nil {
		http.Error(w, "Failed to write PDF to response", http.StatusInternalServerError)
		return
	}
}

// cleanupOldFiles removes files older than the specified duration from the given directory
func cleanupOldFiles(dir string, maxAge time.Duration) {
	for {
		time.Sleep(1 * time.Hour) // Check every minute

		files, err := os.ReadDir(dir)
		if err != nil {
			fmt.Println("Failed to read temp directory:", err)
			continue
		}

		for _, file := range files {
			filePath := filepath.Join(dir, file.Name())
			info, err := os.Stat(filePath)
			if err != nil {
				fmt.Println("Failed to get file info:", err)
				continue
			}

			// Check if the file is older than maxAge
			if time.Since(info.ModTime()) > maxAge {
				if err := os.Remove(filePath); err != nil {
					fmt.Println("Failed to delete file:", err)
				} else {
					fmt.Println("Deleted old file:", filePath)
				}
			}
		}
	}
}

func authMiddleware(expectedToken string, next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		token := r.Header.Get("x-auth-token")
		if token == "" || token != expectedToken {
			http.Error(w, "Unauthorized", http.StatusUnauthorized)
			return
		}
		next.ServeHTTP(w, r)
	}
}

func addPaddingToPDF(inputPath string, marginMM float64) (string, error) {
	pageCount, err := api.PageCountFile(inputPath)
	if err != nil {
		return "", fmt.Errorf("count pages: %w", err)
	}

	if pageCount == 0 {
		return "", fmt.Errorf("pdf has no pages")
	}

	outputPath := strings.TrimSuffix(inputPath, ".pdf") + "_padded.pdf"

	pdf := fpdf.New("P", "mm", "", "")
	for page := 1; page <= pageCount; page++ {
		tpl := gofpdi.ImportPage(pdf, inputPath, page, "/MediaBox")
		width, height := gofpdi.GetTemplateSize(pdf, tpl)
		pdf.AddPageFormat("P", fpdf.SizeType{Wd: width + marginMM*2, Ht: height + marginMM*2})
		gofpdi.UseImportedTemplate(pdf, tpl, marginMM, marginMM, width, height)
	}

	if err := pdf.OutputFileAndClose(outputPath); err != nil {
		return "", fmt.Errorf("write padded pdf: %w", err)
	}

	return outputPath, nil
}
