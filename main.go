package main

import (
	"archive/zip"
	"bytes"
	"embed"
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"log"
	"mime/multipart"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"supply-tablefusion/internal/excel"
)

//go:embed web/dist
var staticFiles embed.FS

//go:embed 示例文件/硬件产品信息.xlsx
var hardwareProductMapping []byte

const maxUploadSize = 32 << 20

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	mux, err := newMux()
	if err != nil {
		log.Fatalf("initialize server: %v", err)
	}

	addr := "127.0.0.1:" + port
	url := "http://" + addr
	go openBrowser(url)

	server := &http.Server{
		Addr:              addr,
		Handler:           mux,
		ReadHeaderTimeout: 5 * time.Second,
	}

	log.Printf("listening on %s", url)
	if err := server.ListenAndServe(); err != nil && !errors.Is(err, http.ErrServerClosed) {
		log.Fatalf("server stopped: %v", err)
	}
}

func newMux() (*http.ServeMux, error) {
	dist, err := fs.Sub(staticFiles, "web/dist")
	if err != nil {
		return nil, err
	}

	mux := http.NewServeMux()
	mux.HandleFunc("/api/health", handleHealth)
	mux.HandleFunc("/api/transform", handleTransform)
	mux.Handle("/", spaHandler(dist))
	return mux, nil
}

func handleHealth(w http.ResponseWriter, _ *http.Request) {
	writeJSON(w, http.StatusOK, map[string]string{"status": "ok"})
}

func handleTransform(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, maxUploadSize)
	if err := r.ParseMultipartForm(maxUploadSize); err != nil {
		http.Error(w, "invalid multipart request", http.StatusBadRequest)
		return
	}

	sourceType := excel.SourceType(r.FormValue("sourceType"))
	files := uploadedFiles(r)
	if len(files) == 0 {
		http.Error(w, "missing file", http.StatusBadRequest)
		return
	}

	if len(files) == 1 {
		result, err := transformUploadedFile(files[0], sourceType)
		if err != nil {
			writeTransformError(w, err)
			return
		}

		logPath, err := writeTransformLog(files[0].Filename, sourceType, result.LogMarkdown)
		if err != nil {
			http.Error(w, fmt.Sprintf("write transform log: %v", err), http.StatusInternalServerError)
			return
		}

		filename := outputWorkbookFilename(files[0].Filename)
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		w.Header().Set("Content-Disposition", contentDisposition(filename))
		w.Header().Set("X-Transform-Log-Path", logPath)
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write(result.Workbook.Bytes())
		return
	}

	zipData, err := transformUploadedFilesToZip(files, sourceType)
	if err != nil {
		writeTransformError(w, err)
		return
	}

	filename := fmt.Sprintf("处理后文件-%s.zip", time.Now().Format("20060102-150405"))
	w.Header().Set("Content-Type", "application/zip")
	w.Header().Set("Content-Disposition", contentDisposition(filename))
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(zipData.Bytes())
}

func contentDisposition(filename string) string {
	encoded := url.PathEscape(filename)
	return fmt.Sprintf(`attachment; filename="download%s"; filename*=UTF-8''%s`, filepath.Ext(filename), encoded)
}

func uploadedFiles(r *http.Request) []*multipart.FileHeader {
	if r.MultipartForm == nil {
		return nil
	}
	if files := r.MultipartForm.File["files"]; len(files) > 0 {
		return files
	}
	return r.MultipartForm.File["file"]
}

func transformUploadedFile(header *multipart.FileHeader, sourceType excel.SourceType) (*excel.Result, error) {
	file, err := header.Open()
	if err != nil {
		return nil, fmt.Errorf("open uploaded file %s: %w", header.Filename, err)
	}
	defer file.Close()

	return excel.Transform(file, sourceType, bytes.NewReader(hardwareProductMapping))
}

func transformUploadedFilesToZip(files []*multipart.FileHeader, sourceType excel.SourceType) (*bytes.Buffer, error) {
	output := &bytes.Buffer{}
	archive := zip.NewWriter(output)

	usedNames := make(map[string]int, len(files))
	for _, file := range files {
		result, err := transformUploadedFile(file, sourceType)
		if err != nil {
			_ = archive.Close()
			return nil, fmt.Errorf("%s: %w", file.Filename, err)
		}

		filename := uniqueZipFilename(outputWorkbookFilename(file.Filename), usedNames)
		writer, err := archive.Create(filename)
		if err != nil {
			_ = archive.Close()
			return nil, fmt.Errorf("create zip entry %s: %w", filename, err)
		}
		if _, err := writer.Write(result.Workbook.Bytes()); err != nil {
			_ = archive.Close()
			return nil, fmt.Errorf("write zip entry %s: %w", filename, err)
		}
	}

	if err := archive.Close(); err != nil {
		return nil, fmt.Errorf("close zip: %w", err)
	}
	return output, nil
}

func writeTransformError(w http.ResponseWriter, err error) {
	switch {
	case errors.Is(err, excel.ErrUnsupportedSourceType):
		http.Error(w, err.Error(), http.StatusBadRequest)
	case errors.Is(err, excel.ErrMappingsNotImplemented):
		http.Error(w, err.Error(), http.StatusNotImplemented)
	default:
		http.Error(w, fmt.Sprintf("transform workbook: %v", err), http.StatusUnprocessableEntity)
	}
}

func outputWorkbookFilename(sourceFilename string) string {
	base := strings.TrimSuffix(filepath.Base(sourceFilename), filepath.Ext(sourceFilename))
	if base == "." || base == "" {
		base = "workbook"
	}
	return "处理后_" + safeLogFilename(base) + ".xlsx"
}

func uniqueZipFilename(filename string, used map[string]int) string {
	count := used[filename]
	used[filename] = count + 1
	if count == 0 {
		return filename
	}
	ext := filepath.Ext(filename)
	base := strings.TrimSuffix(filename, ext)
	return fmt.Sprintf("%s-%d%s", base, count+1, ext)
}

func writeTransformLog(sourceFilename string, sourceType excel.SourceType, content string) (string, error) {
	if err := os.MkdirAll("logs", 0o755); err != nil {
		return "", err
	}

	baseName := strings.TrimSuffix(filepath.Base(sourceFilename), filepath.Ext(sourceFilename))
	if baseName == "." || baseName == "" {
		baseName = "workbook"
	}
	baseName = safeLogFilename(baseName)

	filename := fmt.Sprintf("%s-%s-%s.md", time.Now().Format("20060102-150405"), sourceType, baseName)
	logPath := filepath.Join("logs", filename)
	if err := os.WriteFile(logPath, []byte(content), 0o644); err != nil {
		return "", err
	}
	return logPath, nil
}

func safeLogFilename(value string) string {
	var builder strings.Builder
	for _, r := range value {
		switch {
		case r >= 'a' && r <= 'z':
			builder.WriteRune(r)
		case r >= 'A' && r <= 'Z':
			builder.WriteRune(r)
		case r >= '0' && r <= '9':
			builder.WriteRune(r)
		case r == '-' || r == '_' || r == '.' || r >= 0x4e00 && r <= 0x9fff:
			builder.WriteRune(r)
		default:
			builder.WriteRune('-')
		}
	}
	return strings.Trim(builder.String(), "-")
}

func spaHandler(dist fs.FS) http.Handler {
	fileServer := http.FileServer(http.FS(dist))

	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		requestPath := strings.TrimPrefix(path.Clean(r.URL.Path), "/")
		if requestPath == "." || requestPath == "" {
			requestPath = "index.html"
		}
		if _, err := fs.Stat(dist, requestPath); err != nil {
			r.URL.Path = "/index.html"
		}
		fileServer.ServeHTTP(w, r)
	})
}

func writeJSON(w http.ResponseWriter, status int, value any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(value)
}

func openBrowser(url string) {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "darwin":
		cmd = exec.Command("open", url)
	case "windows":
		cmd = exec.Command("rundll32", "url.dll,FileProtocolHandler", url)
	default:
		cmd = exec.Command("xdg-open", url)
	}
	if err := cmd.Start(); err != nil {
		log.Printf("open browser manually: %s", url)
	}
}
