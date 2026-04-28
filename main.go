package main

import (
	"embed"
	"encoding/json"
	"errors"
	"fmt"
	"io/fs"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path"
	"runtime"
	"strings"
	"time"

	"supply-tablefusion/internal/excel"
)

//go:embed web/dist
var staticFiles embed.FS

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
	file, header, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "missing file", http.StatusBadRequest)
		return
	}
	defer file.Close()

	output, err := excel.Transform(file, sourceType)
	if err != nil {
		switch {
		case errors.Is(err, excel.ErrUnsupportedSourceType):
			http.Error(w, err.Error(), http.StatusBadRequest)
		case errors.Is(err, excel.ErrMappingsNotImplemented):
			http.Error(w, err.Error(), http.StatusNotImplemented)
		default:
			http.Error(w, fmt.Sprintf("transform workbook: %v", err), http.StatusUnprocessableEntity)
		}
		return
	}

	filename := "standard-" + strings.TrimSuffix(header.Filename, path.Ext(header.Filename)) + ".xlsx"
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", fmt.Sprintf(`attachment; filename="%s"`, filename))
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(output.Bytes())
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
