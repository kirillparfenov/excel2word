package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// findFileByExt ищет ровно один файл с заданным расширением в папке dir.
// Файлы вида ~$*.docx / ~$*.xlsx (временные файлы Word/Excel) и то, что
// программа сама создаёт в output/, игнорируются.
func findFileByExt(dir, ext string) (string, error) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return "", err
	}

	var found []string
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		name := e.Name()
		if strings.HasPrefix(name, "~$") {
			continue
		}
		if strings.EqualFold(filepath.Ext(name), ext) {
			found = append(found, filepath.Join(dir, name))
		}
	}

	switch len(found) {
	case 0:
		return "", fmt.Errorf("в папке не найден файл %s", ext)
	case 1:
		return found[0], nil
	default:
		return "", fmt.Errorf("в папке найдено несколько файлов %s, оставьте только один: %v", ext, found)
	}
}

// outputDir исходная директория для сохранения новых word-файлов
func outputDir(dir string) (outDir string, err error) {
	outDir = filepath.Join(dir, "output")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return "", fmt.Errorf("не удалось создать папку output: %w", err)
	}
	return
}
