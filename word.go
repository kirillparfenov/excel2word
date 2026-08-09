package main

import (
	"fmt"
	"path/filepath"

	"github.com/lukasjarosch/go-docx"
)

func replaceWordReplacements(dir string, excelReplacements []docx.PlaceholderMap) ([]string, error) {
	wordPath, err := findFileByExt(dir, ".docx")
	if err != nil {
		return nil, err
	}
	fmt.Println("Word-шаблон:", filepath.Base(wordPath))

	outDir, err := outputDir(dir)
	if err != nil {
		return nil, err
	}

	// логируем пропуски реплейсментов в excel -> word
	doc, err := docx.Open(wordPath)
	if err != nil {
		return nil, fmt.Errorf("ошибка при открытии word-файла: %w", err)
	}
	defer doc.Close()
	logMissPlaceholders(doc, excelReplacements[0])

	createdDocs := make([]string, 0, len(excelReplacements)+1)
	for i, replacements := range excelReplacements {
		outName := fmt.Sprintf("document_%03d.docx", i+1)
		outPath := filepath.Join(outDir, outName)

		if err := generateDocx(wordPath, outPath, replacements); err != nil {
			return nil, fmt.Errorf("ошибка в процессе генерации .docx: %w", err)
		}
		createdDocs = append(createdDocs, outName)
	}

	return createdDocs, nil
}

// generateDocx открывает шаблон заново (т.к. замена мутирует документ в памяти),
// подставляет значения вместо {ключ} и сохраняет результат.
func generateDocx(templatePath, outPath string, replacements docx.PlaceholderMap) error {
	doc, err := docx.Open(templatePath)
	if err != nil {
		return fmt.Errorf("не удалось открыть шаблон: %w", err)
	}
	defer doc.Close()

	if err := doc.ReplaceAll(replacements); err != nil {
		return fmt.Errorf("не удалось выполнить замену: %w", err)
	}

	if err := doc.WriteToFile(outPath); err != nil {
		return fmt.Errorf("не удалось сохранить файл: %w", err)
	}
	return nil
}
