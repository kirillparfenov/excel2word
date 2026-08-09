package main

import (
	"fmt"
	"path/filepath"
	"sync"

	"github.com/lukasjarosch/go-docx"
)

// Document
// outDir - директория сохранения файла
// rowIndex - индекс реплейсмента
// wordPath - путь до word файла-шаблона
// replacements - реплейсменты для word файла
type Document struct {
	outDir       string
	rowIndex     int
	wordPath     string
	replacements docx.PlaceholderMap
}

func replaceWordReplacements(dir string, excelReplacements []docx.PlaceholderMap) ([]string, error) {
	wordPath, err := findFileByExt(dir, ".docx")
	if err != nil {
		return nil, err
	}
	fmt.Println("Word-шаблон:", filepath.Base(wordPath))

	// логируем пропуски реплейсментов в excel -> word
	doc, err := docx.Open(wordPath)
	if err != nil {
		return nil, fmt.Errorf("ошибка при открытии word-файла: %w", err)
	}
	defer doc.Close()
	logMissPlaceholders(doc, excelReplacements[0])

	outDir, err := outputDir(dir)
	if err != nil {
		return nil, err
	}

	createdDocs := make([]string, len(excelReplacements))
	errors := make([]error, len(excelReplacements))
	var wg sync.WaitGroup
	wg.Add(len(excelReplacements))

	for i, replacements := range excelReplacements {
		go func(i int, replacements docx.PlaceholderMap) {
			defer wg.Done()
			docName, err := createDoc(&Document{
				outDir:       outDir,
				rowIndex:     i,
				wordPath:     wordPath,
				replacements: replacements,
			})
			if err != nil {
				errors[i] = err
				return
			}
			createdDocs[i] = docName
		}(i, replacements)
	}
	wg.Wait()

	for _, err := range errors {
		if err != nil {
			return nil, err
		}
	}

	return createdDocs, nil
}

// createDoc создает документ
// открывает шаблон заново (т.к. замена мутирует документ в памяти),
// подставляет значения вместо {ключ} и сохраняет результат.
func createDoc(doc *Document) (string, error) {
	word, err := docx.Open(doc.wordPath)
	if err != nil {
		return "", fmt.Errorf("не удалось открыть шаблон word: %w", err)
	}
	defer word.Close()

	if err := word.ReplaceAll(doc.replacements); err != nil {
		return "", fmt.Errorf("не удалось выполнить замену: %w", err)
	}

	outName := fmt.Sprintf("document_%03d.docx", doc.rowIndex+1)
	outPath := filepath.Join(doc.outDir, outName)
	if err := word.WriteToFile(outPath); err != nil {
		return "", fmt.Errorf("не удалось сохранить файл: %w", err)
	}

	return outName, nil
}
