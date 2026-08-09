package main

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/lukasjarosch/go-docx"
	"github.com/xuri/excelize/v2"
)

// excelReplacements достаем все колонки со строками из excel в виде структуры шаблонов для word
func excelReplacements(dir string) ([]docx.PlaceholderMap, error) {
	excelPath, err := findFileByExt(dir, ".xlsx")
	if err != nil {
		return nil, err
	}
	fmt.Println("Excel-файл:", filepath.Base(excelPath))

	headers, dataRows, err := readExcel(excelPath)
	if err != nil {
		return nil, fmt.Errorf("не удалось прочитать Excel: %w", err)
	}
	if len(dataRows) == 0 {
		return nil, fmt.Errorf("в Excel нет строк с данными (кроме заголовка)")
	}

	return extractReplacements(headers, dataRows)
}

func extractReplacements(excelHeaders []string, excelRows [][]string) ([]docx.PlaceholderMap, error) {
	allReplacements := make([]docx.PlaceholderMap, 0, len(excelRows) + 1)
	for _, row := range excelRows {
		replacements := docx.PlaceholderMap{}

		for idx, header := range excelHeaders {
			if header == "" {
				continue
			}
			if idx < len(row) {
				replacements[strings.TrimSpace(header)] = row[idx]
			}

			if header == "дата_договора" {
				textDate, err := dateMonthToText(row[idx])
				if err != nil {
					return nil, err
				}
				replacements["дата_договора_текст"] = textDate
			}
		}
		allReplacements = append(allReplacements, replacements)
	}

	return allReplacements, nil
}

// readExcel читает первый лист: первая строка — заголовки (имена плейсхолдеров
// без { }), остальные строки — данные.
func readExcel(path string) ([]string, [][]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows(f.GetSheetList()[0])
	if err != nil {
		return nil, nil, err
	}

	headers := rows[0]
	dataRows := rows[1:]
	return headers, dataRows, nil
}
