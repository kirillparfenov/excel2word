package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/lukasjarosch/go-docx"
	"github.com/xuri/excelize/v2"
)

func main() {
	if createdDocs, err := run(); err != nil {
		fmt.Println("\nОШИБКА:", err)
	} else {
		fmt.Print("==========\n")
		fmt.Printf("Готово! Документы сохранены в папке 'output': \n%s", strings.Join(createdDocs, "\n"))
		fmt.Print("\n\n")
	}
	waitForEnter()
}

// Run запуск скрипта переноса данных из excel -> word на места плейсхолдеров
func run() ([]string, error) {
	//dir, err := os.Getwd() //при разработке
	execPath, err := os.Executable() //при go build
	if err != nil {
		return nil, fmt.Errorf("не удалось определить текущую папку: %w", err)
	}
	dir := filepath.Dir(execPath)

	excelPath, err := findFileByExt(dir, ".xlsx")
	if err != nil {
		return nil, err
	}
	wordPath, err := findFileByExt(dir, ".docx")
	if err != nil {
		return nil, err
	}

	fmt.Println("Excel-файл:", filepath.Base(excelPath))
	fmt.Println("Word-шаблон:", filepath.Base(wordPath))

	outDir := filepath.Join(dir, "output")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return nil, fmt.Errorf("не удалось создать папку output: %w", err)
	}

	// собираем все данные из excel в качестве реплейсментов для word-а
	allReplacements, err := readAllReplacements(excelPath)
	if err != nil {
		return nil, err
	}

	// логируем пропуски реплейсментов в excel -> word
	doc, err := docx.Open(wordPath)
	if err != nil {
		return nil, fmt.Errorf("ошибка при открытии word-файла: %w", err)
	}
	logMissPlaceholders(doc, allReplacements[0])
	defer doc.Close()

	createdDocs := make([]string, 0, len(allReplacements))
	for i, replacements := range allReplacements {
		outName := fmt.Sprintf("document_%03d.docx", i+1)
		outPath := filepath.Join(outDir, outName)

		if err := generateDocx(wordPath, outPath, replacements); err != nil {
			return nil, fmt.Errorf("ошибка в процессе генерации .docx: %w", err)
		}
		createdDocs = append(createdDocs, outName)
	}

	return createdDocs, nil
}

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

// readAllReplacements достаем все колонки со строками из excel в виде структуры шаблонов для word
func readAllReplacements(excelPath string) ([]docx.PlaceholderMap, error) {
	headers, dataRows, err := readExcel(excelPath)
	if err != nil {
		return nil, fmt.Errorf("не удалось прочитать Excel: %w", err)
	}
	if len(dataRows) == 0 {
		return nil, fmt.Errorf("в Excel нет строк с данными (кроме заголовка)")
	}

	allReplacements := make([]docx.PlaceholderMap, 0, len(dataRows))
	for _, row := range dataRows {
		replacements := docx.PlaceholderMap{}

		for idx, header := range headers {
			if header == "" {
				continue // пропускаем пустые заголовки столбцов
			}
			if idx < len(row) {
				replacements[strings.TrimSpace(header)] = row[idx]
			}
		}
		allReplacements = append(allReplacements, replacements)
	}
	return allReplacements, nil
}

// LogMissPlaceholders логирует пропущенные плейсхолдеры в word.
// Клиент должен дополнить excel этими плейсхолдерами
func logMissPlaceholders(doc *docx.Document, replacements docx.PlaceholderMap) {
	placeholders, err := doc.GetPlaceHoldersList()
	if err != nil {
		fmt.Printf("ошибка во время получения плейсхолдеров %s\n", err)
	}

	missPlaceholders := make(map[string]int)
	for _, placeholder := range placeholders {
		res := placeholder[1 : len(placeholder)-1]
		if replacements[res] == nil {
			missPlaceholders[res] = 1
		}
	}

	foundMiss := make([]string, 0, len(missPlaceholders))
	for k := range missPlaceholders {
		foundMiss = append(foundMiss, k)
	}

	if len(missPlaceholders) > 0 {
		fmt.Print("\n==========\n")
		fmt.Printf("В WORD ПРИСУТСТВУЮТ ШАБЛОНЫ, "+
			"НО ОТСУТСТВУЮТ В EXCEL: \n%s\n\n", strings.Join(foundMiss, ",\n"))
	}
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

// waitForEnter не даёт консольному окну закрыться сразу после завершения
// программы, если пользователь запустил её двойным кликом (актуально для Windows).
func waitForEnter() {
	if runtime.GOOS == "windows" {
		fmt.Println("Нажмите Enter, чтобы закрыть окно...")
		bufio.NewReader(os.Stdin).ReadString('\n')
	}
}
