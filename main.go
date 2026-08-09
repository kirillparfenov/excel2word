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
	//dir, err := developDir()
	dir, err := prodDir()
	if err != nil {
		return nil, fmt.Errorf("не удалось определить текущую папку: %w", err)
	}

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

	outDir, err := outputDir(dir)
	if err != nil {
		return nil, err
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
	defer doc.Close()
	logMissPlaceholders(doc, allReplacements[0])

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

// при разработке
func developDir() (string, error) {
	return os.Getwd()
}

// при go build
func prodDir() (string, error) {
	execPath, err := os.Executable() //при go build
	if err != nil {
		return "", fmt.Errorf("не удалось определить текущую папку: %w", err)
	}
	return filepath.Dir(execPath), nil
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

// outputDir исходная директория для сохранения новых word-файлов
func outputDir(dir string) (outDir string, err error) {
	outDir = filepath.Join(dir, "output")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return "", fmt.Errorf("не удалось создать папку output: %w", err)
	}
	return
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

			if header == "дата_договора" {
				textDate, err := numbersToText(row[idx])
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

var monthMap = map[string]string{
	"01": "января",
	"02": "февраля",
	"03": "марта",
	"04": "апреля",
	"05": "мая",
	"06": "июня",
	"07": "июля",
	"08": "августа",
	"09": "сентября",
	"10": "октября",
	"11": "ноября",
	"12": "декабря",
}

// numbersToText перевод месяца в дате в текст
// 23.04.2026 -> 23 апреля 2026
func numbersToText(date string) (string, error) {
	arrDate := strings.Split(date, ".")
	if len(arrDate) != 3 {
		return "", fmt.Errorf("дата должна содержать день.месяц.год")
	}
	month := arrDate[1]
	if len(month) != 2 {
		return "", fmt.Errorf("месяц должен включать два числа. А включает %d", len(month))
	}

	return fmt.Sprintf("%s %s %s", arrDate[0], monthMap[month], arrDate[2]), nil
}

// LogMissPlaceholders логирует пропущенные плейсхолдеры в word.
// Клиент должен дополнить excel этими плейсхолдерами
func logMissPlaceholders(doc *docx.Document, replacements docx.PlaceholderMap) {
	placeholders, err := doc.GetPlaceHoldersList()
	if err != nil {
		fmt.Printf("ошибка во время получения плейсхолдеров %s\n", err)
	}

	missPlaceholders := make(map[string]struct{})
	for _, placeholder := range placeholders {
		res := placeholder[1 : len(placeholder)-1]
		if _, ok := replacements[res]; !ok {
			missPlaceholders[res] = struct{}{}
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
