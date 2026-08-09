package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
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
		return nil, err
	}

	excelReplacements, err := excelReplacements(dir)
	if err != nil {
		return nil, err
	}

	return replaceWordReplacements(dir, excelReplacements)
}

// при разработке
func developDir() (string, error) {
	return os.Getwd()
}

// при go build
func prodDir() (string, error) {
	execPath, err := os.Executable()
	if err != nil {
		return "", fmt.Errorf("не удалось определить текущую папку: %w", err)
	}
	return filepath.Dir(execPath), nil
}

// waitForEnter не даёт консольному окну закрыться сразу после завершения
// программы, если пользователь запустил её двойным кликом (актуально для Windows).
func waitForEnter() {
	if runtime.GOOS == "windows" {
		fmt.Println("Нажмите Enter, чтобы закрыть окно...")
		bufio.NewReader(os.Stdin).ReadString('\n')
	}
}
