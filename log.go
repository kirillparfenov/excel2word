package main

import (
	"fmt"
	"strings"

	"github.com/lukasjarosch/go-docx"
)

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

	foundMiss := make([]string, 0, len(missPlaceholders)+1)
	for k := range missPlaceholders {
		foundMiss = append(foundMiss, k)
	}

	if len(missPlaceholders) > 0 {
		fmt.Print("\n==========\n")
		fmt.Printf("В WORD ПРИСУТСТВУЮТ ШАБЛОНЫ, "+
			"НО ОТСУТСТВУЮТ В EXCEL: \n%s\n\n", strings.Join(foundMiss, ",\n"))
	}
}
