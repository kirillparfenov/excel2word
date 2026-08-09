package main

import (
	"fmt"
	"strings"
)

var monthToText = map[string]string{
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

// dateMonthToText перевод месяца в дате в текст
// 23.04.2026 -> 23 апреля 2026
func dateMonthToText(date string) (string, error) {
	arrDate := strings.Split(date, ".")
	if len(arrDate) != 3 {
		return "", fmt.Errorf("дата должна содержать день.месяц.год")
	}
	day := arrDate[0]
	month := arrDate[1]
	if len(month) != 2 {
		return "", fmt.Errorf("месяц должен включать два числа. А включает %d", len(month))
	}
	year := arrDate[2]

	return fmt.Sprintf("%s %s %s", day, monthToText[month], year), nil
}
