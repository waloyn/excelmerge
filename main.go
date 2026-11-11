package main

import (
	"log"

	"excelmerge/internal/ui"
)

func main() {
	app := ui.NewAppUI()
	if err := app.Run(); err != nil {
		log.Fatalf("启动失败: %v", err)
	}
}
