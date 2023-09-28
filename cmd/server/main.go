package main

import (
	"github.com/DmytroKha/GoogleSheets/internal/app"
	"log"
)

const (
	credentialsFile = "gkey/gsheets-400309-3b7ce89fc721.json"
)

func main() {
	gApp, err := app.NewGoogleSheetsApp(credentialsFile)
	if err != nil {
		log.Fatalf("Unable to create Google Sheets app: %v", err)
	}

	gApp.Run()
}
