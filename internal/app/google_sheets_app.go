package app

import (
	"bufio"
	"context"
	"fmt"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"os"
	"strings"
)

type GoogleSheetsApp struct {
	srv           *sheets.Service
	spreadsheetID string
}

func NewGoogleSheetsApp(credentialsFile string) (*GoogleSheetsApp, error) {
	ctx := context.Background()
	clientOption := option.WithCredentialsFile(credentialsFile)
	client, err := sheets.NewService(ctx, clientOption)
	if err != nil {
		return nil, err
	}
	return &GoogleSheetsApp{srv: client}, nil
}

func (app *GoogleSheetsApp) Run() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter spreadsheet ID:")
	scanner.Scan()
	app.spreadsheetID = scanner.Text()

	for {
		fmt.Println("Options:")
		fmt.Println("1. Download data from Google Sheets")
		fmt.Println("2. Create a new sheet in Google Sheets")
		fmt.Println("3. Save data in Google Sheets")
		fmt.Println("4. Change sheet name in Google Sheets")
		fmt.Println("5. Update data in Google Sheets")
		fmt.Println("6. Delete sheet in Google Sheets")
		fmt.Println("7. Delete data in Google Sheets")
		fmt.Println("8. Exit")
		fmt.Print("Choose your option (1-8): ")
		scanner.Scan()
		variant := scanner.Text()

		switch variant {
		case "1":
			app.handleDownloadData()
		case "2":
			app.handleCreateSheet()
		case "3":
			app.handleSaveData()
		case "4":
			app.handleChangeSheetName()
		case "5":
			app.handleUpdateData()
		case "6":
			app.handleDeleteSheet()
		case "7":
			app.handleDeleteData()
		case "8":
			fmt.Println("Goodbye!")
			return
		default:
			fmt.Println("Invalid option. Please choose a valid option (1-8).")
		}
	}
}

func (app *GoogleSheetsApp) handleDownloadData() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet names (comma-separated): ")
	scanner.Scan()
	sheetNames := scanner.Text()

	fmt.Print("Enter sheet range (e.g., A1:D5): ")
	scanner.Scan()
	sheetRange := scanner.Text()

	splipSheetNames := strings.Split(sheetNames, ",")

	for _, sheetName := range splipSheetNames {
		err := app.readAndPrintData(sheetName, sheetRange)
		if err != nil {
			log.Printf("Error downloading data: %v", err)
		}
	}
}

func (app *GoogleSheetsApp) handleCreateSheet() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet name: ")
	scanner.Scan()
	sheetName := scanner.Text()

	err := app.createSheet(sheetName)
	if err != nil {
		log.Printf("Error creating sheet: %v", err)
	} else {
		fmt.Printf("Sheet '%s' created successfully.\n", sheetName)
	}
}

func (app *GoogleSheetsApp) handleSaveData() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet name: ")
	scanner.Scan()
	sheetName := scanner.Text()

	fmt.Print("Enter column name: ")
	scanner.Scan()
	columnName := scanner.Text()

	fmt.Print("Enter text to add: ")
	scanner.Scan()
	newText := scanner.Text()

	err := app.saveData(sheetName, columnName, newText)
	if err != nil {
		log.Printf("Error saving data: %v", err)
	} else {
		fmt.Printf("Data saved successfully to sheet '%s'.\n", sheetName)
	}
}

func (app *GoogleSheetsApp) handleChangeSheetName() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter existing sheet name: ")
	scanner.Scan()
	existingSheetName := scanner.Text()

	fmt.Print("Enter new sheet name: ")
	scanner.Scan()
	newSheetName := scanner.Text()

	err := app.changeSheetName(existingSheetName, newSheetName)
	if err != nil {
		log.Printf("Error changing sheet name: %v", err)
	} else {
		fmt.Printf("Sheet name changed successfully from '%s' to '%s'.\n", existingSheetName, newSheetName)
	}
}

func (app *GoogleSheetsApp) handleUpdateData() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet name: ")
	scanner.Scan()
	sheetName := scanner.Text()

	fmt.Print("Enter sheet range: ")
	scanner.Scan()
	sheetRange := scanner.Text()

	fmt.Print("Enter text to update: ")
	scanner.Scan()
	newText := scanner.Text()

	err := app.updateData(sheetName, sheetRange, newText)
	if err != nil {
		log.Printf("Error updating data: %v", err)
	} else {
		fmt.Printf("Data updated successfully in sheet '%s' at range '%s'.\n", sheetName, sheetRange)
	}
}

func (app *GoogleSheetsApp) handleDeleteSheet() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet name: ")
	scanner.Scan()
	sheetName := scanner.Text()

	err := app.deleteSheet(sheetName)
	if err != nil {
		log.Printf("Error deleting sheet: %v", err)
	} else {
		fmt.Printf("Sheet '%s' has been deleted.\n", sheetName)
	}
}

func (app *GoogleSheetsApp) handleDeleteData() {
	scanner := bufio.NewScanner(os.Stdin)

	fmt.Print("Enter sheet name: ")
	scanner.Scan()
	sheetName := scanner.Text()

	fmt.Print("Enter sheet range (e.g., A1:D5): ")
	scanner.Scan()
	sheetRange := scanner.Text()

	err := app.deleteDataInRange(sheetName, sheetRange)
	if err != nil {
		log.Printf("Error deleting data: %v", err)
	} else {
		fmt.Printf("Data in range '%s' on sheet '%s' has been deleted.\n", sheetRange, sheetName)
	}
}

func (app *GoogleSheetsApp) readAndPrintData(sheetName string, sheetRange string) error {
	var readRange string
	if sheetRange != "" {
		readRange = sheetName + "!" + sheetRange
	} else {
		readRange = sheetName
	}

	resp, err := app.srv.Spreadsheets.Values.Get(app.spreadsheetID, readRange).Do()
	if err != nil {
		return err
	}

	if len(resp.Values) == 0 {
		fmt.Printf("No data found in sheet %s\n", sheetName)
		return nil
	}

	fmt.Printf("Data from Google Sheet (%s):\n", sheetName)
	for _, row := range resp.Values {
		fmt.Println(row)
	}

	return nil
}

func (app *GoogleSheetsApp) createSheet(sheetName string) error {
	request := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				AddSheet: &sheets.AddSheetRequest{
					Properties: &sheets.SheetProperties{
						Title: sheetName,
					},
				},
			},
		},
	}

	_, err := app.srv.Spreadsheets.BatchUpdate(app.spreadsheetID, request).Do()
	if err != nil {
		return err
	}
	return nil
}

func (app *GoogleSheetsApp) saveData(sheetName, columnName, newText string) error {
	readRange := sheetName + "!" + columnName + "1:" + columnName

	resp, err := app.srv.Spreadsheets.Values.Get(app.spreadsheetID, readRange).Do()
	if err != nil {
		return err
	}

	rowNumber := len(resp.Values) + 1

	writeRange := fmt.Sprintf("%s!%s%d", sheetName, columnName, rowNumber)

	valueRange := &sheets.ValueRange{
		Values: [][]interface{}{{newText}},
	}

	_, err = app.srv.Spreadsheets.Values.Append(app.spreadsheetID, writeRange, valueRange).ValueInputOption("RAW").Do()
	if err != nil {
		return err
	}
	return nil
}

func (app *GoogleSheetsApp) changeSheetName(existingSheetName, newSheetName string) error {
	sheetID := app.getSheetIDByName(existingSheetName)
	if sheetID == -1 {
		return fmt.Errorf("Sheet '%s' not found in Google Sheets.", existingSheetName)
	}

	requests := []*sheets.Request{
		{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Properties: &sheets.SheetProperties{
					Title:   newSheetName,
					SheetId: sheetID,
				},
				Fields: "title",
			},
		},
	}

	_, err := app.srv.Spreadsheets.BatchUpdate(app.spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return err
	}
	return nil
}

func (app *GoogleSheetsApp) getSheetIDByName(sheetName string) int64 {
	resp, err := app.srv.Spreadsheets.Get(app.spreadsheetID).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve spreadsheet: %v", err)
	}

	for _, sheet := range resp.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId
		}
	}

	return -1
}

func (app *GoogleSheetsApp) updateData(sheetName, sheetRange, newText string) error {
	writeRange := sheetName + "!" + sheetRange
	valueRange := &sheets.ValueRange{
		Values: [][]interface{}{{newText}},
	}
	_, err := app.srv.Spreadsheets.Values.Update(app.spreadsheetID, writeRange, valueRange).ValueInputOption("RAW").Do()
	if err != nil {
		return err
	}
	return nil
}

func (app *GoogleSheetsApp) deleteSheet(sheetName string) error {
	sheetID := app.getSheetIDByName(sheetName)
	if sheetID == -1 {
		return fmt.Errorf("Sheet '%s' not found in Google Sheets.", sheetName)
	}

	requests := []*sheets.Request{
		{
			DeleteSheet: &sheets.DeleteSheetRequest{
				SheetId: sheetID,
			},
		},
	}

	_, err := app.srv.Spreadsheets.BatchUpdate(app.spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return err
	}
	return nil
}

func (app *GoogleSheetsApp) deleteDataInRange(sheetName string, rangeToDelete string) error {
	clearRequest := &sheets.ClearValuesRequest{}
	_, err := app.srv.Spreadsheets.Values.Clear(app.spreadsheetID, sheetName+"!"+rangeToDelete, clearRequest).Do()

	if err != nil {
		return err
	}
	return nil
}
