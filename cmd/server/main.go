package main

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

func main() {
	// Створюємо новий об'єкт сканера для зчитування введеного рядка з консолі.
	scanner := bufio.NewScanner(os.Stdin)

	// Шлях до файлу конфігурації автентифікації, який ви отримаєте від Google Cloud.
	credentialsFile := "gkey/gsheets-400309-3b7ce89fc721.json"

	// Створюємо контекст та використовуємо файл конфігурації для автентифікації.
	ctx := context.Background()
	clientOption := option.WithCredentialsFile(credentialsFile)
	client, err := sheets.NewService(ctx, clientOption)
	if err != nil {
		log.Fatalf("Unable to create Google Sheets service: %v", err)
	}

	// ID вашого Google Sheets документа та ім'я аркуша, з яким ви працюєте.
	var variant, yesNo, spreadsheetID, sheetsNames, sheetName, sheetRange, newText, columnName string

	fmt.Println("1. Download data from Google Sheets")
	fmt.Println("2. Saving data in Google Sheets")
	fmt.Println("3. Update data in Google Sheets")
	fmt.Println("4. Delete data in Google Sheets")
	fmt.Println("Choose your variant (enter the number):")
	scanner.Scan()
	variant = scanner.Text()

	if variant == "1" {

		fmt.Println("Enter spreadsheet ID:")
		//fmt.Scan(&spreadsheetID) //"1eXuEKYP35Y94PIgI6-Nym46ke7GY4_R7ZUXig1zaiBM"
		scanner.Scan()
		spreadsheetID = scanner.Text()
		fmt.Println("Enter sheets names:")
		//fmt.Scan(&sheetsNames) // "Аркуш1"
		scanner.Scan()
		sheetsNames = scanner.Text()
		fmt.Println("Enter sheet range:")
		//fmt.Scan(&sheetRange) //"!A1:D5"
		scanner.Scan()
		sheetRange = scanner.Text()

		// Викликаємо функцію для завантаження та виведення даних.
		readAndPrintData(client, spreadsheetID, sheetsNames, sheetRange)

	} else if variant == "2" {
		fmt.Println("Do you want to create new sheet? (y - Yes/n - No):")
		//fmt.Scan(&yesNo)
		scanner.Scan()
		yesNo = scanner.Text()
		if yesNo == "y" {

		} else {

			fmt.Println("Enter spreadsheet ID:")
			//fmt.Scan(&spreadsheetID)
			scanner.Scan()
			spreadsheetID = scanner.Text()
			fmt.Println("Enter sheet name:")
			//fmt.Scan(&sheetName)
			scanner.Scan()
			sheetName = scanner.Text()
			fmt.Println("Enter column name:")
			//fmt.Scan(&columnName)
			scanner.Scan()
			columnName = scanner.Text()
			fmt.Println("Enter text to add:")
			//a, _ := fmt.Scan(&newText)
			scanner.Scan()
			newText = scanner.Text()

			// Викликаємо функцію для збереження даних в аркуші.
			writeData(client, spreadsheetID, sheetName, columnName, []interface{}{newText})

		}
	}

}

func readAndPrintData(srv *sheets.Service, spreadsheetID string, sheetsNames string, sheetRange string) {
	// Зчитуємо дані з вказаного аркуша.
	splipSheetsNames := strings.Split(sheetsNames, ",")
	if len(splipSheetsNames) != 0 {
		for _, v := range splipSheetsNames {
			readRange := v + sheetRange
			resp, err := srv.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
			if err != nil {
				log.Fatalf("Unable to retrieve data from sheet: %v", err)
			}

			// Перевіряємо, чи маємо дані.
			if len(resp.Values) == 0 {
				fmt.Println("No data found.")
				return
			}

			// Виводимо отримані дані на консоль.
			fmt.Printf("Data from Google Sheet (%s): \n", v)
			for _, row := range resp.Values {
				fmt.Println(row)
			}
		}
	}

}

func writeData(srv *sheets.Service, spreadsheetID string, sheetName string, columnName string, data []interface{}) {
	//// Створюємо новий аркуш з вказаним ім'ям у документі.
	//newSheet := &sheets.Sheet{
	//	Properties: &sheets.SheetProperties{
	//		Title: sheetName,
	//	},
	//}
	//_, err := srv.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
	//	Requests: []*sheets.Request{
	//		{
	//			AddSheet: newSheet,
	//		},
	//	},
	//}).Do()
	//if err != nil {
	//	log.Fatalf("Unable to create new sheet: %v", err)
	//}
	readRange := sheetName + "!" + columnName + "1:" + columnName

	resp, err := srv.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
	if err != nil {
		log.Fatalf("Unable to retrieve data from sheet: %v", err)
	}

	// Calculate the row number where new data should be appended.
	// Add 2 to the length of the existing data to account for the header row and 1-based indexing.
	rowNumber := len(resp.Values) + 1

	// Specify the range to append data to.
	writeRange := fmt.Sprintf("%s!%s%d", sheetName, columnName, rowNumber)

	// Create a ValueRange with the data to append.
	valueRange := &sheets.ValueRange{
		Values: [][]interface{}{data},
	}

	// Append the data to the specified sheet.
	_, err = srv.Spreadsheets.Values.Append(spreadsheetID, writeRange, valueRange).ValueInputOption("RAW").Do()
	if err != nil {
		log.Fatalf("Unable to append data to sheet: %v", err)
	}
}
