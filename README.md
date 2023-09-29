# Console Application

This console application allows you to interact with Google Sheets and perform the following functionalities:

## Choose an Action
Upon launching the application, you will be prompted to choose one of the following actions:
1. Download data from Google Sheets and display it on the console.
2. Create a new sheet in Google Sheets.
3. Save data in Google Sheets.
4. Change the name of an existing sheet in Google Sheets.
5. Update data in Google Sheets.
6. Delete a sheet in Google Sheets.
7. Delete data from a specified range in Google Sheets.

## How to Use
After selecting an action, the program will request necessary information from you and execute the corresponding operation on Google Sheets. Below is a description of each action:

### 1. Download Data from Google Sheets
This operation allows you to download data from a specified sheet and range in Google Sheets and display it on the console. You will need to provide the document ID, sheet name(s), and the range from which to download data.

### 2. Create a New Sheet in Google Sheets
This operation enables you to create a new sheet with a specified name in Google Sheets. You will need to provide the document ID and the name for the new sheet.

### 3. Save Data in Google Sheets
This operation allows you to add new data to a specified sheet and column in Google Sheets. You will need to provide the document ID, sheet name, column name, and the data to be added.

### 4. Change the Name of an Existing Sheet in Google Sheets
This operation allows you to change the name of an existing sheet in Google Sheets. You will need to provide the document ID, the current sheet name, and the new sheet name.

### 5. Update Data in Google Sheets
This operation lets you update data in a specified range of a sheet in Google Sheets. You will need to provide the document ID, sheet name, range, and the data to update.

### 6. Delete a Sheet in Google Sheets
This operation allows you to delete a specified sheet from Google Sheets. You will need to provide the document ID and the name of the sheet to delete.

### 7. Delete Data from a Specified Range in Google Sheets
This operation enables you to delete data from a specified range in a sheet in Google Sheets. You will need to provide the document ID, sheet name, and the range from which to delete data.

## Configuration Requirements
To use this application, you need to have an authentication configuration file from Google Cloud (e.g., `gsheets-credentials.json`). This file should be located in the `gkey` folder. You should also have Google Sheets API access set up in your Google account. Also, when you start working with the application, you need to initialize the spreadsheetID, it can be found in the link line to the Google Sheet page, for example, in the above line it will be "1eXuEKYP35Y94PIgI6-Nym46ke7GY4_R7ZUXig1zaiBM": https://docs.google.com/spreadsheets/d/1eXuEKYP35Y94PIgI6-Nym46ke7GY4_R7ZUXig1zaiBM/edit#gid=1635680430 

## Installation and Execution
1. Clone the repository to your computer.
2. Run the `go run main.go` command in your terminal to execute the program.

## Author
Dmytro Khlypun