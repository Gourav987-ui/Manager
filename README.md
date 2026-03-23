# Test Sheet Manager

A simple web app to manage your manual test case Excel sheets (`.xlsx`).

## Features

- **List** all test sheets with Jira ticket, filename, size, and last modified
- **Search** by ticket key (e.g. `INVST-123`) or filename
- **Download** sheets
- **Upload** new `.xlsx` or `.xls` files
- **Delete** sheets

## Setup

1. Install dependencies:
   ```bash
   cd D:\TestSheetManager
   npm install
   ```

2. Configure the sheets folder in `config.json` (default: `D:\TestSheetManager\Testsheets`):
   ```json
   {
     "testSheetsPath": "D:\\TestSheetManager\\Testsheets",
     "port": 3456
   }
   ```

3. Start the server:
   ```bash
   npm start
   ```

4. Open http://localhost:3456 in your browser.

## Config

| Key             | Description                         | Default                        |
|-----------------|-------------------------------------|--------------------------------|
| `testSheetsPath`| Folder where Excel sheets are stored| `D:\TestSheetManager\Testsheets`|
| `port`          | HTTP port                           | `3456`                         |
