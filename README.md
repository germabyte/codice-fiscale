# Google Sheets Italian Fiscal Code Calculator

This custom function for Google Sheets allows for easy calculation of the Italian Fiscal Code (Codice Fiscale).  It also provides a function to extract the data from a valid fiscal code.

**Purpose:** This tool simplifies the process of generating Italian Fiscal Codes within a Google Sheets environment. It eliminates the need for manual calculations or external tools, providing a quick and convenient solution directly within your spreadsheet. It can also be used to check the validity of a provided fiscal code (using the inverse function), and extract information from a valid code.

**Value Proposition:**  This program is beneficial for anyone who needs to generate or validate Italian Fiscal Codes regularly, such as administrative staff, human resources personnel, or anyone working with Italian personal data.  It saves time and reduces the risk of errors compared to manual calculations.

## Dependencies (Required Software)

*   **Google Account:** You need a Google account to access and use Google Sheets.  Create one for free at [https://accounts.google.com/signup](https://accounts.google.com/signup).
*   **Google Sheets:**  This program is designed to run as a custom function within Google Sheets. Access Google Sheets from your Google account by going to [https://sheets.google.com](https://sheets.google.com).
* **`comuni` sheet.** A sheet named 'comuni' with three columns: Comune, Provincia, and Codice Catastale.

## Getting Started (Installation and Execution)

Since this is a Google Apps Script custom function, "installation" is different from traditional software. There is no ZIP file to download. You simply copy and paste the provided code into your Google Sheet's Script editor.

1.  **Open a Google Sheet:** Open the Google Sheet where you want to use the Fiscal Code calculator, or create a new one.

2.  **Open the Script Editor:**
    *   Go to "Tools" > "Script editor".

3.  **Copy and Paste the Code:**
    *   Delete any existing code in the Script editor.
    *   Copy the complete code provided for the `CODICEFISCALE` *and* `CODICEFISCALEINVERSO` functions (including all helper functions) and paste it into the Script editor.

4.  **Create the 'comuni' Sheet:**
      *   Click on the **+** button in your Google Sheet, to open a second Sheet.
      *   Rename the sheet "comuni" (without the quotes).
      *   In the first row (the header row), create three headers, in these three columns.
            * Column A: "Comune"
            * Column B: "Provincia"
            * Column C: "Codice"
    *   Populate the data starting from the *second* row (skipping the header row).  You'll need to populate this sheet with data listing the "Comune" (Municipality), "Provincia" (Province), and "Codice" (Codice Catastale, the municipality code) for each location.  You can find resources online for Italian municipality codes. *This is a crucial step; the code will not work without a properly populated "comuni" sheet.*

5.  **Save the Script:**
    *   Click on the "Save" icon (it looks like a floppy disk).  You may be prompted to give your project a name.  You can use a name like "FiscalCodeCalculator".

6.  **Refresh your Google Sheet.** Close and then re-open the Google Sheet to allow time to integrate the changes.

That's it! The `CODICEFISCALE()` and `CODICEFISCALEINVERSO()` functions are now available as custom functions in your Google Sheet.

## Using the Program (User Guide)

### `CODICEFISCALE()` (Calculating the Fiscal Code)

The `CODICEFISCALE()` function takes five arguments:

*   **`surname`** (Text): The person's surname.
*   **`name`** (Text): The person's first name.
*   **`birthdate`** (Text, Number, or Date): The person's birthdate.  You can enter this in several ways:
    *   As text in the format `DD/MM/YYYY` (e.g., `25/12/1980`).
    *   As a date value recognized by Google Sheets (usually by entering it in a cell formatted as a date).
    * As a serial date.
*   **`gender`** (Text): The person's gender, either "M" for Male or "F" for Female.
*   **`comune`** (Text): The name of the Italian municipality (Comune) where the person was born. *This must exactly match an entry in the first column of the `comuni` sheet.*

**To use the function:**

1.  In an empty cell, type `=CODICEFISCALE(`.

2.  Enter the five arguments, separated by commas. For example:

    `=CODICEFISCALE("Rossi", "Mario", "25/12/1980", "M", "Roma")`

3.  Press "Enter".

4.  The calculated Fiscal Code will appear in the cell. If there's an error (e.g., a missing or invalid input, or the comune is not found in your list), the cell will display an error message starting with "ERROR: ".

**Output:**

The function returns a 16-character string representing the calculated Italian Fiscal Code, or an error message if there was a problem with the input data or if the 'comune' provided is not present on the dedicated sheet.

### `CODICEFISCALEINVERSO()` (Extracting Data from a Fiscal Code)

The `CODICEFISCALEINVERSO()` function takes one argument:

*   **`codiceFiscale`** (Text): A 16-character Italian Fiscal Code.

**To use the function:**

1. In an empty cell type: =`CODICEFISCALEINVERSO(`.

2. Type the fiscal code into the parenthesis. For example:
 =`CODICEFISCALEINVERSO("RSSMRA80T25H501U")`

3.  Press "Enter".

The program will then provide an output which will require separate cells, detailing the Surname, Name, Birthdate, Gender, and Comune from the imputed Fiscal Code. Should the fiscal code inputted have an error, it will be reflected in an ERROR return.

**Output:**

The `CODICEFISCALEINVERSO` function returns the original inputs, so long as they can be found within the `comuni` sheet. If there is no 'comune' that is represented by the Catastale Code presented, the program will return an error. It will also return an error if the input does not match exactly 16 characters in length.

## Use Cases and Examples

**1. Employee Onboarding:**

*   **Situation:**  A company is hiring a new employee and needs to generate their Fiscal Code for administrative purposes.
*   **Example:** The employee's information is: Surname: `Bianchi`, Name: `Giulia`, Birthdate: `15/06/1992`, Gender: `F`, Comune: `Milano`.
*   **Input:**  `=CODICEFISCALE("Bianchi", "Giulia", "15/06/1992", "F", "Milano")`
*   **Expected Output:**  `BNCGLI92H55F205Z` (Note: The last character is a checksum and might vary).

**2. Data Verification:**

*   **Situation:** An organization has a list of Fiscal Codes and needs to extract the birthdate and gender from each code for reporting.
*    **Example:** The codice fiscale is RSSMRA80T25H501U
*    **Input:** `=CODICEFISCALEINVERSO("RSSMRA80T25H501U")`
*   **Expected Output:** Will return data as such:
    *  Cell 1: RSS
    *  Cell 2: MRA
    *  Cell 3: 25/12/80
    *  Cell 4: M
    *  Cell 5: ROMA

**3. Database Population:**

*   **Situation:**  A user is building a database of Italian contacts and wants to automatically generate Fiscal Codes based on the information they have.
*   **Example:** The user has the following data in their spreadsheet:  Surname in cell A2 (`Ferrari`), Name in cell B2 (`Luca`), Birthdate in cell C2 (`02/03/1975`), Gender in cell D2 (`M`), and Comune in cell E2 (`Bologna`).
*   **Input (in cell F2):** `=CODICEFISCALE(A2, B2, C2, D2, E2)`
*   **Expected Output:** `FRRLCU75C02A944Y` (Note: The last character is a checksum and might vary).

## Disclaimer

This repository is subject to updates at any time. These updates may render portions of this README file outdated. No guarantee is made that the README will be updated to reflect changes in the repository. The `CODICEFISCALE` and `CODICEFISCALEINVERSO` custom functions are provided "as is" without warranty of any kind. The accuracy of the calculated Fiscal Code depends entirely on the accuracy and completeness of the input data provided by the user and the information within the `comuni` Sheet, as well as the user following this guide correctly.
