# 🇮🇹 Italian Fiscal Code Calculator for Google Sheets

## 1. Introduction and Purpose

### 📌 Introduction  
This program provides two powerful custom functions for **Google Sheets** that allow users to **calculate** and **decode** an **Italian Fiscal Code (Codice Fiscale)** directly within their spreadsheet.

### 🎯 Purpose & Problem Statement  
Italian citizens are assigned a unique 16-character Fiscal Code based on their personal data (name, surname, gender, date, and place of birth). Manually calculating or interpreting this code is error-prone and time-consuming. This tool automates that process within a spreadsheet, removing the need for manual calculations or third-party websites.

### 💡 Value Proposition  
- Easily generate a legally valid Italian Fiscal Code directly in a spreadsheet.  
- Reverse an existing Fiscal Code to extract meaningful personal information.  
- Automatically pulls municipality codes from a reference sheet.  
- Eliminates manual errors and saves time for citizens, businesses, and public offices.

---

## 2. Dependencies (Required Software/Libraries)

### ✅ Required Platform
This script is written in **Google Apps Script** and is designed to run within **Google Sheets**. No additional software is required beyond a standard web browser and a Google account.

### 📄 Required Sheet: "comuni"
The script relies on a worksheet named `comuni` that must contain:
- Column A: Municipality Name  
- Column C: Codice Catastale (Italian municipal code)  

Ensure this sheet is correctly populated. It acts as a lookup table to match municipality names to official codes.

---

## 3. Getting Started (Installation & Execution)

### 📥 Downloading the Script

1. Open or create a Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Paste the entire code provided into the script editor.
4. Save the project (e.g., name it `CodiceFiscaleCalculator`).
5. Return to the spreadsheet.

### 📂 Setting Up the Comune Sheet
1. Add a new sheet and rename it to `comuni`.
2. Populate it with data:  
   - **Column A**: Municipality name (e.g., *Roma*)  
   - **Column C**: Codice catastale (e.g., *H501*)  

   > Column B can remain blank or contain extra info—it’s ignored by the program.

### ▶️ Running the Program

#### 🧮 To Generate a Fiscal Code:
Use the formula in a cell:
```
=CODICEFISCALE("Rossi", "Mario", "01/01/1980", "M", "Roma")
```

#### 🔁 To Decode a Fiscal Code:
Use the formula in a cell:
```
=CODICEFISCALEINVERSO("RSSMRA80A01H501U")
```

Both functions return results directly in the spreadsheet.

---

## 4. User Guide (How to Effectively Use the Program)

### ✅ Function 1: CODICEFISCALE

#### 📥 Input Parameters:
| Parameter   | Format       | Description                            |
|-------------|--------------|----------------------------------------|
| `surname`   | Text         | Last name (e.g., "Rossi")              |
| `name`      | Text         | First name (e.g., "Mario")             |
| `birthdate` | DD/MM/YYYY, number, or Date object | Person’s birthdate |
| `gender`    | "M" or "F"   | Gender                                  |
| `comune`    | Text         | Birth municipality (must match "comuni" sheet) |

#### 📤 Output:
Returns a **16-character Codice Fiscale**, or an error message if the data is invalid.

---

### ✅ Function 2: CODICEFISCALEINVERSO

#### 📥 Input:
- A valid **16-character Codice Fiscale** (e.g., `RSSMRA80A01H501U`)

#### 📤 Output:
An object with:
- Surname code (first 3 letters)
- Name code (next 3 letters)
- Birthdate in DD/MM/YY format
- Gender ("M" or "F")
- Birth comune name

**Note:** This is **not guaranteed** to give the full original name, just code fragments used to construct the fiscal code.

---

## 5. Use Cases and Real-World Examples

### 🧑‍⚖️ Use Case 1: Citizen Identity Generation  
**Scenario:** A public office needs to quickly generate Codici Fiscali for multiple new Italian citizens.  
**Input:**  
```excel
=CODICEFISCALE("Bianchi", "Giulia", "23/05/1995", "F", "Milano")
```  
**Output:**  
`BNCGLI95E63F205Y`  

---

### 🏢 Use Case 2: Reverse Identity Check  
**Scenario:** A bank receives a Codice Fiscale and needs to extract info for verification.  
**Input:**  
```excel
=CODICEFISCALEINVERSO("BNCGLI95E63F205Y")
```  
**Output:**  
```json
{
  "surname": "BNC",
  "name": "GLI",
  "birthdate": "23/05/95",
  "gender": "F",
  "comune": "Milano"
}
```

---

### 🧾 Use Case 3: Form Validation in Business Apps  
**Scenario:** An HR spreadsheet auto-generates Codice Fiscale from staff records.  
**Input:**  
Each row pulls employee details from columns A–E:  
```excel
=CODICEFISCALE(A2, B2, C2, D2, E2)
```  
**Output:**  
Fiscal Code appears in column F.

---

## 6. Disclaimer & Important Notices

- This repository and its contents may be updated at any time without notice.
- Such updates may render parts of the provided README file obsolete.
- No commitment is made to maintain or update the README to reflect future changes.
- The provided code is delivered **"as-is"** and no guarantees—explicit or implied—are made regarding functionality, reliability, compatibility, or correctness.
