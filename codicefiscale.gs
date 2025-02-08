/**
 * Calculates the Italian Fiscal Code (Codice Fiscale) based on input data.
 *
 * @param {string} surname The person's surname.
 * @param {string} name The person's name.
 * @param {string|number|Date} birthdate The person's birthdate.  Can be DD/MM/YYYY, a number (serial date), or a Date object.
 * @param {string} gender The person's gender (M or F).
 * @param {string} comune The person's birth comune (municipality).
 * @return {string} The calculated Fiscal Code.
 * @customfunction
 */
function CODICEFISCALE(surname, name, birthdate, gender, comune) {
  if (!surname || !name || !birthdate || !gender || !comune) {
    throw new Error("All parameters (surname, name, birthdate, gender, comune) are required.");
  }

  // Helper functions (moved inside the main function for Apps Script compatibility)
  function estraiConsonanti(parola) {
    let consonanti = "";
    for (let char of parola.toUpperCase()) {
      if ("BCDFGHJKLMNPQRSTVWXYZ".includes(char)) {
        consonanti += char;
      }
    }
    return consonanti;
  }

  function estraiVocali(parola) {
    let vocali = "";
    for (let char of parola.toUpperCase()) {
      if ("AEIOU".includes(char)) {
        vocali += char;
      }
    }
    return vocali;
  }


  function meseCodice(mese) {
    const codiciMese = {
      '01': 'A', '02': 'B', '03': 'C', '04': 'D', '05': 'E', '06': 'H',
      '07': 'L', '08': 'M', '09': 'P', '10': 'R', '11': 'S', '12': 'T'
    };
    return codiciMese[mese];
  }

  function calcolaCarattereDiControllo(codiceParziale) {
    const valoriDispari = {
      '0': 1, '1': 0, '2': 5, '3': 7, '4': 9, '5': 13, '6': 15, '7': 17, '8': 19, '9': 21,
      'A': 1, 'B': 0, 'C': 5, 'D': 7, 'E': 9, 'F': 13, 'G': 15, 'H': 17, 'I': 19, 'J': 21,
      'K': 2, 'L': 4, 'M': 18, 'N': 20, 'O': 11, 'P': 3, 'Q': 6, 'R': 8, 'S': 12, 'T': 14,
      'U': 16, 'V': 10, 'W': 22, 'X': 25, 'Y': 24, 'Z': 23
    };

    const valoriPari = {
      '0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
      'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9,
      'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19,
      'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25
    };

    let somma = 0;
    for (let i = 0; i < codiceParziale.length; i++) {
      const char = codiceParziale[i];
      if ((i + 1) % 2 === 1) { // Odd positions (1, 3, 5, ...)  (i+1 because index starts at 0)
        somma += valoriDispari[char] || 0;
      } else { // Even positions (2, 4, 6, ...)
        somma += valoriPari[char] || 0;
      }
    }

    const caratteriControllo = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    return caratteriControllo[somma % 26];
  }


  function getCodiceCatastale(comune) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("comuni");
    if (!sheet) {
      throw new Error("Sheet 'comuni' not found.");
    }

    const data = sheet.getDataRange().getValues();
    // Start from row 1 to skip header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toUpperCase() === comune.toUpperCase()) { // Case-insensitive comparison
        return data[i][2]; // Codice is in the third column
      }
    }
    return null; // Or throw an error: throw new Error("Comune not found in 'comuni' sheet.");
  }

  function formatDate(date) {
      if (typeof date === 'string') {
        // Assume DD/MM/YYYY format if already a string
        return date;
      }
      if (typeof date === 'number') {
        // Convert serial number to Date object
        date = new Date(Math.round((date - 25569) * 86400 * 1000));
      }
      if (date instanceof Date) {
            const d = date.getDate();
            const m = date.getMonth() + 1; // Month is 0-indexed
            const y = date.getFullYear();
            return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
        }

      throw new Error("Invalid date format.");
  }



  // --- Main calculation logic ---
  try {
    const cognomeCod = (estraiConsonanti(surname) + estraiVocali(surname) + "XXX").substring(0, 3);

    let nomeCod = estraiConsonanti(name);
    if (nomeCod.length >= 4) {
      nomeCod = nomeCod[0] + nomeCod[2] + nomeCod[3];
    }
    nomeCod = (nomeCod + estraiVocali(name) + "XXX").substring(0, 3);


    // ********* DATE HANDLING IMPROVEMENT **********
    const formattedDate = formatDate(birthdate);
    const birthdateParts = formattedDate.split("/");

    if (birthdateParts.length !== 3) {
      throw new Error("Invalid birthdate format. Use DD/MM/YYYY.");
    }
    const day = parseInt(birthdateParts[0], 10);
    const month = birthdateParts[1];
    const year = birthdateParts[2];
     if (isNaN(day) || day < 1 || day > 31) {
      throw new Error("Invalid day in birthdate.");
    }
    if (isNaN(parseInt(month, 10)) || parseInt(month, 10) < 1 || parseInt(month, 10) > 12) {
      throw new Error("Invalid month in birthdate.");
    }

    // ***********************************************

    const annoCod = year.substring(2);
    const meseCod = meseCodice(month);
    let giornoCod = (gender.toUpperCase() === 'F') ? String(day + 40).padStart(2, '0') : String(day).padStart(2, '0');


    const codiceCatastale = getCodiceCatastale(comune);
    if (!codiceCatastale) {
      throw new Error("Codice catastale not found for " + comune);
    }

    const codiceParziale = cognomeCod + nomeCod + annoCod + meseCod + giornoCod + codiceCatastale;
    const carattereControllo = calcolaCarattereDiControllo(codiceParziale);

    return codiceParziale + carattereControllo;

  } catch (error) {
    return "ERROR: " + error.message;
  }
}



/**
 * Extracts personal data from an Italian Fiscal Code.
 *
 * @param {string} codiceFiscale The 16-character Italian Fiscal Code.
 * @return {Object} An object containing the extracted data (surname, name, birthdate, gender, comune).
 * @customfunction
 */
function CODICEFISCALEINVERSO(codiceFiscale) {
    if (!codiceFiscale || codiceFiscale.length !== 16) {
        throw new Error("Invalid Fiscal Code.  It must be 16 characters long.");
    }

    codiceFiscale = codiceFiscale.toUpperCase();

  // Helper functions (inside the main function for Apps Script)
  function meseInverso(codiceMese) {
    const codiciMeseInverso = {
      'A': '01', 'B': '02', 'C': '03', 'D': '04', 'E': '05', 'H': '06',
      'L': '07', 'M': '08', 'P': '09', 'R': '10', 'S': '11', 'T': '12'
    };
    return codiciMeseInverso[codiceMese];
  }

  function calcolaCarattereDiControllo(codiceParziale) {
    const valoriDispari = {
      '0': 1, '1': 0, '2': 5, '3': 7, '4': 9, '5': 13, '6': 15, '7': 17, '8': 19, '9': 21,
      'A': 1, 'B': 0, 'C': 5, 'D': 7, 'E': 9, 'F': 13, 'G': 15, 'H': 17, 'I': 19, 'J': 21,
      'K': 2, 'L': 4, 'M': 18, 'N': 20, 'O': 11, 'P': 3, 'Q': 6, 'R': 8, 'S': 12, 'T': 14,
      'U': 16, 'V': 10, 'W': 22, 'X': 25, 'Y': 24, 'Z': 23
    };

    const valoriPari = {
      '0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
      'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9,
      'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19,
      'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25
    };

    let somma = 0;
    for (let i = 0; i < codiceParziale.length; i++) {
      const char = codiceParziale[i];
      if ((i + 1) % 2 === 1) { // Odd positions (1, 3, 5, ...)
        somma += valoriDispari[char] || 0;
      } else { // Even positions (2, 4, 6, ...)
        somma += valoriPari[char] || 0;
      }
    }

    const caratteriControllo = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    return caratteriControllo[somma % 26];
  }


    function getComuneByCodiceCatastale(codiceCatastale) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("comuni");
        if (!sheet) {
            throw new Error("Sheet 'comuni' not found.");
        }

        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][2] === codiceCatastale) {
                return data[i][0]; // Comune name is in the first column
            }
        }
        return null;  //Or: throw new Error("Comune not found for codice catastale: " + codiceCatastale);
    }


    // --- Main inverse calculation logic ---
    try {
        const cognomeCodice = codiceFiscale.substring(0, 3);
        const nomeCodice = codiceFiscale.substring(3, 6);
        const anno = codiceFiscale.substring(6, 8);
        const meseCodice = codiceFiscale.charAt(8);
        const mese = meseInverso(meseCodice);
        let giornoCodice = parseInt(codiceFiscale.substring(9, 11), 10);
        let sesso = 'M';
        if (giornoCodice > 40) {
            sesso = 'F';
            giornoCodice -= 40;
        }
        const giorno = String(giornoCodice).padStart(2, '0');

        const codiceCatastale = codiceFiscale.substring(11, 15);
        const comune = getComuneByCodiceCatastale(codiceCatastale);
          if (!comune) {
            throw new Error("Comune not found for codice catastale: " + codiceCatastale);
          }

        // Control character check
        const calculatedControlChar = calcolaCarattereDiControllo(codiceFiscale.substring(0, 15));
        if (calculatedControlChar !== codiceFiscale.charAt(15)) {
            throw new Error("Invalid control character.");
        }

        return {
            surname: cognomeCodice,
            name: nomeCodice,
            birthdate: `${giorno}/${mese}/${anno}`,
            gender: sesso,
            comune: comune
        };

    } catch (error) {
        return "ERROR: " + error.message;
    }
}
