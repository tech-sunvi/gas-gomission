const SPREADSHEET_ID = '1MTNQLZMBJCyirap6tKG9dbkHEFiRMl3drtJ6LO3bEIk';
const destinationFolderId = '1Z_JzwMsJ90gguIXmTwAbSLJOW1pc5hgB';
const personalTemplateId = '1ilY1k5UKbnihJ0khIfbFiS-tNcVwJkKT8oYL1ZKSXVA';
const MISSING_DATA = '-----';

// Helper to get Spreadsheet instance efficiently
let _ss = null;
function getSpreadsheet() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

function doGet(e) {
  if (e.parameter && e.parameter.page === 'edit') {
    return HtmlService.createHtmlOutputFromFile('editPersonnel');
  }
  return HtmlService.createHtmlOutputFromFile('startingForm');
}

// ==========================================
// AUTOCOMPLETE FUNCTIONS
// ==========================================

/**
 * Generic helper to search a column in a sheet
 * @param {string} sheetName - Name of the sheet
 * @param {string} searchColName - Header name of the column to search
 * @param {string} query - The search query
 * @param {Array<string>} returnColNames - Optional: Array of other columns to return. If null, returns the search col value.
 */
function searchSheet(sheetName, searchColName, query, returnColNames) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const searchIndex = headers.indexOf(searchColName);
  if (searchIndex === -1) throw new Error(`Column "${searchColName}" not found in ${sheetName}`);

  // Resolve indices for return columns if provided
  const returnIndices = returnColNames ? returnColNames.map(name => headers.indexOf(name)) : null;

  const queryLower = query.toLowerCase();
  const results = [];

  // Start from 1 to skip header
  for (let i = 1; i < data.length; i++) {
    const cellValue = String(data[i][searchIndex]);
    if (cellValue.toLowerCase().includes(queryLower)) {
      if (returnIndices) {
        // Return array of requested columns
        results.push(returnIndices.map(idx => data[i][idx]));
      } else {
        // Return just the value
        results.push(cellValue);
      }
    }
  }
  return results;
}

function getEmployeeNames(hint) {
  // Returns: [EmployeeId, Prénoms, Nom, Fonction]
  // We search in First Name ('Prénoms') OR Last Name ('Nom'). Custom logic needed here.

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Personnel');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxId = headers.indexOf('EmployeeId');
  const idxFirst = headers.indexOf('Prénoms');
  const idxLast = headers.indexOf('Nom');
  const idxFunc = headers.indexOf('Fonction');

  const query = hint.toLowerCase();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const first = String(data[i][idxFirst]).toLowerCase();
    const last = String(data[i][idxLast]).toLowerCase();

    if (first.includes(query) || last.includes(query)) {
      results.push([
        data[i][idxId],
        data[i][idxFirst],
        data[i][idxLast],
        data[i][idxFunc]
      ]);
    }
  }
  return results;
}

function getDestinations(hint) {
  return searchSheet('Destinations', 'Destination', hint);
}

function getTransportMeans(hint) {
  return searchSheet('Transport', 'Moyen de transport', hint);
}

function getBudgets(hint) {
  return searchSheet('Budget', 'Budget', hint);
}

// ==========================================
// PERSONNEL UPDATE LOGIC
// ==========================================

function getPersonnelRecord(employeeId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Personnel');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('EmployeeId');

  if (idIndex === -1) throw new Error('Column "EmployeeId" not found.');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(employeeId)) {
      const record = {};
      headers.forEach((header, index) => {
        // Convert dates to string for easier frontend handling
        let vals = data[i][index];
        if (vals instanceof Date) {
          vals = Utilities.formatDate(vals, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        record[header] = vals;
      });
      return record;
    }
  }
  return null;
}

function updatePersonnelRecord(formObject) {
  try {
    let employeeId = formObject.EmployeeId;

    // Explicit list of allowed columns to edit to allow safety
    const editableColumns = [
      "Nom", "Prénoms", "Civilité", "Fonction", "Date de naissance",
      "Lieu de naissance", "Grade", "Indice", "Matricule",
      "IFU", "Adresse complète", "Telephone", "Email"
    ];

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Personnel');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('EmployeeId');

    // === CREATE MODE ===
    if (!employeeId) {
      // 1. Generate new ID
      // Filter out non-numeric IDs if any, find max
      const ids = data.slice(1).map(r => parseInt(r[idIndex])).filter(val => !isNaN(val));
      const nextId = ids.length > 0 ? Math.max(...ids) + 1 : 1000;
      employeeId = nextId; // Set the new ID

      // 2. Prepare new row (initialized with empty strings)
      const newRow = new Array(headers.length).fill('');

      // Set EmployeeId
      newRow[idIndex] = nextId;

      // Map form fields to new row columns
      editableColumns.forEach(colName => {
        if (formObject.hasOwnProperty(colName)) {
          const colIndex = headers.indexOf(colName);
          if (colIndex !== -1) {
            const val = formObject[colName];
            // Handle date format
            if (colName === "Date de naissance" && val) {
              newRow[colIndex] = new Date(val); // Will be set, but formatting needs range access roughly
            } else {
              newRow[colIndex] = val;
            }
          }
        }
      });

      // Append row
      sheet.appendRow(newRow);

      // Post-Correction for Date Formatting on the last row
      const lastRow = sheet.getLastRow();
      const dobColIndex = headers.indexOf("Date de naissance");
      if (dobColIndex !== -1 && formObject["Date de naissance"]) {
        sheet.getRange(lastRow, dobColIndex + 1).setNumberFormat('yyyy-MM-dd');
      }

      return { success: true, message: "Nouveau dossier créé avec succès (ID: " + nextId + ")" };
    }

    // === UPDATE MODE ===

    // Find row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(employeeId)) {
        rowIndex = i + 1; // 1-based index for Sheet API
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Personnel record not found.");

    editableColumns.forEach(colName => {
      if (formObject.hasOwnProperty(colName)) {
        const colIndex = headers.indexOf(colName);
        if (colIndex !== -1) {
          const cell = sheet.getRange(rowIndex, colIndex + 1);

          if (colName === "Date de naissance") {
            // Handle date specifically to ensure clean formatting in Sheet
            const dateVal = formObject[colName] ? new Date(formObject[colName]) : null;
            if (dateVal) {
              cell.setValue(dateVal).setNumberFormat('yyyy-MM-dd');
            } else {
              cell.clearContent();
            }
          } else {
            // Default handling
            cell.setValue(formObject[colName]);
          }
        }
      }
    });

    return { success: true };
  } catch (e) {
    Logger.log("Update Error: " + e.message);
    return { success: false, message: e.message };
  }
}


// ==========================================
// BACKEND LOGIC
// ==========================================

/**
 * Loads all personnel data into a Map for O(1) access.
 * Reads the sheet ONLY ONCE.
 * @returns {Map<string, Object>} Map of EmployeeId -> Record Object
 */
function getPersonnelMap() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Personnel');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const personnelMap = new Map();
  const idIndex = headers.indexOf('EmployeeId');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const record = {};
    headers.forEach((header, colIndex) => {
      record[header] = row[colIndex];
    });
    // Store by ID (as string to be safe)
    personnelMap.set(String(row[idIndex]), record);
  }

  return personnelMap;
}

function processMissionData(data) {
  try {
    const ss = getSpreadsheet();
    const missionSheet = ss.getSheetByName('Missions');

    // 1. Load all personnel data efficiently
    // This is the key optimization: we stop reading the sheet inside loops
    const personnelMap = getPersonnelMap();

    // 2. Prepare Driver Data
    // Logic: Filter members who have function 'Conducteur...'
    // Since we have the map, we can just check the members provided in data.members

    const driverIds = [];
    data.members.forEach(memberId => {
      const record = personnelMap.get(String(memberId));
      if (record && record['Fonction'] === 'Conducteur de véhicules administratifs') {
        driverIds.push(memberId);
      }
    });

    data.drivers = driverIds; // Store found drivers

    // 3. Save to "Missions" Sheet
    const missionId = `ODM-${Date.now()}`;
    const headers = missionSheet.getRange(1, 1, 1, missionSheet.getLastColumn()).getValues()[0];

    const newRowData = headers.map(header => {
      if (header === 'MissionId') return missionId;
      if (header === 'CreatedAt') {
        const now = new Date();
        return `${now.getFullYear()}-${(now.getMonth() + 1).toString().padStart(2, '0')}-${now.getDate().toString().padStart(2, '0')} ${now.getHours()}:${now.getMinutes()}`;
      }

      // Map data fields to headers
      // Note: data keys usually match headers but lowerCamelCase vs Header Name needs mapping if strictly required.
      // In the original code, it checked `if(data[el])`. Let's assume headers in sheet match keys in data object somewhat or data object keys are used directly?
      // The original code used: if(data[el]).
      // This implies the Sheet Headers MUST match the keys in the `data` object (e.g. 'reference', 'odmType').
      // Let's preserve that logic.

      const value = data[header]; // This relies on Header Name == JSON Key Name
      if (value) {
        return Array.isArray(value) ? value.join(' - ') : value;
      }
      return '';
    });

    // Since original code had specific header mapping logic that might be fragile (it relied on data[headerName]),
    // but the JSON keys are 'reference', 'odmType' etc. and headers might be 'Reference', 'Type' etc.
    // The original code: `if(data[el])`. `el` is the header name.
    // So if Header is 'Reference', data['Reference'] must exist.
    // BUT the frontend sends `reference` (lowercase).
    // Use keys map if needed, or rely on original behavior (which imply headers match keys exactly or keys were added to data object).
    // To be safe and identical to functionality, I will trust the original logic's assumption or Map it if broken.
    // Original: `var newRowData = [missionId] ... headers.forEach(el => ...)`
    // Actually, original code pushed `missionId` first, then looped headers? Use careful reconstruction.
    // Correction: It initialized `newRowData = [missionId]`.
    // Then looped headers. CONSTANT WARNING: If `headers` includes 'MissionId' at index 0, it would double add.
    // Let's stick to a robust approach:

    const rowToAppend = [];
    headers.forEach(header => {
      if (header === 'MissionId') {
        rowToAppend.push(missionId);
      } else if (header === 'CreatedAt') {
        const now = new Date(); // Simple formatting
        rowToAppend.push(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      } else {
        // Try exact match or match with common case issues
        let val = data[header] || data[header.charAt(0).toLowerCase() + header.slice(1)];
        if (val) {
          rowToAppend.push(Array.isArray(val) ? val.join(' - ') : val);
        } else {
          rowToAppend.push('');
        }
      }
    });

    missionSheet.appendRow(rowToAppend);

    // 4. Generate Document
    if (data.odmType === 'individual') {
      generateIndividualDocument(data, personnelMap);
    }

    return true;

  } catch (e) {
    Logger.log('ERROR in processMissionData: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    throw e; // Re-throw to be caught by failure handler
  }
}

/**
 * Generates the Google Doc.
 * @param {Object} data - Form data
 * @param {Map} personnelMap - Cached personnel data (ID -> Record)
 */
function generateIndividualDocument(data, personnelMap) {
  const docName = data.docName ? `Ordre de Mission - ${data.docName}` : `Ordre de Mission - ${data.reference} - ${new Date().toLocaleDateString()}`;

  const templateFile = DriveApp.getFileById(personalTemplateId);
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const newDocFile = templateFile.makeCopy(docName, destinationFolder);
  const doc = DocumentApp.openById(newDocFile.getId());
  const body = doc.getBody();

  // Helper for dates
  const formatDate = (dateStr) => {
    if (!dateStr) return MISSING_DATA;
    return new Date(dateStr).toLocaleDateString('fr-FR', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });
  };

  // Base ODM Data
  const replacements = {
    reference: data.reference || MISSING_DATA,
    destinations: (data.destinations || []).join(', ') || MISSING_DATA,
    dateDepart: formatDate(data.departureDate),
    dateRetour: formatDate(data.returnDate),
    transportMeans: (data.transportMeans || []).join(', ') || MISSING_DATA,
    budgets: (data.budgets || []).join(', ') || MISSING_DATA
  };

  // Driver Info (Primary Driver)
  if (data.drivers && data.drivers.length > 0) {
    const driverRecord = personnelMap.get(String(data.drivers[0]));
    replacements.driver = driverRecord ? `${driverRecord['Nom']} ${driverRecord['Prénoms']}` : MISSING_DATA;
  } else {
    replacements.driver = MISSING_DATA;
  }

  // Process Each Member
  // Note: If template is meant for one person (Individual ODM), usually we generate ONE doc per person? 
  // OR one doc containing all? The original code loops `data.members.forEach` and does `body.replaceText`. 
  // If there are multiple members in "Individual" mode (which sounds contradictory but possible), 
  // replacing '{{nom}}' once will replace it for ALL occurrences. 
  // The original code seems to assume 1 member OR it overwrites placeholders (which only works effectively for 1 member).
  // If 'Individual' ODM implies 1 form per person, logic might be needed to duplicate the template/pages.
  // HOWEVER, based on exact refactoring: I will preserve original behavior: Loop members and replace. 
  // (Warning: If multiple members, later ones might find no placeholders if they were replaced by the first one).

  // Assuming typically 1 member for Individual ODM or the template is designed to handle list?
  // Original code: `data.members.forEach(memberId => { ... body.replaceText ... })`
  // I will strictly follow this logic but use the Map.

  data.members.forEach(memberId => {
    const member = personnelMap.get(String(memberId));
    if (!member) return;

    // Build specific member replacements
    const memberReplacements = {
      ...replacements, // Inherit base replacements
      nom: member['Nom'],
      prenom: member['Prénoms'],
      fullName: `${member['Nom']} ${member['Prénoms']}`,
      civilite: member['Civilité'],
      fonction: member['Fonction'] || MISSING_DATA,
      grade: member['Grade'] || MISSING_DATA,
      indice: member['Indice'] || MISSING_DATA,
      matricule: member['Matricule'] || MISSING_DATA,
      ifu: member['IFU'] || MISSING_DATA,
      adresse: member['Adresse complète'] || MISSING_DATA,
      lieuNaissance: member['Lieu de naissance'] || MISSING_DATA,
      dateNaissance: member['Date de naissance'] ? new Date(member['Date de naissance']).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' }) : MISSING_DATA,
    };

    // Phone formatting
    let phone = member['Telephone'] ? String(member['Telephone']) : '';
    if (phone.includes('+229')) phone = phone.replace('+229', '');
    memberReplacements.phone = phone.replace(/\s/g, '').trim() || MISSING_DATA;

    // Special Mission Object logic
    memberReplacements.missionObject = (member['Fonction'] === 'Conducteur de véhicules administratifs')
      ? `conduire l'équipe chargée de ${data.missionObject}`
      : data.missionObject;

    // Apply replacements
    for (const [key, val] of Object.entries(memberReplacements)) {
      body.replaceText(`{{${key}}}`, String(val));
    }
  });

  doc.saveAndClose();
  return newDocFile.getUrl();
}

// Dummy data just for testing
// mission request
function testSubmission() {
  /*const data = {
    "reference": "6985",
    "odmType": "individual",
    "destinations": [
        "Cotonou"
    ],
    "members": [
        "303",
        "75",
        "468",
        "266"
    ],
    "missionObject": "Some object comme le chant. du livre se détachait de. sans cause, incompréhensible, comme une chose vraiment obscure. Je. vite que je n'avais pas le temps",
    "departureDate": "2025-06-14",
    "returnDate": "2025-06-17",
    "transportMeans": [
        "BN 1263 RB",
        "IPY 6595 RB"
    ],
    "budgets": [],
    "docName": "test objet"
  }*/

  /* const data = {
    "reference": "",
    "docName": "mission sans frais",
    "odmType": "collective",
    "destinations": [
        "Agouna",
        "Zagnanado"
    ],
    "members": [
        "599",
        "594",
        "79"
    ],
    "missionObject": "aller faire quelque chose que je ne connais pas",
    "departureDate": "2025-06-14",
    "returnDate": "2025-06-14",
    "transportMeans": [
        "BN 2813 RB"
    ],
    "budgets": []
  }*/

  const data = {
    "reference": "9484/SRTB/DG/CSAJ",
    "docName": "mission solo",
    "odmType": "individual",
    "destinations": [
      "Natitingou"
    ],
    "members": [
      "594"
    ],
    "missionObject": "Une ballade de santé",
    "departureDate": "2025-06-16",
    "returnDate": "2025-06-29",
    "transportMeans": [
      "avion"
    ],
    "budgets": [
      "Budget organisateur"
    ]
  }

  processMissionData(data);
}

