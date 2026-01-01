const SPREADSHEET_ID = '1MTNQLZMBJCyirap6tKG9dbkHEFiRMl3drtJ6LO3bEIk';
const destinationFolderId = '1Z_JzwMsJ90gguIXmTwAbSLJOW1pc5hgB';
const personalTemplateId = '1ilY1k5UKbnihJ0khIfbFiS-tNcVwJkKT8oYL1ZKSXVA';
const noteServiceMultipleTemplateId = '1kx1tsVfDj1dBcaSqgRKeJ3wqLHDvp6z6zWu70d912bc';
const noteServiceSingleTemplateId = '1GI_co10kNRzpkJwB1M7Td_lokBc-j-ESQU2UdRpyjs4';
const MISSING_DATA = '-----';

// Helper to get Spreadsheet instance efficiently
let _ss = null;
function getSpreadsheet() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

function doGet(e) {
  if (e.parameter && e.parameter.page === 'edit') {
    return HtmlService.createHtmlOutputFromFile('editPersonnel')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('startingForm')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
    const employeeId = formObject.EmployeeId;
    if (!employeeId) throw new Error("EmployeeId is missing.");

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Personnel');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('EmployeeId');

    // Find row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(employeeId)) {
        rowIndex = i + 1; // 1-based index for Sheet API
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Personnel record not found.");

    // Update columns
    // We iterate over the received formObject keys and if they match a header, we update that cell.
    // Explicit list of allowed columns to edit to allow safety
    const editableColumns = [
      "Nom", "Prénoms", "Civilité", "Fonction", "Date de naissance",
      "Lieu de naissance", "Grade", "Indice", "Matricule",
      "IFU", "Adresse complète", "Telephone", "Email"
    ];

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
  const docName = data.docName ? `Ordre de Mission - ${new Date().toLocaleDateString()} ${data.docName}` : `Ordre de Mission - ${data.reference} - ${new Date().toLocaleDateString()}`;

  // Helper for dates
  const formatDate = (dateStr) => {
    if (!dateStr) return MISSING_DATA;
    return new Date(dateStr).toLocaleDateString('fr-FR', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
    });
  };

  // 1. Create Final Empty Document
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  // Create doc in root first (default behavior of create), then move? Or just create.
  // DocumentApp.create() creates in root. We need to move it.
  const finalODMDoc = DocumentApp.create(docName);
  const finalODMDocFile = DriveApp.getFileById(finalODMDoc.getId());
  finalODMDocFile.moveTo(destinationFolder);
  
  const finalODMDocBody = finalODMDoc.getBody();

  // Base ODM Data
  const replacements = {
    reference: data.reference || MISSING_DATA,
    destinations: (data.destinations || []).join(', ') || MISSING_DATA,
    dateDepart: formatDate(data.departureDate),
    dateRetour: formatDate(data.returnDate),
    transportMeans: (data.transportMeans || []).join(', ') || MISSING_DATA,
    budgets: (data.budgets || []).join(', ') || 'Budget SRTB',
    // Add missing fields from original logic
    datesString: data.departureDate === data.returnDate ? `le ${formatDate(data.departureDate)}` : `du ${formatDate(data.departureDate)} au ${formatDate(data.returnDate)}`
  };

  // Driver Info (Primary Driver)
  if (data.drivers && data.drivers.length > 0) {
    const driverRecord = personnelMap.get(String(data.drivers[0]));
    replacements.driver = driverRecord ? `${driverRecord['Nom']} ${driverRecord['Prénoms']}` : MISSING_DATA;
  } else {
    replacements.driver = MISSING_DATA;
  }

  // Determine Mission Object Vowel logic (Global for all members?)
  // Original logic checked data.missionObject first char for 'conduire l'équipe...' prefix
  // but applied it individually per member if they are 'Conducteur...'
  const vowels = ['a', 'e', 'é', 'è', 'ê', 'i', 'î', 'ï', 'o', 'ô', 'ö', 'u', 'ù', 'û', 'ü', 'y'];
  let driverMissionObjectIntro = '';
  if (data.missionObject && vowels.includes(data.missionObject.charAt(0).toLowerCase())) {
    driverMissionObjectIntro = "conduire l'équipe chargée d'";
  } else {
    driverMissionObjectIntro = "conduire l'équipe chargée de ";
  }

  // 2. Process Members
  const noteDeServiceData = [];

  data.members.forEach(memberId => {
    const member = personnelMap.get(String(memberId));
    if (!member) return;

    // Collect data for Note De Service
    noteDeServiceData.push(`- ${member['Civilité']} ${member['Nom']} ${member['Prénoms']}, ${member['Fonction']}`);

    // Build specific member replacements
    const memberReplacements = {
      ...replacements,
      nom: member['Nom'],
      prenom: member['Prénoms'],
      fullName: `${member['Nom']} ${member['Prénoms']}`,
      civilite: member['Civilité'],
      fonction: member['Fonction'] || MISSING_DATA,
      grade: member['Grade'] || MISSING_DATA,
      indice: member['Indice'] || MISSING_DATA,
      matricule: member['Matricule'] || MISSING_DATA,
      ifu: member['IFU'] || MISSING_DATA,
      adresse: (member['Adresse complète'] || MISSING_DATA).replace(/;\s*$/, ''),
      lieuNaissance: member['Lieu de naissance'] || MISSING_DATA,
      dateNaissance: member['Date de naissance'] ? new Date(member['Date de naissance']).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' }) : MISSING_DATA,
      charge: member['Civilité'] === 'Monsieur' ? 'chargé' : 'chargée'
    };

    // Phone formatting
    let phone = member['Telephone'] ? String(member['Telephone']) : '';
    if (phone.includes('+229')) phone = phone.replace('+229', '');
    memberReplacements.phone = phone.replace(/\s/g, '').trim() || MISSING_DATA;

    // Mission Object Logic
    if (member['Fonction'] && member['Fonction'].toLowerCase() === 'conducteur de véhicules administratifs') {
       memberReplacements.missionObject = `${driverMissionObjectIntro}${data.missionObject}`;
    } else {
       memberReplacements.missionObject = data.missionObject;
    }

    // Creating member list for Note de Service
    // The line break is added to each member in the array to ensure we have each element on a new line
    // The dash previously added in the loop (`- ${member['Civilité']} ${member['Nom']} ${member['Prénoms']}, ${member['Fonction']}`)
    // is put to simulate bullet list in the final document
    // as bullet list is difficult to create programmatically in Google Docs
    memberReplacements.membersList = noteDeServiceData.join('\n');

    // 3. Create Temp Doc for this member
    const tempDocFile = DriveApp.getFileById(personalTemplateId).makeCopy();
    const tempDoc = DocumentApp.openById(tempDocFile.getId());
    const tempBody = tempDoc.getBody();

    // Replace Text
    for (const [key, val] of Object.entries(memberReplacements)) {
      tempBody.replaceText(`{{${key}}}`, String(val));
    }
    tempDoc.saveAndClose();

    // 4. Append Temp Doc content to Final Doc
    const tempDocFilled = DocumentApp.openById(tempDocFile.getId());
    const tempDocFilledBody = tempDocFilled.getBody();
    const totalElements = tempDocFilledBody.getNumChildren();

    for (let i = 0; i < totalElements; i++) {
        const element = tempDocFilledBody.getChild(i).copy();
        appendElement(finalODMDocBody, element);
    }

    // Formatting fix ported from backup
    const paragraphs = finalODMDocBody.getParagraphs();
    // We only need to format the newly added paragraphs? No, iterating all is safer/easier as per original
    paragraphs.forEach(function (paragraph) {
        const text = paragraph.getText().trim();
        if (text.toUpperCase() === 'ORDRE DE MISSION') {
            paragraph.setFontFamily('Calibri')
                .setFontSize(24)
                .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        } else {
            // Only apply to defaults? Careful not to overwrite if template had multiple fonts.
            // Original code applied this indiscriminately.
            paragraph.setFontFamily('Calibri')
                .setFontSize(13)
                .setSpacingAfter(5);
        }
    });

    finalODMDocBody.appendPageBreak();
    tempDocFile.setTrashed(true);
  });

  // 5. Generate Note de Service
  // Prepare data for note de service
  // Original logic: Remove missionObject to prevent driverIntro logic carrying over?
  // And re-add original mission object.
  const npsData = {
      ...replacements,
      missionObject: data.missionObject,
      membersList: noteDeServiceData.join('\n')
  };

  const noteServiceDocId = generateNoteDeService(npsData, noteDeServiceData);
  const tempNoteServiceDoc = DocumentApp.openById(noteServiceDocId);
  const tempNoteServiceDocBody = tempNoteServiceDoc.getBody();
  const totalElements = tempNoteServiceDocBody.getNumChildren();

  for (let i = 0; i < totalElements; i++) {
      const element = tempNoteServiceDocBody.getChild(i).copy();
      appendElement(finalODMDocBody, element);
  }

  DriveApp.getFileById(noteServiceDocId).setTrashed(true);
  finalODMDoc.saveAndClose();
  
  return finalODMDocFile.getUrl();
}

/**
 * Helper to append different element types to a document body.
 */
function appendElement(docBody, element) {
  const type = element.getType();
  switch (type) {
    case DocumentApp.ElementType.PARAGRAPH:
      docBody.appendParagraph(element); // .copy() not needed if element is already a copy? logic says element.copy() in loop.
      break;
    case DocumentApp.ElementType.TABLE:
      docBody.appendTable(element);
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      docBody.appendListItem(element);
      break;
    case DocumentApp.ElementType.INLINE_IMAGE:
      docBody.appendImage(element);
      break;
    default:
      Logger.log('Unsupported element type: ' + type);
  }
}

/**
 * Generates the Note de Service temp doc.
 */
function generateNoteDeService(odmData, membersDataArray) {
    const noteServiceTemplateId = membersDataArray.length > 1 ? noteServiceMultipleTemplateId : noteServiceSingleTemplateId;
    const tempFile = DriveApp.getFileById(noteServiceTemplateId).makeCopy();
    const tempDoc = DocumentApp.openById(tempFile.getId());
    const body = tempDoc.getBody();

    for (const [key, val] of Object.entries(odmData)) {
        body.replaceText(`{{${key}}}`, String(val));
    }
    tempDoc.saveAndClose();
    return tempFile.getId();
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

