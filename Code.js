const SPREADSHEET_ID = '1MTNQLZMBJCyirap6tKG9dbkHEFiRMl3drtJ6LO3bEIk';
const destinationFolderId = '1Z_JzwMsJ90gguIXmTwAbSLJOW1pc5hgB';
const personalTemplateId = '1ilY1k5UKbnihJ0khIfbFiS-tNcVwJkKT8oYL1ZKSXVA';
const noteServiceMultipleTemplateId = '1kx1tsVfDj1dBcaSqgRKeJ3wqLHDvp6z6zWu70d912bc';
const noteServiceSingleTemplateId = '1GI_co10kNRzpkJwB1M7Td_lokBc-j-ESQU2UdRpyjs4';
const MISSING_DATA = '-----';
const MISSION_GROUPS_SHEET_NAME = 'MissionGroups';

// Helper to get Spreadsheet instance efficiently
let _ss = null;
function getSpreadsheet() {
  if (!_ss) {
    _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Ensure MissionGroups sheet exists
    if (!_ss.getSheetByName(MISSION_GROUPS_SHEET_NAME)) {
      const sheet = _ss.insertSheet(MISSION_GROUPS_SHEET_NAME);
      sheet.appendRow(['MissionID', 'Vehicle', 'DriverID', 'PassengerIDs']);
    }
  }
  return _ss;
}

function doGet(e) {
  let template;
  let page = e.parameter.page;
  
  if (page === 'edit') {
    template = HtmlService.createTemplateFromFile('editPersonnel');
  } else if (page === 'add_employee') {
    template = HtmlService.createTemplateFromFile('form_employee');
  } else if (page === 'add_vehicle') {
    template = HtmlService.createTemplateFromFile('form_vehicle');
  } else if (page === 'add_destination') {
    template = HtmlService.createTemplateFromFile('form_destination');
  } else {
    // Default to 'new_mission' or root
    template = HtmlService.createTemplateFromFile('startingForm');
  }
  
  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
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

    const rowToAppend = [];
    headers.forEach(header => {
      // Fix: Check for both MissionId and MissionID to be safe
      if (header === 'MissionId' || header === 'MissionID') {
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

    // [New] Save Groups to 'MissionGroups' Sheet
    if (data.groups && data.groups.length > 0) {
        const groupsSheet = ss.getSheetByName(MISSION_GROUPS_SHEET_NAME);
        // Headers: MissionID, Vehicle, DriverID, PassengerIDs
        // Ensure we follow order or dynamically map if needed. For now simple append.
        data.groups.forEach(group => {
           groupsSheet.appendRow([
               missionId,
               group.vehicle || '',
               group.driverId || '',
               (group.passengers || []).join(',')
           ]);
        });
    }

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
  // Determine Mission Object Vowel logic (Global for all members?)
  const vowels = ['a', 'e', 'é', 'è', 'ê', 'i', 'î', 'ï', 'o', 'ô', 'ö', 'u', 'ù', 'û', 'ü', 'y'];
  let globalDriverMissionObjectIntro = '';
  if (data.missionObject && vowels.includes(data.missionObject.charAt(0).toLowerCase())) {
    globalDriverMissionObjectIntro = "conduire l'équipe chargée d'";
  } else {
    globalDriverMissionObjectIntro = "conduire l'équipe chargée de ";
  }

  // 2. Process Groups & Members
  const noteDeServiceData = [];

  // Logic: Iterate groups.
  // data.groups is expected. If missing (legacy usage?), use data.members as specific group?
  // Let's normalize: if no groups but members exist, treat as 1 group.
  let groupsToProcess = data.groups || [];
  if (groupsToProcess.length === 0 && data.members && data.members.length > 0) {
      // Fallback for legacy calls or simple mode if logic still routes here
      groupsToProcess.push({
          vehicle: (data.transportMeans && data.transportMeans[0]) || '',
          driverId: (data.drivers && data.drivers[0]) || '',
          passengers: data.members
      });
  }

  groupsToProcess.forEach(group => {
      // Resolve Driver for this group
      let groupDriverName = MISSING_DATA;
      if (group.driverId) {
          const r = personnelMap.get(String(group.driverId));
          if (r) groupDriverName = `${r['Nom']} ${r['Prénoms']}`;
      } else if (group.driverName) {
           groupDriverName = group.driverName;
      }

      // Iterate passengers in this group
      const passengers = group.passengers || [];
      passengers.forEach(memberId => {
        const member = personnelMap.get(String(memberId));
        if (!member) return;

        // Collect data for Note De Service (Global list)
        noteDeServiceData.push(`- ${member['Civilité']} ${member['Nom']} ${member['Prénoms']}, ${member['Fonction']}`);

        // Build specific member replacements
        const memberReplacements = {
          ...replacements,
          // Overwrite transport/driver with GROUP specific info
          transportMeans: group.vehicle || replacements.transportMeans,
          driver: groupDriverName,
          
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
           memberReplacements.missionObject = `${globalDriverMissionObjectIntro}${data.missionObject}`;
        } else {
           memberReplacements.missionObject = data.missionObject;
        }

        // Creating member list for Note de Service
        // This is tricky: Note de Service needs ALL members from ALL groups.
        // But we are inside the loop.
        // Wait, the logic for "Note de Service" depends on `noteDeServiceData` which is being built HERE.
        // IF we join it here, it will only contain members processed SO FAR.
        // CORRECT LOGIC: We must process ALL members first to build `noteDeServiceData`, THEN generate documents?
        // OR pass a reference?
        // In previous working code, we iterated members, updated `noteDeServiceData`, then joined it.
        // BUT `noteDeServiceData` only had *previous* + *current* member.
        // Is that desired? No, usually Note de Service lists EVERYONE.
        // If the original code worked by accumulating, it means the LAST page had everyone?
        // Or did every page list everyone processed *so far*?
        // Actually, for individual ODM, `membersList` placeholder might strictly refer to the Note De Service page attached at the end.
        // BUT `generateIndividualDocument` makes a TEMP doc for *each* member.
        // Valid question: Does the individual ODM page (page 1) show the list of everyone?
        // Usually not. It shows "Order de Mission" for X.
        // The list is on the "Note de Service".
        // SO: We can defer the Note De Service generation to the end (which we do).
        // BUT: Does `memberReplacements` use `membersList`?
        // Lines 413-418 in original: `memberReplacements.membersList = noteDeServiceData.join('\n');`
        // This implies the individual ODM template *might* use it.
        // If it does, using a partial list is buggy or weird.
        // Ideally, we should pre-calculate the full list.
        
        // REFACTOR: Build full member list first.
        // Since we are iterating anyway, let's keep the flow but maybe pre-calculate `noteDeServiceData` if possible?
        // Or just accept the accumulation behavior (it works for the final Note de Service, but for individual pages if they show it, it's partial).
        // Given the prompt "we find out that some missions have so many participants...", the individual page likely doesn't list everyone.
        // It's the Note de Service that lists everyone.
        // I will stick to the accumulation logic but ideally, I should fix it. 
        // Let's stick to accumulation to minimize regression risk on behavior I haven't fully inspected in the template.
        // memberReplacements.membersList = noteDeServiceData.join('\n');
        
        // UPDATE: I will NOT change the logic of accumulation *inside* the loop for now,
        // but I will ensure `noteDeServiceData` aggregates across GROUPS.
      }); 
  });
  
  // NOW iterate again to generate docs?
  // No, we need to generate docs *as* we iterate or after.
  // If we want `membersList` to be complete on all pages (if used), we must pre-process.
  // Let's do 2 passes. Pass 1: Collect all data. Pass 2: Generate.
  
  // Pass 1: Build comprehensive list
  const allMembersForNote = [];
  groupsToProcess.forEach(group => {
       const groupMembers = [...(group.passengers || [])];
       // Add driver if present and distinct
       if (group.driverId && !groupMembers.includes(group.driverId)) {
           groupMembers.unshift(group.driverId);
       }
       
       groupMembers.forEach(pId => {
           const m = personnelMap.get(String(pId));
           if(m) allMembersForNote.push(`- ${m['Civilité']} ${m['Nom']} ${m['Prénoms']}, ${m['Fonction']}`);
       });
  });
  const fullMembersListString = allMembersForNote.join('\n');

  // Pass 2: Generate
  groupsToProcess.forEach(group => {
      // Group Driver
      let groupDriverName = MISSING_DATA;
      if (group.driverId) {
          const r = personnelMap.get(String(group.driverId));
          if (r) groupDriverName = `${r['Nom']} ${r['Prénoms']}`;
      } else if (group.driverName) {
           groupDriverName = group.driverName;
      }
      
      const passengers = group.passengers || [];
      // Combine driver and passengers for ODM generation
      // We want to generate an ODM for the driver too.
      const membersToGenerate = [...passengers];
      if (group.driverId && !membersToGenerate.includes(group.driverId)) {
          membersToGenerate.unshift(group.driverId);
      }

      membersToGenerate.forEach(memberId => {
        const member = personnelMap.get(String(memberId));
        if (!member) return;
        
        const memberReplacements = {
          ...replacements,
          transportMeans: group.vehicle || replacements.transportMeans,
          driver: groupDriverName,
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
          charge: member['Civilité'] === 'Monsieur' ? 'chargé' : 'chargée',
          membersList: fullMembersListString, // Use full list
          phone: (member['Telephone'] ? String(member['Telephone']).replace('+229', '').replace(/\s/g, '').trim() : MISSING_DATA)
        };
        
        if ((member['Fonction'] && member['Fonction'].toLowerCase() === 'conducteur de véhicules administratifs') || String(memberId) === String(group.driverId)) {
           memberReplacements.missionObject = `${globalDriverMissionObjectIntro}${data.missionObject}`;
        } else {
           memberReplacements.missionObject = data.missionObject;
        }

        // Generate Temp Doc
        const tempDocFile = DriveApp.getFileById(personalTemplateId).makeCopy();
        const tempDoc = DocumentApp.openById(tempDocFile.getId());
        const tempBody = tempDoc.getBody();

        for (const [key, val] of Object.entries(memberReplacements)) {
          tempBody.replaceText(`{{${key}}}`, String(val));
        }
        tempDoc.saveAndClose();

        // Append to Final
        const tempDocFilled = DocumentApp.openById(tempDocFile.getId());
        const tempDocFilledBody = tempDocFilled.getBody();
        const totalElements = tempDocFilledBody.getNumChildren();

        for (let i = 0; i < totalElements; i++) {
             const element = tempDocFilledBody.getChild(i).copy();
             appendElement(finalODMDocBody, element);
        }
        
        // Formatting
        const paragraphs = finalODMDocBody.getParagraphs();
        paragraphs.forEach(function (paragraph) {
            const text = paragraph.getText().trim();
            if (text.toUpperCase() === 'ORDRE DE MISSION') {
                paragraph.setFontFamily('Calibri').setFontSize(24).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            } else {
                paragraph.setFontFamily('Calibri').setFontSize(13).setSpacingAfter(5);
            }
        });

        finalODMDocBody.appendPageBreak();
        tempDocFile.setTrashed(true);
      });
  });

  // 5. Generate Note de Service
  // Prepare data for note de service
  // Original logic: Remove missionObject to prevent driverIntro logic carrying over?
  // And re-add original mission object.
  const npsData = {
      ...replacements,
      missionObject: data.missionObject,
      membersList: fullMembersListString
  };

  const noteServiceDocId = generateNoteDeService(npsData, allMembersForNote);
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


// ==========================================
// RESOURCE MANAGEMENT FUNCTIONS
// ==========================================

function addEmployee(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Personnel');
  // Generate EmployeeId: Max existing ID + 1
  // ID is in column 1. Get values, flatten, convert to number, find max.
  // Warning: getRange(2, 1, lastRow-1) might fail if lastRow < 2 (empty sheet).
  let newId = 1;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(Number).filter(n => !isNaN(n));
    if (existingIds.length > 0) {
      newId = Math.max(...existingIds) + 1;
    }
  }

  // PRODUCT_SPEC Order:
  // EmployeeId, Nom, Prénoms, Civilité, Fonction, Date de naissance, Lieu de naissance, Grade, Indice, Matricule, IFU, Adresse complète, Telephone, Sexe, Email
  const row = [
    newId,
    data.nom,
    data.prenoms,
    data.civilite,
    data.fonction,
    data.dateNaissance ? new Date(data.dateNaissance) : '', // Store as date object or string? Sheet prefers Date object for date columns usually
    data.lieuNaissance,
    data.grade,
    data.indice,
    data.matricule,
    data.ifu,
    data.adresse,
    data.telephone,
    data.sexe,
    data.email
  ];
  sheet.appendRow(row);
  return newId;
}

function addVehicle(vehicleName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Transport');
  sheet.appendRow([vehicleName]);
  return true;
}

function addDestination(destinationName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Destination');
  sheet.appendRow([destinationName]);
  return true;
}
