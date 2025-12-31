const SPREADSHEET_ID = '1MTNQLZMBJCyirap6tKG9dbkHEFiRMl3drtJ6LO3bEIk';
const rootFolderId = DriveApp.getRootFolder().getId();
const destinationFolderId = '1Z_JzwMsJ90gguIXmTwAbSLJOW1pc5hgB';
const odmIndividualTemplateId = '1ilY1k5UKbnihJ0khIfbFiS-tNcVwJkKT8oYL1ZKSXVA';

const noteServiceMultipleTemplateId = '1kx1tsVfDj1dBcaSqgRKeJ3wqLHDvp6z6zWu70d912bc';
const noteServiceSingleTemplateId = '1GI_co10kNRzpkJwB1M7Td_lokBc-j-ESQU2UdRpyjs4';

const odmSFMMultipleTemplateId = '1PPnzQ3BIZMOUuDn4w0R1STPbLNd1uHxysG6wHacxVlY';
const odmSFMSingleTemplateId = '1saRnddr7fLxwdeJBQ87ILu6z2dvaEwL-1pRzDO2_Va8';

// Get today's date
const today = new Date();
const day = today.getDate().toString().padStart(2, '0');
const month = (today.getMonth() + 1).toString().padStart(2, '0');
const year = today.getFullYear();
const hours = today.getHours();
const minutes = today.getMinutes() + 1;
const seconds = today.getSeconds();

const formattedDate = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

// Missing data placeholder
const missingDataPlaceholder = '-----';

var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

function doGet() {
    return HtmlService.createHtmlOutputFromFile('startingForm');
}

function getEmployeeNames(hint) {
    var hintLower = hint.toLowerCase();
    var employeesSheet = ss.getSheetByName('Personnel');

    var data = employeesSheet.getDataRange().getValues();

    // Assuming the first row contains headers
    var headers = data[0];
    var firstNameIndex = headers.indexOf('Prénoms');
    var lastNameIndex = headers.indexOf('Nom');

    var employeeIdIndex = headers.indexOf('EmployeeId');
    var fonctionIndex = headers.indexOf('Fonction');

    if (firstNameIndex == -1 || lastNameIndex == -1) {
        throw new Error('FirstName or LastName column not found.');
    }

    var matchingRows = [];

    for (var i = 1; i < data.length; i++) {
        var firstName = data[i][firstNameIndex].toLowerCase();
        var lastName = data[i][lastNameIndex].toLowerCase();

        if (firstName.includes(hintLower) || lastName.includes(hintLower)) {
            // matchingRows.push(data[i]);
            matchingRows.push([data[i][employeeIdIndex], data[i][firstNameIndex], data[i][lastNameIndex], data[i][fonctionIndex]]);
        }
    }
    return matchingRows;
}

function getDestinations(hint) {
    var destinationsSheet = ss.getSheetByName('Destinations');

    var data = destinationsSheet.getDataRange().getValues();

    // Assuming the first row contains headers
    var headers = data[0];
    var desinationNameIndex = headers.indexOf('Destination');


    if (desinationNameIndex == -1) {
        throw new Error('La colonne Destination n\'est pas trouvée.');
    }

    // Flatten the 2D array to a 1D array
    var destinations = data.map(function (row) {
        return row[0];
    });

    var matchinDestinations = [];

    destinations.forEach((destination) => {
        if (destination.toLowerCase().includes(hint.toLowerCase()))
            matchinDestinations.push(destination);
    });

    return matchinDestinations;
}

function getTransportMeans(hint) {
    var transportMeansSheet = ss.getSheetByName('Transport');

    var data = transportMeansSheet.getDataRange().getValues();

    // Assuming the first row contains headers
    var headers = data[0];
    var transportMeansNameIndex = headers.indexOf('Moyen de transport');

    if (transportMeansNameIndex == -1) {
        throw new Error('La colonne Moyen de transport n\'est pas trouvée.');
    }

    // Flatten the 2D array to a 1D array
    var transportMeans = data.map(function (row) {
        return row[0];
    });

    var matchingTransportMeans = [];

    transportMeans.forEach((transportMean) => {
        if (transportMean.toLowerCase().includes(hint.toLowerCase()))
            matchingTransportMeans.push(transportMean);
    });

    return matchingTransportMeans;
}

function getBudgets(hint) {
    var budgetSheet = ss.getSheetByName('Budget');

    var data = budgetSheet.getDataRange().getValues();

    // Assuming the first row contains headers
    var headers = data[0];
    var budgetNameIndex = headers.indexOf('Budget');

    if (budgetNameIndex == -1) {
        throw new Error('La colonne Budget n\'est pas trouvée.');
    }

    // Flatten the 2D array to a 1D array
    var budgets = data.map(function (row) {
        return row[0];
    });

    var matchingBudgets = [];

    budgets.forEach((budget) => {
        if (budget.toLowerCase().includes(hint.toLowerCase()))
            matchingBudgets.push(budget);
    });

    return matchingBudgets;
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

function getFilteredRecords(employeeIds, columnName, desiredValue) {
    const sheet = ss.getSheetByName('Personnel');
    const data = sheet.getDataRange().getValues();

    const headers = data[0];
    const employeeIdIndex = headers.indexOf('EmployeeId');
    const targetColumnIndex = headers.indexOf(columnName);

    if (employeeIdIndex === -1 || targetColumnIndex === -1) {
        Logger.log(`Column "EmployeeId" or "${columnName}" not found.`);
        return [];
    }

    const matchingRecords = [];

    for (let i = 1; i < data.length; i++) { // Start from row 1 to skip headers
        let row = data[i];
        let currentEmployeeId = row[employeeIdIndex];
        let targetValue = row[targetColumnIndex];
        Logger.log('Current EmployeeId: ' + currentEmployeeId + ' Target Value: ' + targetValue);

        if (employeeIds.includes(currentEmployeeId.toString()) && targetValue.toLowerCase() === desiredValue.toLowerCase()) {
            matchingRecords.push(row);
        }
    }

    return matchingRecords;
}

// Get an employee record by its id an returns data as an object
function getEmployeeRecordById(employeeId) {
    var sheet = ss.getSheetByName('Personnel');

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var employeeIdIndex = headers.indexOf('EmployeeId');

    if (employeeIdIndex === -1) {
        throw new Error('Column "EmployeeId" not found.');
    }

    for (var i = 1; i < data.length; i++) { // Skip header row
        if (data[i][employeeIdIndex] == employeeId) { // Non-strict match to avoid type issues
            var record = {};
            headers.forEach(function (header, index) {
                record[header] = data[i][index];
            });
            return record; // Return the matched employee as an object
        }
    }

    return null; // No match found
}

function processMissionData(data) {
    // Add missing data
    data.budgets = data.budgets || ['Budget SRTB'];
    data.reference = data.reference || 'Sans référence';

    const missionId = `ODM-${Date.now()}`;

    // Mission sheet headers

    var missionSheet = ss.getSheetByName('Missions');
    var headers = missionSheet.getRange(1, 1, 1, missionSheet.getLastColumn()).getValues()[0];

    var newRowData = [missionId];

    // Get the driver from the members
    // 1. From members ids, get the "Personnel" sheet records matching the ids
    const drivers = getFilteredRecords(data.members, 'Fonction', 'Conducteur de véhicules administratifs');

    var driversIds = [];

    if (drivers.length > 0) {
        drivers.forEach(driver => {
            driversIds.push(driver[0]);
        });
    }

    data.drivers = driversIds;

    // Logger.log('Final data' + JSON.stringify(data))

    headers.forEach(el => {
        if (data[el]) {
            if (typeof data[el] === "object") {
                newRowData.push(data[el].join(' - '))
            }
            else {
                newRowData.push(data[el]);
            }
        }
    });

    // Find the first empty row
    var lastRow = missionSheet.getLastRow();
    var nextRow = lastRow + 1;

    // Get the range where the row will be added
    var range = missionSheet.getRange(nextRow, 1, 1, newRowData.length);

    // Set the values in the range
    if (range.setValues([newRowData])) {
        // Update CreatedAt cell
        const createdDateCell = missionSheet.getRange(nextRow, headers.indexOf('CreatedAt') + 1, 1, 1);
        createdDateCell.setValue(`${year}-${month}-${day} ${hours}:${minutes}:${seconds}`);

        if (data.odmType === 'individual') {
            generateCombinedIndividualOrdreDeMission(data);
        } else {
            generateSFMOrdreDeMission(data);
        }
        return true;
    }
    else {
        return false
    }
}

// ===============================================
// Generate SFM ODM document
// ===============================================

function generateSFMOrdreDeMission(data) {

    // Create the final ODM document
    const destination = data.destinations[0];
    const docName = data.docName || 'sans nom';
    const docTitle = `ODM ${formattedDate} ${destination} ${docName}`;

    // Make a copy of the template
    var templateFile = DriveApp.getFileById(data.members.length > 1 ? odmSFMMultipleTemplateId : odmSFMSingleTemplateId);
    var newDoc = templateFile.makeCopy(docTitle.split(' ').join('-'), DriveApp.getFolderById(rootFolderId));

    var doc = DocumentApp.openById(newDoc.getId());
    var body = doc.getBody();

    // Build ODM shared data

    const dateDepartString = new Date(data.departureDate).toLocaleDateString('fr-FR', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    })

    const dateRetourString = new Date(data.returnDate).toLocaleDateString('fr-FR', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    })

    const odmData = {
        reference: data.reference,
        destinations: Array.isArray(data.destinations) ? data.destinations.join(', ') : missingDataPlaceholder,
        missionObject: data.missionObject,
        dateDepart: dateDepartString,
        dateRetour: data.departureDate === data.returnDate ? 'même jour' : dateRetourString,
        datesString: data.departureDate === data.returnDate ? `le ${dateDepartString}` : `du ${dateDepartString} au ${dateRetourString}`,
        transportMeans: Array.isArray(data.transportMeans) ? data.transportMeans.join(', ') : missingDataPlaceholder,
        budgets: data.budgets ? Array.isArray(data.budgets) ? data.budgets.join(', ') : 'Budget SRTB' : 'Budget SRTB',
    }

    const membersList = [];

    if (data.members.length > 1) {
        data.members.forEach(memberId => {
            const memberData = getEmployeeRecordById(memberId);
            membersList.push(`- ${memberData['Civilité']} ${memberData['Nom']} ${memberData['Prénoms']}, ${memberData['Fonction']}`);
        });
    } else {
        const memberData = getEmployeeRecordById(data.members[0]);
        odmData.nom = memberData['Nom'];
        odmData.prenom = memberData['Prénoms'];
        odmData.fullName = `${memberData['Nom']} ${memberData['Prénoms']}`;
        odmData.civilite = memberData['Civilité'];
        odmData.fonction = memberData['Fonction'] || missingDataPlaceholder;
        odmData.charge = memberData['Civilité'] === 'Monsieur' ? 'chargé' : 'chargée';
    }

    odmData.membersList = membersList.join('\n');

    // Replace each placeholder
    for (var key in odmData) {
        if (odmData.hasOwnProperty(key)) {
            body.replaceText('{{' + key + '}}', odmData[key]);
        }
    }

    doc.saveAndClose();

    Logger.log('Document created: ' + newDoc.getUrl());
}

function generateCombinedIndividualOrdreDeMission(data) {
    // Create the final ODM document
    const destination = data.destinations[0];
    const docName = data.docName || 'sans nom';
    const docTitle = `ODM ${formattedDate} ${destination} ${docName}`;

    const finalODMDoc = DocumentApp.create(docTitle.split(' ').join('-'));
    const finalODMDocBody = finalODMDoc.getBody();

    // Array for Note de service data

    const noteDeServiceData = []

    // Build ODM shared data

    const dateDepartString = new Date(data.departureDate).toLocaleDateString('fr-FR', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    })

    const dateRetourString = new Date(data.returnDate).toLocaleDateString('fr-FR', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    })

    let budgetString;

    if (data.budgets && data.budgets.length > 0) {
        budgetString = Array.isArray(data.budgets) ? data.budgets.join(', ') : 'Budget SRTB';
    } else {
        budgetString = 'Budget SRTB';
    }

    const odmData = {
        reference: data.reference,
        destinations: Array.isArray(data.destinations) ? data.destinations.join(', ') : missingDataPlaceholder,
        dateDepart: dateDepartString,
        dateRetour: data.departureDate === data.returnDate ? 'même jour' : dateRetourString,
        datesString: data.departureDate === data.returnDate ? `le ${dateDepartString}` : `du ${dateDepartString} au ${dateRetourString}`,
        transportMeans: Array.isArray(data.transportMeans) ? data.transportMeans.join(', ') : missingDataPlaceholder,
        budgets: budgetString,
    }

    // Driver data
    const driverData = getEmployeeRecordById(data.drivers[0]);
    odmData.driver = driverData ? `${driverData['Nom']} ${driverData['Prénoms']}` : missingDataPlaceholder;

    // Build ODM data for each member
    data.members.forEach((memberId, index) => {
        const memberData = getEmployeeRecordById(memberId);

        // Fill note de service data array
        noteDeServiceData.push(`- ${memberData['Civilité']} ${memberData['Nom']} ${memberData['Prénoms']}, ${memberData['Fonction']}`)

        let phoneNumber;
        if (memberData['Telephone']) {
            if (memberData['Telephone'].includes('+229')) {
                phoneNumber = memberData['Telephone'].substring(4);
            }
            else {
                phoneNumber = memberData['Telephone'].substring(3);
            }
            phoneNumber = phoneNumber.replace(/\s/g, '');
        } else {
            phoneNumber = missingDataPlaceholder;
        }

        odmData.nom = memberData['Nom'];
        odmData.prenom = memberData['Prénoms'];
        odmData.fullName = `${memberData['Nom']} ${memberData['Prénoms']}`;
        odmData.civilite = memberData['Civilité'];
        odmData.fonction = memberData['Fonction'] || missingDataPlaceholder;
        odmData.grade = memberData['Grade'] || missingDataPlaceholder;
        odmData.indice = memberData['Indice'] || missingDataPlaceholder;
        odmData.matricule = memberData['Matricule'] || missingDataPlaceholder;
        odmData.ifu = memberData['IFU'] || missingDataPlaceholder;
        odmData.adresse = (memberData['Adresse complète'] || missingDataPlaceholder).replace(/;\s*$/, '');
        odmData.phone = phoneNumber;
        odmData.dateNaissance = new Date(memberData['Date de naissance']).toLocaleDateString('fr-FR', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        }) || missingDataPlaceholder;
        odmData.lieuNaissance = memberData['Lieu de naissance'] || missingDataPlaceholder;

        const vowels = ['a', 'e', 'é', 'è', 'ê', 'i', 'î', 'ï', 'o', 'ô', 'ö', 'u', 'ù', 'û', 'ü', 'y'];

        let driverMissionObjectIntro = '';

        if (vowels.includes(data.missionObject.charAt(0).toLowerCase())) {
            driverMissionObjectIntro = 'conduire l\'équipe chargée d\'';
        }
        else {
            driverMissionObjectIntro = 'conduire l\'équipe chargée de ';
        }
        // Mission object for Conducteur de véhicules administratifs
        odmData.missionObject = memberData.Fonction.toLowerCase() === 'conducteur de véhicules administratifs' ? `${driverMissionObjectIntro}${data.missionObject}` : data.missionObject;
        odmData.charge = memberData.Civilité === 'Monsieur' ? 'chargé' : 'chargée';

        odmData.membersList = noteDeServiceData.join('\n');

        // Logger.log('ODM Data: ' + JSON.stringify(odmData));

        // Make a copy of the template

        const tempDocId = DriveApp.getFileById(odmIndividualTemplateId).makeCopy().getId();
        const tempDoc = DocumentApp.openById(tempDocId);
        const body = tempDoc.getBody();

        // Replace each placeholder in template copy and save
        for (var key in odmData) {
            if (odmData.hasOwnProperty(key)) {
                body.replaceText('{{' + key + '}}', odmData[key]);
            }
        }

        tempDoc.saveAndClose();

        // Open the filled document again to copy its elements
        const tempDocFilled = DocumentApp.openById(tempDocId);
        const tempDocFilledBody = tempDocFilled.getBody();
        const totalElements = tempDocFilledBody.getNumChildren();

        // Copy each element (preserves formatting)
        for (let i = 0; i < totalElements; i++) {
            const element = tempDocFilledBody.getChild(i).copy();

            appendElement(finalODMDocBody, element)
        }

        var paragraphs = finalODMDocBody.getParagraphs();
        paragraphs.forEach(function (paragraph) {
            var text = paragraph.getText().trim();

            if (text.toUpperCase() === 'ORDRE DE MISSION') {
                // Apply specific formatting to the title
                paragraph.setFontFamily('Calibri')
                    .setFontSize(24)
                    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            } else {
                // Apply general formatting to all other paragraphs
                paragraph.setFontFamily('Calibri')
                    .setFontSize(13)
                    .setSpacingAfter(5);
            }
        });

        finalODMDocBody.appendPageBreak();

        // Optionally delete temporary document
        DriveApp.getFileById(tempDocId).setTrashed(true);
    });

    // Prepare data for note de service
    // 1. Remove missionObject from odmData for potential driverMissionObjectIntro
    delete odmData.missionObject;

    // 2. Add original missionObject to odmDataForNoteService
    const odmDataForNoteService = {
        ...odmData,
        missionObject: data.missionObject
    }

    // Append Note de service
    const noteServiceDocId = generateNoteDeService(odmDataForNoteService, noteDeServiceData)

    // Open the filled Note de service document again to copy its elements
    const tempNoteServiceDoc = DocumentApp.openById(noteServiceDocId);
    const tempNoteServiceDocBody = tempNoteServiceDoc.getBody();
    const totalElements = tempNoteServiceDocBody.getNumChildren();

    // Copy all elements from Note de service document (preserves formatting)
    for (let i = 0; i < totalElements; i++) {
        const element = tempNoteServiceDocBody.getChild(i).copy();

        appendElement(finalODMDocBody, element)
    }

    // Delete Note de service document
    DriveApp.getFileById(noteServiceDocId).setTrashed(true);

    Logger.log('Generated Document URL: ' + finalODMDoc.getUrl());
}

// ===============================================
// Note de service functions
// Generate note de service for multiple members
// Generate note de service for single member
// ===============================================

function generateNoteDeService(odmData, membersDataArray) {

    const noteServiceTemplateId = membersDataArray.length > 1 ? noteServiceMultipleTemplateId : noteServiceSingleTemplateId;
    // var noteServiceDoc = DocumentApp.openById(noteServiceTemplateId);
    // var body = noteServiceDoc.getBody();

    const tempNoteServiceDocId = DriveApp.getFileById(noteServiceTemplateId).makeCopy().getId();
    const tempNoteServiceDoc = DocumentApp.openById(tempNoteServiceDocId);
    const tempNoteServiceDocBody = tempNoteServiceDoc.getBody();

    // Replace each placeholder in template copy and save
    for (var key in odmData) {
        if (odmData.hasOwnProperty(key)) {
            tempNoteServiceDocBody.replaceText('{{' + key + '}}', odmData[key]);
        }
    }

    tempNoteServiceDoc.saveAndClose();

    // Return id of the filled note de service document
    return tempNoteServiceDoc.getId();
}

function appendElement(doc, element) {
    const type = element.getType();

    switch (type) {
        case DocumentApp.ElementType.PARAGRAPH:
            doc.appendParagraph(element.copy());
            break;
        case DocumentApp.ElementType.TABLE:
            doc.appendTable(element.copy());
            break;
        case DocumentApp.ElementType.LIST_ITEM:
            doc.appendListItem(element.copy());
            break;
        case DocumentApp.ElementType.INLINE_IMAGE:
            doc.appendImage(element.copy());
            break;
        default:
            Logger.log('Unsupported element type: ' + type);
    }
}