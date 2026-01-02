# GoMissionSoft Product Specification

**GoMissionSoft** is an office automation system that leverages Google Workspace to automate mission assignment issuing and management.

## Features

The main interface of the system is a web application accessible from any device with an internet connection. The system is built using HTML, CSS, and JavaScript, and is hosted on a Google App Engine instance (or Google Apps Script Web App). The system is designed to be used by office workers who need to issue and manage missions for their employees.

The user interface is a form that allows the user to issue missions to their employees. The user can select the employees, the type of mission, the start date and end date of the mission, transportation means, and the destination.

The system is backed by a Google Sheet that serves as a database for the missions. The Google Sheet is the source that feeds the form with the data for autocomplete fields (employees, destinations, transportation means). The Google Sheet is also used to store the missions issued by the user. Issued missions data are processed to generate a Google Doc that can be exported to a Word document.

## The Google Sheet Structure

The database consists of the following sheets:

### 'Personnel' Sheet
Stores employee information.

| Column Name | Description |
| :--- | :--- |
| **EmployeeId** | Unique identifier for the employee |
| **Nom** | Last Name |
| **Prénoms** | First Names |
| **Civilité** | Civility (e.g., Mr, Mrs) |
| **Fonction** | Job Title / Function |
| **Date de naissance** | Date of Birth |
| **Lieu de naissance** | Place of Birth |
| **Grade** | Valid administrative grade |
| **Indice** | Salary index |
| **Matricule** | Employee ID Number |
| **IFU** | Tax ID Number |
| **Adresse complète** | Full Address |
| **Telephone** | Contact Number |
| **Sexe** | Gender |
| **Email** | Email Address |

### 'Transport' Sheet
Stores available transport options.

| Column Name | Description |
| :--- | :--- |
| **Moyen de transport** | Unique column holding vehicle registration number (e.g., "V123456"), or type (e.g., "plane", "Taxi", "Other"). |

### 'Destination' Sheet
Stores valid destinations.

| Column Name | Description |
| :--- | :--- |
| **Destination** | Unique column holding the destination name. |

### 'Missions' Sheet
Stores the history of issued missions.

| Column Name | Description |
| :--- | :--- |
| **MissionID** | Unique column holding mission ID in the format `ODM-${Date.now()}`. |
| **reference** | Reference number of the authorization document (User input). |
| **members** | List of employee IDs (Foreign Key to 'Personnel'). Multiple IDs are separated by a dash (` - `). |
| **missionObject** | Description of the mission's objective (User input). |
| **destinations** | List of destination names (Foreign Key to 'Destination'). Multiple destinations separated by a dash (` - `). |
| **departureDate** | Start date of the mission (Format: `YYYY-MM-DD`). |
| **returnDate** | End date of the mission (Format: `YYYY-MM-DD`). |
| **transportMeans** | Transport means used (Foreign Key to 'Transport'). |
| **budgets** | Budget source (Foreign Key to 'Budget' sheet if exists). |
| **drivers** | List of driver employee IDs (Foreign Key to 'Personnel'). Multiple drivers separated by a dash (` - `). |
| **Completed** | `Boolean`. `true` if the mission is completed, `false` otherwise. |
| **Archived** | `Boolean`. `true` if the mission is archived, `false` otherwise. |
| **CreatedAt** | Timestamp of creation. Format: `YYYY-MM-DD HH:mm:ss`. |
