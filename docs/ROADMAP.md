# Project Roadmap: Multi-Vehicle Mission Support

This document outlines the strategic plan to restructure GoMissionSoft to support complex missions involving multiple groups, vehicles, and drivers within a single mission order.

## Phase 1: Database Restructuring
**Objective**: Enhance the Google Sheets database to support relational data (Mission -> Groups).
- [ ] Create new sheet `'MissionGroups'` to store group-level details (Vehicle, Driver, Members).
- [ ] Update `'Missions'` sheet to store only high-level mission metadata (Reference, Dates, Object).

## Phase 2: User Interface Overhaul
**Objective**: Update the web application form to enable dynamic group creation.
- [ ] Redesign `startingForm.html` to split the flow:
    -   **Step 1: General Info**: Reference, Dates, Object, Destination.
    -   **Step 2: Group Management**: Interactive interface to add/remove groups.
    -   **Step 3: Assignment**: For each group, select 1 Vehicle, 1 Driver, and N Passengers.
- [ ] Implement validation to ensure all members are assigned to a group.

## Phase 3: Backend Logic & Persistence
**Objective**: Update `Code.js` to handle the new data structure.
- [ ] Update `processMissionData` to parse the structured group data.
- [ ] Implement transaction-like saving:
    -   Save main record to `'Missions'`.
    -   Save row(s) to `'MissionGroups'` (one row per group).

## Phase 4: Document Generation Logic
**Objective**: Ensure generated documents reflect the specific group assignments.
- [ ] Refactor `generateIndividualDocument` to:
    -   Iterate through each **Group**.
    -   For each member in the group, generate the ODM using that group's Vehicle and Driver info.
    -   Compile all pages into the single final document.

## Phase 5: Migration & Cleanup
**Objective**: Ensure backward compatibility or clean transition.
- [ ] Archive logic for "simple" missions (or treat a simple mission as a single group).
- [ ] Update helper functions to read from the new schema.
