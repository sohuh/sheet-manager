 const SIGNUPS_SHEET = 'Signups'; // Format: Name, Student ID
 const SUBMISSIONS_SHEET = 'Submissions;  // Format: Name, Student ID, Mission, Photo
 const TRACKER_SHEET = 'Master Tracker';  // Name, Student ID, then the mission columns
 const MISSIONS_SHEET = 'Missions';       // Reference list of the 40 missions.

 const TRACKER_ID_COL = 2; // Student id column, it is the second column 
 const TRACKER_FIRST_MISSION_COL = 3;  // Missions column is the third column.

 const SUB_ID_COL = 3; // Student ID column in submissions sheet.
 const SUB_MISSION_COL = 4; // Mission column in submissions sheet.
 const SUB_STATUS_COL = 6; // Status column 



 // Triggers, when someone submits in the signup form --> add them to the master tracker if not already there.

 function onSignupSubmit(e) {
    const name = e.values[1];
    const id = e.values[2];

    const tracker = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET);
    const lastRow = tracker.getLastRow();
    const ids = lastRow > 1
    ? tracker.getRange(2, TRACKER_ID_COL, lastRow -1).getValues().flat()
    : [];

    if (ids.indexOf(id) === -1) {
        tracker.appendRow([name,id]);
    }
 }   


 / Triggers when someone submits a mission, it marks it as pending in the submissions row.

function onSubmissionSubmit(e) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SUBMISSIONS_SHEET);
    const row = e.range.getRow();

    sheet.getRange(row, SUB_STATUS_COL).setValue('Pending');

    const studentId = e.values[SUB_ID_COL - 1];
    const mission = e.values[SUB_MISSION_COL - 1];
    setTrackerStatus(studentId, mission, 'Pending');
}


/ Triggers when a staff edits the status cell for a submitted mission.

function onEdit(e) {
    const sheet = e.range.getSheet();
    if (sheet.getName() != SUBMISSIONS_SHEET) return;
    if (e.range.getColumn() != SUB_STATUS_COL) return;

    const row = e.range.getRow();
    const status = e.range.GetValue();
    const studentId = sheet.getRange(row, SUB_ID_COL).getValue();
    const mission = sheet.getRange(row, SUB_MISSION_COL).getValue();


    if (status == 'Approved') {
        setTrackerStatus(studentId, mission, 'Approved');
    }   else if (status == 'Rejected'){
        setTrackerStatus(studentId, mission, '');

    }
}


/ Helper, finds the student row + mission column in the master tracker and sets the cell.

function setTrackerStatus(studentId, missionName, status) {
    const tracker = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET); 
    const lastRow = tracker.getLastRow();
    const lastCol = tracker.getLastColumn();
    if (lastRow < 2) return; 

    const ids = tracker.getRange(2, TRACKER_ID_COL, lastRow - 1).GetValues().flat();
    const headers - tracker.getRange(1,1,1, lastCol).getValues()[0];

    const rowIndex = ids.indexOf(studentId);
    const colIndex = headers.indexOf(missionName);
    if (rowIndex === -1 || colIndex === -1) return;

    tracker.getRange(rowIndex +2, colIndex + 1).setValue(status);
    } 



/ Run once, builds Master Tracker row from the missions list.

function setupTrackerHeaders() {
    const missionsSheet = SpreadsheetApp.getActive().getSheetByName(MISSIONS_SHEET);
    const missions = missionsSheet.GetRange(2,1,missionsSheet.getLastRow() - 1).getValues().flat().filter(string);


    const tracker = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET);
    tracker.getRange(1,1,1,2).setValues([['Name', 'Student ID]]);
    tracker.getRange(1,TRACKER_FIRST_MISSION_COL, 1, missions.length).setValues([missions]);
}