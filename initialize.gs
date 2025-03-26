/**
 * Alliance Management System - Basic Sheets Setup
 * 
 * This script creates the basic structure of all sheets needed for the system.
 */

function setupAllianceManagementSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all the required sheets
  createDashboard(ss);
  createMembersSheet(ss);
  createActivitiesSheet(ss);
  createParticipationSheet(ss);
  createWeightConfigSheet(ss);
  createEvaluationCriteriaSheet(ss);
  
  // Set Dashboard as the active sheet
  ss.setActiveSheet(ss.getSheetByName('Dashboard'));
}

/**
 * Creates the Dashboard sheet with basic structure
 */
function createDashboard(ss) {
  // Check if Dashboard already exists
  if (ss.getSheetByName('Dashboard')) {
    ss.deleteSheet(ss.getSheetByName('Dashboard'));
  }
  
  // Create Dashboard sheet
  var dashboard = ss.insertSheet('Dashboard');
  
  // Set up basic dashboard structure
  dashboard.setColumnWidth(1, 200);
  dashboard.setColumnWidths(2, 6, 150);
  
  // Add title
  dashboard.getRange('A1:G1').merge();
  dashboard.getRange('A1').setValue('聯盟管理系統');
  dashboard.getRange('A1').setFontSize(16);
  dashboard.getRange('A1').setFontWeight('bold');
  dashboard.getRange('A1').setHorizontalAlignment('center');
  dashboard.getRange('A1').setBackgroundRGB(50, 50, 150);
  dashboard.getRange('A1').setFontColor('white');
  
  // Add sections for key metrics
  dashboard.getRange('A3').setValue('關鍵指標');
  dashboard.getRange('A3').setFontWeight('bold');
  dashboard.getRange('A3').setBackgroundRGB(220, 220, 220);
  
  // Add placeholders for metrics
  dashboard.getRange('A4').setValue('總成員數：');
  dashboard.getRange('A5').setValue('R1成員：');
  dashboard.getRange('A6').setValue('R2成員：');
  dashboard.getRange('A7').setValue('R3成員：');
  dashboard.getRange('A8').setValue('R4成員：');
  dashboard.getRange('A9').setValue('R5成員：');
  dashboard.getRange('A10').setValue('非活躍成員：');
  
  // Add section for important alerts
  dashboard.getRange('A12').setValue('重要警報');
  dashboard.getRange('A12').setFontWeight('bold');
  dashboard.getRange('A12').setBackgroundRGB(220, 220, 220);
  
  // Add placeholder for alerts
  dashboard.getRange('A13:G13').merge();
  dashboard.getRange('A13').setValue('目前沒有重要警報。');
  
  // Add section for rank change suggestions
  dashboard.getRange('A15').setValue('等級變更建議');
  dashboard.getRange('A15').setFontWeight('bold');
  dashboard.getRange('A15').setBackgroundRGB(220, 220, 220);
  
  // Add placeholders for rank suggestions  
  dashboard.getRange('A16:G16').merge();
  dashboard.getRange('A16').setValue('No rank change suggestions at this time.');
  
  // Add Quick Links section
  dashboard.getRange('C3').setValue('快速連結');
  dashboard.getRange('C3').setFontWeight('bold');
  dashboard.getRange('C3').setBackgroundRGB(220, 220, 220);
  
  // Add quick links placeholders
  dashboard.getRange('C4').setValue('記錄活動');
  dashboard.getRange('C5').setValue('更新成員列表');
  dashboard.getRange('C6').setValue('生成報告');
}

/**
 * Creates the Members sheet with appropriate columns and data validation
 */
function createMembersSheet(ss) {
  // Check if Members sheet already exists
  if (ss.getSheetByName('Members')) {
    ss.deleteSheet(ss.getSheetByName('Members'));
  }
  
  // Create Members sheet
  var members = ss.insertSheet('Members');
  
  // Set column widths
  members.setColumnWidth(1, 80);  // MemberID
  members.setColumnWidth(2, 150); // GameName
  members.setColumnWidth(3, 60);  // Rank
  members.setColumnWidth(4, 80);  // Power
  members.setColumnWidth(5, 100); // JoinDate
  members.setColumnWidth(6, 100); // LastActiveDate
  members.setColumnWidth(7, 100); // TotalScore
  members.setColumnWidth(8, 150); // RankSuggestion
  members.setColumnWidth(9, 200); // Notes
  
  // Add headers
  var headers = [
    'MemberID', 'GameName', 'Rank', 'Power', 'JoinDate', 
    'LastActiveDate', 'TotalScore', 'RankSuggestion', 'Notes'
  ];
  
  members.getRange(1, 1, 1, headers.length).setValues([headers]);
  members.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  members.getRange(1, 1, 1, headers.length).setBackgroundRGB(220, 220, 220);
  
  // Freeze the header row
  members.setFrozenRows(1);
  
  // Add data validation for Rank column
  var rankValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['R1', 'R2', 'R3', 'R4', 'R5', 'X'], true)
    .build();
  members.getRange('C2:C1000').setDataValidation(rankValidation);
  
  // Set date format for date columns
  members.getRange('E2:F1000').setNumberFormat('yyyy-mm-dd');
  
  // Set number format for TotalScore
  members.getRange('G2:G1000').setNumberFormat('0.00');
}

/**
 * Creates the Activities sheet with appropriate columns and data validation
 */
function createActivitiesSheet(ss) {
  // Check if Activities sheet already exists
  if (ss.getSheetByName('Activities')) {
    ss.deleteSheet(ss.getSheetByName('Activities'));
  }
  
  // Create Activities sheet
  var activities = ss.insertSheet('Activities');
  
  // Set column widths
  activities.setColumnWidth(1, 100);  // ActivityID
  activities.setColumnWidth(2, 200);  // ActivityName
  activities.setColumnWidth(3, 60);   // Level
  activities.setColumnWidth(4, 100);  // ParentID
  activities.setColumnWidth(5, 100);  // Type
  activities.setColumnWidth(6, 80);   // BaseWeight
  activities.setColumnWidth(7, 80);   // DecayRate
  activities.setColumnWidth(8, 100);  // EnableDecay
  activities.setColumnWidth(9, 100);  // MinThreshold
  activities.setColumnWidth(10, 100); // MaxThreshold
  activities.setColumnWidth(11, 100); // LowScoreFactor
  activities.setColumnWidth(12, 100); // HighScoreFactor
  
  // Add headers
  var headers = [
    'ActivityID', 'ActivityName', 'Level', 'ParentID', 'Type', 
    'BaseWeight', 'DecayRate', 'EnableDecay', 'MinThreshold', 
    'MaxThreshold', 'LowScoreFactor', 'HighScoreFactor'
  ];
  
  activities.getRange(1, 1, 1, headers.length).setValues([headers]);
  activities.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  activities.getRange(1, 1, 1, headers.length).setBackgroundRGB(220, 220, 220);
  
  // Freeze the header row
  activities.setFrozenRows(1);
  
  // Add data validation for Level column
  var levelValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3'], true)
    .build();
  activities.getRange('C2:C1000').setDataValidation(levelValidation);
  
  // Add data validation for Type column
  var typeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Attendance', 'Score'], true)
    .build();
  activities.getRange('E2:E1000').setDataValidation(typeValidation);
  
  // Add data validation for EnableDecay column
  var boolValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .build();
  activities.getRange('H2:H1000').setDataValidation(boolValidation);
  
  // Set number format for numeric columns
  activities.getRange('F2:G1000').setNumberFormat('0.00');
  activities.getRange('I2:J1000').setNumberFormat('0.00');
  activities.getRange('K2:L1000').setNumberFormat('0.00');
}

/**
 * Creates the Participation sheet with appropriate columns and data validation
 */
function createParticipationSheet(ss) {
  // Check if Participation sheet already exists
  if (ss.getSheetByName('Participation')) {
    ss.deleteSheet(ss.getSheetByName('Participation'));
  }
  
  // Create Participation sheet
  var participation = ss.insertSheet('Participation');
  
  // Set column widths
  participation.setColumnWidth(1, 100);  // RecordID
  participation.setColumnWidth(2, 100);  // ActivityID
  participation.setColumnWidth(3, 100);  // Date
  participation.setColumnWidth(4, 100);  // MemberID
  participation.setColumnWidth(5, 150);  // ParticipationStatus
  participation.setColumnWidth(6, 80);   // Score
  participation.setColumnWidth(7, 100);  // MilestoneRating
  participation.setColumnWidth(8, 100);  // ActualScore
  participation.setColumnWidth(9, 100);  // RecordTime
  participation.setColumnWidth(10, 200); // Notes
  
  // Add headers
  var headers = [
    'RecordID', 'ActivityID', 'Date', 'MemberID', 'ParticipationStatus', 
    'Score', 'MilestoneRating', 'ActualScore', 'RecordTime', 'Notes'
  ];
  
  participation.getRange(1, 1, 1, headers.length).setValues([headers]);
  participation.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  participation.getRange(1, 1, 1, headers.length).setBackgroundRGB(220, 220, 220);
  
  // Freeze the header row
  participation.setFrozenRows(1);
  
  // Add data validation for ParticipationStatus column
  var statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Present', 'Absent-Excused', 'Absent-Unexcused', 'N/A'], true)
    .build();
  participation.getRange('E2:E1000').setDataValidation(statusValidation);
  
  // Set date format for Date column
  participation.getRange('C2:C1000').setNumberFormat('yyyy-mm-dd');
  
  // Set datetime format for RecordTime column
  participation.getRange('I2:I1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // Set number format for Score and ActualScore columns
  participation.getRange('F2:F1000').setNumberFormat('0.00');
  participation.getRange('H2:H1000').setNumberFormat('0.00');
}

/**
 * Creates the WeightConfig sheet with appropriate columns and data validation
 */
function createWeightConfigSheet(ss) {
  // Check if WeightConfig sheet already exists
  if (ss.getSheetByName('WeightConfig')) {
    ss.deleteSheet(ss.getSheetByName('WeightConfig'));
  }
  
  // Create WeightConfig sheet
  var weightConfig = ss.insertSheet('WeightConfig');
  
  // Set column widths
  weightConfig.setColumnWidth(1, 150);  // ConfigType
  weightConfig.setColumnWidth(2, 100);  // BaseValue
  weightConfig.setColumnWidth(3, 100);  // WeightPercentage
  weightConfig.setColumnWidth(4, 150);  // ApplicableRank
  weightConfig.setColumnWidth(5, 200);  // Description
  
  // Add headers
  var headers = [
    'ConfigType', 'BaseValue', 'WeightPercentage', 'ApplicableRank', 'Description'
  ];
  
  weightConfig.getRange(1, 1, 1, headers.length).setValues([headers]);
  weightConfig.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  weightConfig.getRange(1, 1, 1, headers.length).setBackgroundRGB(220, 220, 220);
  
  // Freeze the header row
  weightConfig.setFrozenRows(1);
  
  // Set number format for numeric columns
  weightConfig.getRange('B2:C1000').setNumberFormat('0.00');
  
  // Add initial configuration data
  var configData = [
    ['ActivityAttendance', 1.00, 40, 'All', 'Base weight for activity attendance'],
    ['ActivityScore', 1.00, 30, 'All', 'Base weight for activity scores'],
    ['PowerGrowth', 1.00, 5, 'All', 'Base weight for power growth'],
    ['AbsentExcused', 0.50, 0, 'All', 'Multiplier for excused absences'],
    ['AbsentUnexcused', 0.00, 0, 'All', 'Multiplier for unexcused absences']
  ];
  
  weightConfig.getRange(2, 1, configData.length, headers.length).setValues(configData);
}

/**
 * Creates the EvaluationCriteria sheet with appropriate columns
 */
function createEvaluationCriteriaSheet(ss) {
  // Check if EvaluationCriteria sheet already exists
  if (ss.getSheetByName('EvaluationCriteria')) {
    ss.deleteSheet(ss.getSheetByName('EvaluationCriteria'));
  }
  
  // Create EvaluationCriteria sheet
  var criteria = ss.insertSheet('EvaluationCriteria');
  
  // Set column widths
  criteria.setColumnWidth(1, 120);  // EvaluationType
  criteria.setColumnWidth(2, 100);  // FromRank
  criteria.setColumnWidth(3, 100);  // ToRank
  criteria.setColumnWidth(4, 100);  // PowerRequirement
  criteria.setColumnWidth(5, 130);  // ActivityRequirement
  criteria.setColumnWidth(6, 120);  // AttendanceRate
  criteria.setColumnWidth(7, 100);  // TotalScoreRequired
  criteria.setColumnWidth(8, 100);  // EvaluationPeriod
  criteria.setColumnWidth(9, 200);  // Description
  
  // Add headers
  var headers = [
    'EvaluationType', 'FromRank', 'ToRank', 'PowerRequirement', 
    'ActivityRequirement', 'AttendanceRate', 'TotalScoreRequired', 
    'EvaluationPeriod', 'Description'
  ];
  
  criteria.getRange(1, 1, 1, headers.length).setValues([headers]);
  criteria.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  criteria.getRange(1, 1, 1, headers.length).setBackgroundRGB(220, 220, 220);
  
  // Freeze the header row
  criteria.setFrozenRows(1);
  
  // Add data validation for EvaluationType column
  var typeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Promotion', 'Demotion', 'Removal'], true)
    .build();
  criteria.getRange('A2:A1000').setDataValidation(typeValidation);
  
  // Add data validation for FromRank and ToRank columns
  var rankValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['R1', 'R2', 'R3', 'R4', 'R5', 'X'], true)
    .build();
  criteria.getRange('B2:C1000').setDataValidation(rankValidation);
  
  // Set number format for numeric columns
  criteria.getRange('D2:G1000').setNumberFormat('0.00');
  
  // Add initial criteria data
  var criteriaData = [
    ['Promotion', 'R1', 'R2', 5000, 70, 80, 75, '30', 'Criteria for promoting from R1 to R2'],
    ['Promotion', 'R2', 'R3', 10000, 75, 85, 80, '45', 'Criteria for promoting from R2 to R3'],
    ['Promotion', 'R3', 'R4', 20000, 85, 90, 90, '60', 'Criteria for promoting from R3 to R4'],
    ['Demotion', 'R3', 'R2', 8000, 60, 70, 65, '30', 'Criteria for demoting from R3 to R2'],
    ['Demotion', 'R2', 'R1', 4000, 50, 60, 55, '30', 'Criteria for demoting from R2 to R1'],
    ['Removal', 'R1', 'X', 3000, 40, 50, 45, '30', 'Criteria for removing from Alliance']
  ];
  
  criteria.getRange(2, 1, criteriaData.length, headers.length).setValues(criteriaData);
}

// Function to run all setup
function runSetup() {
  setupAllianceManagementSystem();
  Browser.msgBox("Alliance Management System has been set up successfully.");
}