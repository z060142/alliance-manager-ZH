/**
 * Alliance Management System - Menu Management
 * Central file for all menu creation and management functions
 */

/**
 * Creates the Alliance System menu when the spreadsheet is opened
 * This is the single integrated menu creation function replacing all individual onOpen functions
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create the main menu with all integrated options
  ui.createMenu('聯盟系統')
    // Activity and Member Management
    .addItem('記錄活動', 'showActivityRecordingDialog')
    .addItem('管理成員', 'showMemberManagementDialog')
    .addItem('管理活動', 'showActivityManagementDialog')
    .addSeparator()
    
    // Analysis
    .addItem('成員排名', 'showRankedMemberList')
    .addItem('查看活動記錄', 'showActivityViewerDialog')
    .addItem('更新儀表板', 'updateDashboard')
    .addSeparator()
    
    // Evaluation
    .addItem('計算分數', 'calculateAllMemberScores')
    .addItem('生成等級建議', 'generateAllRankSuggestions')
    .addSeparator()

    // System Administration
    //.addItem('Setup System', 'runSetup')
    .addItem('設置觸發器', 'setupTriggers')
    .addItem('重新計算活動分數', 'recalculateActualScores')
    .addItem('除錯：顯示活動結構', 'showActivityStructure')
    .addToUi();
}

/**
 * Opens the activity recording as a modal dialog
 */
function showActivityRecordingDialog() {
  var html = HtmlService.createTemplateFromFile('ActivityRecording')
    .evaluate()
    .setWidth(1000)
    .setHeight(600)
    .setTitle('Record Activity');
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Activity');
}

/**
 * @deprecated Use showActivityRecordingDialog() instead
 * This function is kept for backward compatibility but will be removed in future versions
 */
function showActivityRecordingSidebar() {
  // Redirect to the dialog version
  showActivityRecordingDialog();
  
  // Optionally show a deprecation notice
  SpreadsheetApp.getActive().toast(
    'The sidebar version is deprecated. Using dialog version instead.',
    'Notice',
    5
  );
}

/**
 * Include external HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}