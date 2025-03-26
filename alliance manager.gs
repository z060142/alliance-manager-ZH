/**
 * Alliance Management System - Activity Management Module
 * Functions to add to AllianceManagementSystem.gs


// Update onOpen function to include activity management
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alliance System')
    .addItem('Record Activity', 'showActivityRecordingDialog') // Changed function name
    .addItem('Manage Members', 'showMemberManagementSidebar')
    .addItem('Manage Activities', 'showActivityManagementDialog')
    .addSeparator()
    .addItem('Calculate Scores', 'calculateAllMemberScores')
    .addItem('Generate Rank Suggestions', 'generateAllRankSuggestions')
    .addItem('Update Dashboard', 'updateDashboard')
    .addItem('Member Rankings', 'showRankedMemberList') // New menu item
    .addSeparator()
    //.addItem('Setup System', 'runSetup')
    .addItem('Setup Triggers', 'setupTriggers')
    .addItem('Recalculate Activity Scores', 'recalculateActualScores') // 新的简化函数
    .addItem('Debug: Print Activity Structure', 'showActivityStructure') // 添加调试选项
    .addToUi();
}
 */

/**
 * Opens the activity management dialog
 */
function showActivityManagementDialog() {
  var html = HtmlService.createTemplateFromFile('ActivityManager')
    .evaluate()
    .setWidth(900)
    .setHeight(600)
    .setTitle('Multi-Level Activity Management');
  SpreadsheetApp.getUi().showModalDialog(html, 'Multi-Level Activity Management');
}

/**
 * Opens the activity recording as a modal dialog instead of sidebar
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
 * Get all activities in a hierarchical structure
 */
function getActivityHierarchy() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  // Find column indices
  var idIdx = headers.indexOf('ActivityID');
  var nameIdx = headers.indexOf('ActivityName');
  var levelIdx = headers.indexOf('Level');
  var parentIdIdx = headers.indexOf('ParentID');
  var typeIdx = headers.indexOf('Type');
  var baseWeightIdx = headers.indexOf('BaseWeight');
  var decayRateIdx = headers.indexOf('DecayRate');
  var enableDecayIdx = headers.indexOf('EnableDecay');
  var minThresholdIdx = headers.indexOf('MinThreshold');
  var maxThresholdIdx = headers.indexOf('MaxThreshold');
  var lowScoreFactorIdx = headers.indexOf('LowScoreFactor');
  var highScoreFactorIdx = headers.indexOf('HighScoreFactor');
  
  // Convert to objects for easier manipulation
  var activities = data.map(function(row) {
    return {
      id: row[idIdx],
      name: row[nameIdx],
      level: row[levelIdx],
      parentId: row[parentIdIdx],
      type: row[typeIdx],
      baseWeight: row[baseWeightIdx],
      decayRate: row[decayRateIdx],
      enableDecay: row[enableDecayIdx],
      minThreshold: row[minThresholdIdx],
      maxThreshold: row[maxThresholdIdx],
      lowScoreFactor: row[lowScoreFactorIdx],
      highScoreFactor: row[highScoreFactorIdx],
      children: []
    };
  });
  
  // Calculate absolute weight
  activities = calculateAbsoluteWeights(activities);
  
  // Build hierarchy
  var hierarchy = [];
  var activityMap = {};
  
  // Create a map for quick lookup
  activities.forEach(function(activity) {
    activityMap[activity.id] = activity;
  });
  
  // Build the tree
  activities.forEach(function(activity) {
    if (activity.level == 1) {
      // Root level activity
      hierarchy.push(activity);
    } else if (activity.parentId && activityMap[activity.parentId]) {
      // Child activity
      activityMap[activity.parentId].children.push(activity);
    }
  });
  
  return hierarchy;
}

/**
 * Calculate absolute weights based on parent-child relationships
 */
function calculateAbsoluteWeights(activities) {
  // Get level 1 activities total weight
  var level1TotalWeight = 0;
  var level1Activities = activities.filter(function(a) { return a.level == 1; });
  level1Activities.forEach(function(a) { level1TotalWeight += a.baseWeight; });
  
  // Create a map for quick lookup
  var activityMap = {};
  activities.forEach(function(activity) {
    activityMap[activity.id] = activity;
    // Initialize absoluteWeight with baseWeight
    activity.absoluteWeight = activity.baseWeight;
    
    // Calculate relative weight (for display)
    if (activity.level == 1) {
      activity.relativeWeight = (activity.baseWeight / level1TotalWeight * 100).toFixed(0);
    } else {
      // Will be calculated later when we have parent info
      activity.relativeWeight = 0;
    }
  });
  
  // Calculate child relative weights for each parent
  activities.forEach(function(activity) {
    if (activity.level > 1 && activity.parentId && activityMap[activity.parentId]) {
      var siblings = activities.filter(function(a) { 
        return a.parentId === activity.parentId; 
      });
      
      var siblingTotalWeight = 0;
      siblings.forEach(function(s) { siblingTotalWeight += s.baseWeight; });
      
      if (siblingTotalWeight > 0) {
        activity.relativeWeight = (activity.baseWeight / siblingTotalWeight * 100).toFixed(0);
      }
      
      // Calculate absolute weight based on parent
      var parent = activityMap[activity.parentId];
      activity.absoluteWeight = (parent.absoluteWeight * (activity.baseWeight / siblingTotalWeight)).toFixed(0);
    }
  });
  
  return activities;
}

/**
 * Get activity by ID
 */
function getActivityById(activityId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var idIdx = headers.indexOf('ActivityID');
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][idIdx] === activityId) {
      var activity = {};
      headers.forEach(function(header, index) {
        activity[header] = data[i][index];
      });
      return activity;
    }
  }
  
  return null;
}

/**
 * Add a new activity
 */
function addActivity(activityData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  // Generate a unique activity ID based on parent path
  var activityId = generateActivityId(activityData);
  
  // Set additional data
  activityData.ActivityID = activityId;
  
  // Set default values for missing fields
  if (!activityData.DecayRate) activityData.DecayRate = 0.05;
  if (activityData.EnableDecay === undefined) activityData.EnableDecay = true;
  if (!activityData.MinThreshold) activityData.MinThreshold = 0;
  if (!activityData.MaxThreshold) activityData.MaxThreshold = 0;
  if (!activityData.LowScoreFactor) activityData.LowScoreFactor = -0.2;
  if (!activityData.HighScoreFactor) activityData.HighScoreFactor = 0.2;
  
  // Create row array in the correct order
  var newRow = headers.map(function(header) {
    return activityData[header] !== undefined ? activityData[header] : '';
  });
  
  // Add new row
  sheet.appendRow(newRow);
  
  return activityId;
}

/**
 * Generate an activity ID based on parent path
 */
function generateActivityId(activityData) {
  // For level 1, use a simple prefix
  if (activityData.Level == 1) {
    return "ACT" + new Date().getTime().toString().substring(7);
  }
  
  // For child activities, incorporate parent info
  var parentActivity = getActivityById(activityData.ParentID);
  if (!parentActivity) {
    return "ACT" + new Date().getTime().toString().substring(7);
  }
  
  // Create a shorthand based on parent ID and activity name
  var shortName = activityData.ActivityName.replace(/[^A-Z0-9]/gi, '').substring(0, 5).toUpperCase();
  var uniqueSuffix = new Date().getTime().toString().substring(10);
  
  return parentActivity.ActivityID + "_" + shortName + uniqueSuffix;
}

/**
 * Update an existing activity
 */
function updateActivity(activityData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var idIdx = headers.indexOf('ActivityID');
  var activityId = activityData.ActivityID;
  
  // Find the row index for the activity
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][idIdx] === activityId) {
      rowIndex = i + 2; // +2 because we removed headers and sheet is 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Activity not found: ' + activityId);
  }
  
  // Update the row with new data
  headers.forEach(function(header, colIndex) {
    if (activityData.hasOwnProperty(header)) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(activityData[header]);
    }
  });
  
  return true;
}

/**
 * Delete an activity and its children
 */
function deleteActivity(activityId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var idIdx = headers.indexOf('ActivityID');
  var parentIdIdx = headers.indexOf('ParentID');
  
  // Find activities to delete (the activity itself and all its children)
  var idsToDelete = [activityId];
  var children = findAllChildren(data, activityId, idIdx, parentIdIdx);
  idsToDelete = idsToDelete.concat(children);
  
  // Delete from bottom to top to avoid shifting issues
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 0; i--) {
    if (idsToDelete.indexOf(data[i][idIdx]) !== -1) {
      rowsToDelete.push(i + 2); // +2 because we removed headers and sheet is 1-indexed
    }
  }
  
  // Delete rows
  for (var i = 0; i < rowsToDelete.length; i++) {
    sheet.deleteRow(rowsToDelete[i]);
  }
  
  return idsToDelete.length;
}

/**
 * Recursively find all children of an activity
 */
function findAllChildren(data, parentId, idIdx, parentIdIdx) {
  var children = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][parentIdIdx] === parentId) {
      var childId = data[i][idIdx];
      children.push(childId);
      // Recursively find grandchildren
      var grandchildren = findAllChildren(data, childId, idIdx, parentIdIdx);
      children = children.concat(grandchildren);
    }
  }
  
  return children;
}

/**
 * Get activity path
 */
function getActivityPath(activityId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var idIdx = headers.indexOf('ActivityID');
  var nameIdx = headers.indexOf('ActivityName');
  var parentIdIdx = headers.indexOf('ParentID');
  
  var path = [];
  var currentId = activityId;
  
  while (currentId) {
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (data[i][idIdx] === currentId) {
        path.unshift(data[i][nameIdx]); // Add to beginning of path
        currentId = data[i][parentIdIdx]; // Move to parent
        found = true;
        break;
      }
    }
    
    if (!found) {
      break;
    }
  }
  
  return path.join(' > ');
}

/**
 * Save participation records to the Participation sheet
 * Modified to only save records with actual data
 */
function saveParticipationRecords(records) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var timestamp = new Date();
  
  var activitiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var activityIdIdx = activitiesHeaders.indexOf('ActivityID');
  var typeIdx = activitiesHeaders.indexOf('Type');
  var minThresholdIdx = activitiesHeaders.indexOf('MinThreshold');
  var maxThresholdIdx = activitiesHeaders.indexOf('MaxThreshold');
  var lowScoreFactorIdx = activitiesHeaders.indexOf('LowScoreFactor');
  var highScoreFactorIdx = activitiesHeaders.indexOf('HighScoreFactor');
  
  // Create activity property mapping
  var activityMap = {};
  for (var i = 0; i < activitiesData.length; i++) {
    var row = activitiesData[i];
    activityMap[row[activityIdIdx]] = {
      type: row[typeIdx],
      minThreshold: row[minThresholdIdx],
      maxThreshold: row[maxThresholdIdx],
      lowScoreFactor: row[lowScoreFactorIdx],
      highScoreFactor: row[highScoreFactorIdx]
    };
  }
  
  // Get attendance multipliers
  var weightConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WeightConfig');
  var weightData = weightConfigSheet.getDataRange().getValues();
  var attendanceMult = 1.0; // Default
  var absentExcusedMult = 0.5; // Default
  var absentUnexcusedMult = 0.0; // Default
  
  // Find attendance multipliers in WeightConfig
  for (var i = 1; i < weightData.length; i++) {
    var configType = weightData[i][0];
    if (configType === "ActivityAttendance") {
      attendanceMult = weightData[i][1];
    } else if (configType === "AbsentExcused") {
      absentExcusedMult = weightData[i][1];
    } else if (configType === "AbsentUnexcused") {
      absentUnexcusedMult = weightData[i][1];
    }
  }
  
  var newRecords = [];
  
  records.forEach(function(record, index) {
    var activity = activityMap[record.activityId];
    
    if (!activity) {
      console.error("Activity not found: " + record.activityId);
      return;
    }
    
    // Skip records with empty values
    if (activity.type === "Score" && (!record.score || record.score === '')) {
      return; // Skip empty score records
    }
    
    if (activity.type === "Attendance" && (!record.status || record.status === '')) {
      return; // Skip empty status records
    }
    
    var recordId = "REC" + (lastRow + newRecords.length);
    var milestoneRating = "";
    var actualScore = 0;
    
    // Calculate milestone rating and actual score
    if (activity.type === "Score") {
      var score = parseFloat(record.score) || 0;
      
      // Using new scoring from 0 to max milestone
      if (score < activity.minThreshold) {
        milestoneRating = "Below";
        // Calculate score below min threshold
        var proportionOfMin = (activity.minThreshold > 0) ? 
                              (score / activity.minThreshold) : 0;
        
        var baseScore = 1000 * (activity.minThreshold / activity.maxThreshold) * proportionOfMin;
        actualScore = baseScore * (1 + activity.lowScoreFactor);
      } else if (score > activity.maxThreshold) {
        milestoneRating = "Above";
        actualScore = 1000 * (1 + activity.highScoreFactor);
      } else {
        milestoneRating = "Within";
        actualScore = 1000 * (score / activity.maxThreshold);
      }
    } else { // Attendance type activity
      var status = record.status;
      milestoneRating = "N/A";
      
      // Calculate score based on attendance status
      if (status === "Present") {
        actualScore = 1000 * attendanceMult;
      } else if (status === "Absent-Excused") {
        actualScore = 1000 * absentExcusedMult;
      } else { // Absent-Unexcused or other
        actualScore = 1000 * absentUnexcusedMult;
      }
    }
    
    // Create the new record
    newRecords.push([
      recordId,                    // RecordID
      record.activityId,           // ActivityID
      new Date(record.date),       // Date
      record.memberId,             // MemberID
      record.status,               // ParticipationStatus
      record.score || "",          // Score
      milestoneRating,             // MilestoneRating
      actualScore,                 // ActualScore
      timestamp,                   // RecordTime
      record.notes || ""           // Notes
    ]);
  });
  
  // Write all new records at once
  if (newRecords.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRecords.length, 10).setValues(newRecords);
  }
  
  return newRecords.length;
}

/**
 * Alliance Management System - Activity Viewer Module
 * This module allows viewing activity participation records by date or activity
 */

/**
 * Opens the activity viewer dialog
 */
function showActivityViewerDialog() {
  var html = HtmlService.createTemplateFromFile('ActivityViewer')
    .evaluate()
    .setWidth(900)
    .setHeight(600)
    .setTitle('Activity Viewer');
  SpreadsheetApp.getUi().showModalDialog(html, 'Activity Viewer');
}

/**
 * Get all dates with activity records
 */
function getActivityDates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var dateIdx = headers.indexOf('Date');
  
  if (dateIdx === -1) {
    return [];
  }
  
  // Extract all dates and remove duplicates
  var dates = data.map(function(row) {
    return row[dateIdx];
  });
  
  // Convert to date strings in yyyy-MM-dd format
  var dateStrings = dates.map(function(date) {
    if (date instanceof Date) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return '';
  }).filter(function(date) {
    return date !== '';
  });
  
  // Remove duplicates
  var uniqueDates = [];
  dateStrings.forEach(function(date) {
    if (uniqueDates.indexOf(date) === -1) {
      uniqueDates.push(date);
    }
  });
  
  // Sort dates in descending order (newest first)
  uniqueDates.sort(function(a, b) {
    return new Date(b) - new Date(a);
  });
  
  return uniqueDates;
}

/**
 * Get all activities with records
 */
function getActivitiesWithRecords() {
  var participationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var participationData = participationSheet.getDataRange().getValues();
  var headers = participationData.shift(); // Remove headers
  
  var activityIdIdx = headers.indexOf('ActivityID');
  
  if (activityIdIdx === -1) {
    return [];
  }
  
  // Extract all activity IDs and remove duplicates
  var activityIds = participationData.map(function(row) {
    return row[activityIdIdx];
  });
  
  var uniqueActivityIds = [];
  activityIds.forEach(function(id) {
    if (uniqueActivityIds.indexOf(id) === -1) {
      uniqueActivityIds.push(id);
    }
  });
  
  // Get activity names from Activities sheet
  var activitiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var idIdx = activitiesHeaders.indexOf('ActivityID');
  var nameIdx = activitiesHeaders.indexOf('ActivityName');
  
  if (idIdx === -1 || nameIdx === -1) {
    return [];
  }
  
  // Map activity IDs to names
  var activityMap = {};
  activitiesData.forEach(function(row) {
    activityMap[row[idIdx]] = row[nameIdx];
  });
  
  // Create result array with both ID and name
  var activities = uniqueActivityIds.map(function(id) {
    return {
      id: id,
      name: activityMap[id] || id
    };
  });
  
  // Sort by name
  activities.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  
  return activities;
}

/**
 * Get dates for a specific activity
 */
function getDatesForActivity(activityId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var dateIdx = headers.indexOf('Date');
  var activityIdIdx = headers.indexOf('ActivityID');
  
  if (dateIdx === -1 || activityIdIdx === -1) {
    return [];
  }
  
  // Filter records for the specific activity
  var filteredData = data.filter(function(row) {
    return row[activityIdIdx] === activityId;
  });
  
  // Extract dates and remove duplicates
  var dates = filteredData.map(function(row) {
    return row[dateIdx];
  });
  
  // Convert to date strings
  var dateStrings = dates.map(function(date) {
    if (date instanceof Date) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return '';
  }).filter(function(date) {
    return date !== '';
  });
  
  // Remove duplicates
  var uniqueDates = [];
  dateStrings.forEach(function(date) {
    if (uniqueDates.indexOf(date) === -1) {
      uniqueDates.push(date);
    }
  });
  
  // Sort dates in descending order (newest first)
  uniqueDates.sort(function(a, b) {
    return new Date(b) - new Date(a);
  });
  
  return uniqueDates;
}

/**
 * Get activities for a specific date
 */
function getActivitiesForDate(dateString) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var dateIdx = headers.indexOf('Date');
  var activityIdIdx = headers.indexOf('ActivityID');
  
  if (dateIdx === -1 || activityIdIdx === -1) {
    return [];
  }
  
  // Convert the dateString to a Date object
  var targetDate = new Date(dateString);
  
  // Filter records for the specific date
  var filteredData = data.filter(function(row) {
    if (row[dateIdx] instanceof Date) {
      var rowDate = row[dateIdx];
      return rowDate.getFullYear() === targetDate.getFullYear() &&
             rowDate.getMonth() === targetDate.getMonth() &&
             rowDate.getDate() === targetDate.getDate();
    }
    return false;
  });
  
  // Extract activity IDs and remove duplicates
  var activityIds = filteredData.map(function(row) {
    return row[activityIdIdx];
  });
  
  var uniqueActivityIds = [];
  activityIds.forEach(function(id) {
    if (uniqueActivityIds.indexOf(id) === -1) {
      uniqueActivityIds.push(id);
    }
  });
  
  // Get activity names from Activities sheet
  var activitiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var idIdx = activitiesHeaders.indexOf('ActivityID');
  var nameIdx = activitiesHeaders.indexOf('ActivityName');
  
  if (idIdx === -1 || nameIdx === -1) {
    return [];
  }
  
  // Map activity IDs to names
  var activityMap = {};
  activitiesData.forEach(function(row) {
    activityMap[row[idIdx]] = row[nameIdx];
  });
  
  // Create result array with both ID and name
  var activities = uniqueActivityIds.map(function(id) {
    return {
      id: id,
      name: activityMap[id] || id
    };
  });
  
  // Sort by name
  activities.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  
  return activities;
}

/**
 * Get activity details by ID
 */
function getActivityDetails(activityId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var idIdx = headers.indexOf('ActivityID');
  var nameIdx = headers.indexOf('ActivityName');
  var typeIdx = headers.indexOf('Type');
  
  if (idIdx === -1 || nameIdx === -1 || typeIdx === -1) {
    return null;
  }
  
  // Find the activity
  for (var i = 0; i < data.length; i++) {
    if (data[i][idIdx] === activityId) {
      return {
        id: data[i][idIdx],
        name: data[i][nameIdx],
        type: data[i][typeIdx]
      };
    }
  }
  
  return null;
}

/**
 * Get member participation for a specific activity and date
 */
function getMemberParticipation(activityId, dateString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all members
  var membersSheet = ss.getSheetByName('Members');
  var membersData = membersSheet.getDataRange().getValues();
  var membersHeaders = membersData.shift();
  
  var memberIdIdx = membersHeaders.indexOf('MemberID');
  var gameNameIdx = membersHeaders.indexOf('GameName');
  var rankIdx = membersHeaders.indexOf('Rank');
  
  if (memberIdIdx === -1 || gameNameIdx === -1 || rankIdx === -1) {
    return { error: 'Member sheet columns not found' };
  }
  
  // Get all active members (not rank 'X')
  var activeMembers = {};
  var inactiveMembers = {};
  
  membersData.forEach(function(row) {
    var memberId = row[memberIdIdx];
    var memberInfo = {
      id: memberId,
      name: row[gameNameIdx],
      rank: row[rankIdx],
      status: 'Not Recorded' // Default status
    };
    
    if (row[rankIdx] === 'X') {
      inactiveMembers[memberId] = memberInfo;
    } else {
      activeMembers[memberId] = memberInfo;
    }
  });
  
  // Get activity details
  var activityDetails = getActivityDetails(activityId);
  if (!activityDetails) {
    return { error: 'Activity not found' };
  }
  
  // Get participation records
  var participationSheet = ss.getSheetByName('Participation');
  var participationData = participationSheet.getDataRange().getValues();
  var participationHeaders = participationData.shift();
  
  var activityIdIdx = participationHeaders.indexOf('ActivityID');
  var dateIdx = participationHeaders.indexOf('Date');
  var memberIdIdx = participationHeaders.indexOf('MemberID');
  var statusIdx = participationHeaders.indexOf('ParticipationStatus');
  var scoreIdx = participationHeaders.indexOf('Score');
  
  if (activityIdIdx === -1 || dateIdx === -1 || memberIdIdx === -1 || statusIdx === -1) {
    return { error: 'Participation sheet columns not found' };
  }
  
  // Convert the dateString to a Date object
  var targetDate = new Date(dateString);
  
  // Process participation records
  participationData.forEach(function(row) {
    if (row[activityIdIdx] === activityId) {
      var rowDate = row[dateIdx];
      if (rowDate instanceof Date && 
          rowDate.getFullYear() === targetDate.getFullYear() &&
          rowDate.getMonth() === targetDate.getMonth() &&
          rowDate.getDate() === targetDate.getDate()) {
        
        var memberId = row[memberIdIdx];
        var status = row[statusIdx];
        var score = row[scoreIdx];
        
        // Update member status if they exist in our member maps
        if (activeMembers[memberId]) {
          if (activityDetails.type === 'Attendance') {
            activeMembers[memberId].status = status;
          } else { // Score type
            activeMembers[memberId].status = score ? 'Present' : 'Not Recorded';
            activeMembers[memberId].score = score;
          }
        } else if (inactiveMembers[memberId]) {
          if (activityDetails.type === 'Attendance') {
            inactiveMembers[memberId].status = status;
          } else { // Score type
            inactiveMembers[memberId].status = score ? 'Present' : 'Not Recorded';
            inactiveMembers[memberId].score = score;
          }
        }
      }
    }
  });
  
  // Convert member maps to arrays
  var activeMembersArray = Object.values(activeMembers);
  var inactiveMembersArray = Object.values(inactiveMembers);
  
  // Sort members by ID
  activeMembersArray.sort(function(a, b) {
    return a.id.localeCompare(b.id);
  });
  
  inactiveMembersArray.sort(function(a, b) {
    return a.id.localeCompare(b.id);
  });
  
  // Return structured response
  return {
    activityDetails: activityDetails,
    date: dateString,
    activeMembers: activeMembersArray,
    inactiveMembers: inactiveMembersArray
  };
}

/**
 * Alliance Management System - Combined Code
 * This file contains the original functions with updated scoring system.


// Keep the original onOpen function
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alliance System')
    .addItem('Record Activity', 'showActivityRecordingSidebar')
    .addItem('Calculate Scores', 'calculateAllMemberScores')
    .addItem('Generate Rank Suggestions', 'generateAllRankSuggestions')
    .addItem('Update Dashboard', 'updateDashboard')
    .addSeparator()
    .addItem('Setup System', 'runSetup')
    .addItem('Setup Triggers', 'setupTriggers')
    .addToUi();
}
 */

/**
 * Calculate scores for all members with configurable power weight
 */
function calculateAllMemberScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var membersSheet = ss.getSheetByName('Members');
  var membersData = membersSheet.getDataRange().getValues();
  var headers = membersData.shift(); // Remove headers
  
  // Find column indices
  var memberIdColIndex = headers.indexOf('MemberID');
  var totalScoreColIndex = headers.indexOf('TotalScore');
  var powerColIndex = headers.indexOf('Power');
  var rankColIndex = headers.indexOf('Rank');
  
  // Skip if required columns not found
  if (memberIdColIndex === -1 || totalScoreColIndex === -1 || powerColIndex === -1) {
    Browser.msgBox('Error', 'Required columns not found in Members sheet.', Browser.Buttons.OK);
    return;
  }
  
  // Get power weight from WeightConfig
  var weightConfigSheet = ss.getSheetByName('WeightConfig');
  var weightData = weightConfigSheet.getDataRange().getValues();
  var powerWeight = 0.05; // Default 5%
  
  // Find the power weight in WeightConfig
  for (var i = 1; i < weightData.length; i++) {
    var configType = weightData[i][0];
    if (configType === "PowerGrowth") {
      powerWeight = weightData[i][2] / 100; // Convert percentage to decimal
      break;
    }
  }
  
  // 记录当前使用的权重，以便调试
  Logger.log("Using power weight: " + powerWeight + " (" + (powerWeight * 100) + "%)");
  
  // 计算活动分数的基准分值 (总分1000 - 战力部分)
  var ACTIVITY_BASE = 1000 * (1 - powerWeight);
  
  // Get all active members
  var activeMembers = [];
  for (var i = 0; i < membersData.length; i++) {
    var row = membersData[i];
    var memberId = row[memberIdColIndex];
    var rank = row[rankColIndex];
    var power = row[powerColIndex];
    
    // Skip inactive members
    if (rank === 'X') continue;
    
    activeMembers.push({
      memberId: memberId,
      power: power,
      rowIndex: i + 2 // +2 because we removed header and 0-indexing
    });
  }
  
  // 将powerWeight传递给calculatePowerDistribution函数
  var powerDistribution = calculatePowerDistribution(activeMembers, powerWeight);
  
  // 记录一些诊断信息
  var diagInfo = "战力分数分布（权重: " + (powerWeight * 100) + "%）：\n";
  var diagCount = Math.min(activeMembers.length, 10); // 最多显示10条
  
  // 按战力排序以便显示最高战力的成员
  var sortedMembers = activeMembers.slice().sort(function(a, b) {
    return b.power - a.power;
  });
  
  for (var i = 0; i < diagCount; i++) {
    var memberId = sortedMembers[i].memberId;
    var power = sortedMembers[i].power;
    var powerScore = powerDistribution[memberId] || 0;
    diagInfo += "成员: " + memberId + ", 战力: " + power + ", 战力分: " + powerScore.toFixed(2) + 
               " (总分" + (powerWeight * 100) + "%)\n";
  }
  
  // Calculate score for each member
  for (var i = 0; i < activeMembers.length; i++) {
    var memberId = activeMembers[i].memberId;
    
    // 战力分数现在是根据配置的权重动态计算的
    var powerScore = powerDistribution[memberId] || 0;
    
    // 活动分数
    var activityScore = calculateMemberActivityScore(memberId);
    
    // 直接相加
    var totalScore = activityScore + powerScore;
    
    // Round to 2 decimal places
    totalScore = Math.round(totalScore * 100) / 100;
    
    // Update the total score in the Members sheet
    membersSheet.getRange(activeMembers[i].rowIndex, totalScoreColIndex + 1).setValue(totalScore);
  }
  
  // 显示带诊断信息的成功消息
  Browser.msgBox('成功', '已計算 ' + activeMembers.length + ' 位成員的分數。', Browser.Buttons.OK);
}
  
/**
 * Calculate power distribution based on normal distribution
 * @param {Array} members - Array of member objects with power values
 * @return {Object} - Object mapping memberIds to power scores
 */
/**
 * 修复战力评分系统 - 使用简单的百分位排名而非有问题的Z分数
 */
function calculatePowerDistribution(members, powerWeight) {
  // 接收powerWeight参数，而不是使用硬编码的比例
  
  // 创建一个成员副本，以便进行排序而不影响原始数组
  var membersCopy = members.slice();
  
  // 按战力从高到低排序
  membersCopy.sort(function(a, b) {
    return b.power - a.power;
  });
  
  // 计算战力最高分 (总分1000 * 配置的权重百分比)
  var MAX_POWER_SCORE = 1000 * powerWeight;
  
  // 基于排名分配分数
  var distribution = {};
  var memberCount = membersCopy.length;
  
  // 为每个成员计算战力分数
  for (var i = 0; i < memberCount; i++) {
    var memberId = membersCopy[i].memberId;
    
    // 计算百分位数: (总数 - 排名) / 总数
    // 排名从0开始，所以第一名的百分位是1.0，最后一名接近0
    var percentile = (memberCount - i) / memberCount;
    
    // 直接将百分位映射到动态计算的最高分范围
    var powerScore = MAX_POWER_SCORE * percentile;
    
    // 储存计算出的分数
    distribution[memberId] = powerScore;
  }
  
  return distribution;
}

/**
 * Normal cumulative distribution function approximation
 * @param {number} z - Z-score
 * @return {number} - Percentile (0-1)
 */
function normalCDF(z) {
  // Approximation of the normal CDF
  if (z < -6) return 0;
  if (z > 6) return 1;
  
  var p = 0.5;
  var t = 1 + z * (0.04986735 + z * (-0.02758434 + z * (-0.00577221 + z * (0.02557185 + z * (0.02546863)))));
  p -= 0.5 * Math.pow(t, -16) * Math.exp(-0.5 * z * z - 1.26551223);
  
  return z > 0 ? 1 - p : p;
}

/**
 * Calculate activity score for a specific member
 * @param {string} memberId - The member's ID
 * @return {number} - The calculated activity score
 */
function calculateMemberActivityScore(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 获取参与记录
  var participationSheet = ss.getSheetByName('Participation');
  var participationData = participationSheet.getDataRange().getValues();
  var participationHeaders = participationData.shift();
  
  var memberIdColIndex = participationHeaders.indexOf('MemberID');
  var activityIdColIndex = participationHeaders.indexOf('ActivityID');
  var dateColIndex = participationHeaders.indexOf('Date');
  var actualScoreColIndex = participationHeaders.indexOf('ActualScore');
  
  // 获取活动信息
  var activitiesSheet = ss.getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var activityIdIdx = activitiesHeaders.indexOf('ActivityID');
  var nameIdx = activitiesHeaders.indexOf('ActivityName');
  var levelIdx = activitiesHeaders.indexOf('Level');
  var parentIdIdx = activitiesHeaders.indexOf('ParentID');
  var baseWeightIdx = activitiesHeaders.indexOf('BaseWeight');
  var decayRateIdx = activitiesHeaders.indexOf('DecayRate');
  var enableDecayIdx = activitiesHeaders.indexOf('EnableDecay');
  
  // 构建活动树结构
  var activityMap = {};
  var level1Activities = [];
  
  // 第一步：映射所有活动
  for (var i = 0; i < activitiesData.length; i++) {
    var activityId = activitiesData[i][activityIdIdx];
    var name = activitiesData[i][nameIdx];
    var level = activitiesData[i][levelIdx];
    var parentId = activitiesData[i][parentIdIdx];
    
    activityMap[activityId] = {
      id: activityId,
      name: name,
      level: level,
      parentId: parentId,
      baseWeight: activitiesData[i][baseWeightIdx],
      decayRate: activitiesData[i][decayRateIdx],
      enableDecay: activitiesData[i][enableDecayIdx],
      children: [],
      records: [],
      adjustedWeight: null // 将在后续步骤计算
    };
    
    if (level == 1) {
      level1Activities.push(activityMap[activityId]);
    }
  }
  
  // 第二步：构建活动树
  for (var id in activityMap) {
    var activity = activityMap[id];
    if (activity.level > 1 && activity.parentId && activityMap[activity.parentId]) {
      activityMap[activity.parentId].children.push(activity);
    }
  }
  
  // 第三步：筛选该成员的参与记录
  var memberRecords = participationData.filter(function(row) {
    return row[memberIdColIndex] === memberId;
  });
  
  // 将记录添加到对应活动
  memberRecords.forEach(function(record) {
    var activityId = record[activityIdColIndex];
    var date = new Date(record[dateColIndex]);
    var score = record[actualScoreColIndex];
    
    if (activityMap[activityId]) {
      activityMap[activityId].records.push({
        date: date,
        score: score
      });
    }
  });
  
  // 第四步：调整同级同父活动的权重
  function adjustSiblingWeights(parentActivity) {
    if (!parentActivity.children || parentActivity.children.length === 0) {
      return;
    }
    
    // 计算同级活动的总权重
    var totalWeight = 0;
    parentActivity.children.forEach(function(child) {
      totalWeight += child.baseWeight;
    });
    
    // 根据总权重调整每个活动的权重
    if (totalWeight === 0) {
      // 如果总权重为0，则平均分配
      var equalWeight = 1.0 / parentActivity.children.length;
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = equalWeight;
      });
    } else if (totalWeight !== 100) {
      // 不为100%时，按比例调整到100%
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = child.baseWeight / totalWeight;
      });
    } else {
      // 正好是100%时，直接使用原始权重的百分比形式
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = child.baseWeight / 100;
      });
    }
    
    // 递归处理每个子活动的子活动
    parentActivity.children.forEach(function(child) {
      adjustSiblingWeights(child);
    });
  }
  
  // 调整一级活动权重
  var totalLevel1Weight = 0;
  level1Activities.forEach(function(activity) {
    totalLevel1Weight += activity.baseWeight;
  });
  
  if (totalLevel1Weight === 0) {
    // 如果总权重为0，则平均分配
    var equalLevel1Weight = 1.0 / level1Activities.length;
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = equalLevel1Weight;
    });
  } else if (totalLevel1Weight !== 100) {
    // 不为100%时，按比例调整到100%
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = activity.baseWeight / totalLevel1Weight;
    });
  } else {
    // 正好是100%时，直接使用原始权重的百分比形式
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = activity.baseWeight / 100;
    });
  }
  
  // 递归调整所有级别的活动权重
  level1Activities.forEach(function(activity) {
    adjustSiblingWeights(activity);
  });
  
  // 第五步：计算每个活动的加权平均分
  function calculateActivityScore(activity) {
    // 如果活动有记录，计算其时间加权平均分
    if (activity.records && activity.records.length > 0) {
      var currentDate = new Date();
      var totalWeight = 0;
      var weightedSum = 0;
      
      activity.records.forEach(function(record) {
        var daysDifference = Math.floor((currentDate - record.date) / (1000 * 60 * 60 * 24));
        
        // 跳过超过60天的记录
        if (daysDifference > 60) return;
        
        // 计算时间衰减系数
        var decayFactor = 1.0;
        if (activity.enableDecay) {
          decayFactor = Math.max(0, 1 - (daysDifference * activity.decayRate));
        }
        
        // 权重为衰减系数
        var weight = decayFactor;
        totalWeight += weight;
        weightedSum += record.score * weight;
      });
      
      // 计算加权平均分
      return totalWeight > 0 ? weightedSum / totalWeight : 0;
    }
    
    // 如果该活动没有直接记录，但有子活动，计算子活动的加权分数
    if (activity.children && activity.children.length > 0) {
      var totalChildScore = 0;
      
      activity.children.forEach(function(child) {
        var childScore = calculateActivityScore(child);
        totalChildScore += childScore * child.adjustedWeight;
      });
      
      return totalChildScore;
    }
    
    // 如果既没有记录也没有子活动，返回0
    return 0;
  }
  
  // 第六步：计算总分
  var totalScore = 0;
  
  // 计算每个一级活动的分数，然后应用一级活动的权重
  level1Activities.forEach(function(activity) {
    var activityScore = calculateActivityScore(activity);
    totalScore += activityScore * activity.adjustedWeight;
  });
  
  // 为了诊断，可以在返回前记录各级活动的权重
  // level1Activities.forEach(function(activity) {
  //   Logger.log("一级活动: " + activity.name + ", 权重: " + (activity.adjustedWeight * 100) + "%");
  //   activity.children.forEach(function(child) {
  //     Logger.log("-- 二级活动: " + child.name + ", 权重: " + (child.adjustedWeight * 100) + "%");
  //     child.children.forEach(function(grandchild) {
  //       Logger.log("---- 三级活动: " + grandchild.name + ", 权重: " + (grandchild.adjustedWeight * 100) + "%");
  //     });
  //   });
  // });
  
  return totalScore;
}

/**
 * 辅助函数 - 打印活动层级结构与权重 (调试用)
 */
function debugPrintActivityStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activitiesSheet = ss.getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var activityIdIdx = activitiesHeaders.indexOf('ActivityID');
  var nameIdx = activitiesHeaders.indexOf('ActivityName');
  var levelIdx = activitiesHeaders.indexOf('Level');
  var parentIdIdx = activitiesHeaders.indexOf('ParentID');
  var baseWeightIdx = activitiesHeaders.indexOf('BaseWeight');
  
  // 构建活动树结构
  var activityMap = {};
  var level1Activities = [];
  
  // 映射所有活动
  for (var i = 0; i < activitiesData.length; i++) {
    var activityId = activitiesData[i][activityIdIdx];
    var name = activitiesData[i][nameIdx];
    var level = activitiesData[i][levelIdx];
    var parentId = activitiesData[i][parentIdIdx];
    
    activityMap[activityId] = {
      id: activityId,
      name: name,
      level: level,
      parentId: parentId,
      baseWeight: activitiesData[i][baseWeightIdx],
      children: []
    };
    
    if (level == 1) {
      level1Activities.push(activityMap[activityId]);
    }
  }
  
  // 构建活动树
  for (var id in activityMap) {
    var activity = activityMap[id];
    if (activity.level > 1 && activity.parentId && activityMap[activity.parentId]) {
      activityMap[activity.parentId].children.push(activity);
    }
  }
  
  // 调整同级同父活动的权重
  function adjustSiblingWeights(parentActivity) {
    if (!parentActivity.children || parentActivity.children.length === 0) {
      return;
    }
    
    // 计算同级活动的总权重
    var totalWeight = 0;
    parentActivity.children.forEach(function(child) {
      totalWeight += child.baseWeight;
    });
    
    // 根据总权重调整每个活动的权重
    if (totalWeight === 0) {
      // 如果总权重为0，则平均分配
      var equalWeight = 1.0 / parentActivity.children.length;
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = equalWeight;
      });
    } else if (totalWeight !== 100) {
      // 不为100%时，按比例调整到100%
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = child.baseWeight / totalWeight;
      });
    } else {
      // 正好是100%时，直接使用原始权重的百分比形式
      parentActivity.children.forEach(function(child) {
        child.adjustedWeight = child.baseWeight / 100;
      });
    }
    
    // 递归处理每个子活动的子活动
    parentActivity.children.forEach(function(child) {
      adjustSiblingWeights(child);
    });
  }
  
  // 调整一级活动权重
  var totalLevel1Weight = 0;
  level1Activities.forEach(function(activity) {
    totalLevel1Weight += activity.baseWeight;
  });
  
  if (totalLevel1Weight === 0) {
    // 如果总权重为0，则平均分配
    var equalLevel1Weight = 1.0 / level1Activities.length;
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = equalLevel1Weight;
    });
  } else if (totalLevel1Weight !== 100) {
    // 不为100%时，按比例调整到100%
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = activity.baseWeight / totalLevel1Weight;
    });
  } else {
    // 正好是100%时，直接使用原始权重的百分比形式
    level1Activities.forEach(function(activity) {
      activity.adjustedWeight = activity.baseWeight / 100;
    });
  }
  
  // 递归调整所有级别的活动权重
  level1Activities.forEach(function(activity) {
    adjustSiblingWeights(activity);
  });
  
  // 打印活动层级结构与权重
  var debugOutput = "活动层级结构与权重:\n\n";
  
  level1Activities.forEach(function(activity) {
    debugOutput += "一级活动: " + activity.name + " (原权重: " + activity.baseWeight + 
                   ", 调整后: " + (activity.adjustedWeight * 100).toFixed(2) + "%)\n";
    
    if (activity.children && activity.children.length > 0) {
      activity.children.forEach(function(child) {
        debugOutput += "  ├─ 二级活动: " + child.name + " (原权重: " + child.baseWeight + 
                       ", 调整后: " + (child.adjustedWeight * 100).toFixed(2) + "%)\n";
        
        if (child.children && child.children.length > 0) {
          child.children.forEach(function(grandchild, idx, arr) {
            var prefix = idx === arr.length - 1 ? "  │  └─ " : "  │  ├─ ";
            debugOutput += prefix + "三级活动: " + grandchild.name + " (原权重: " + grandchild.baseWeight + 
                           ", 调整后: " + (grandchild.adjustedWeight * 100).toFixed(2) + "%)\n";
          });
        }
      });
    }
    
    debugOutput += "\n";
  });
  
  // 计算各级活动权重校验
  var level1WeightSum = 0;
  level1Activities.forEach(function(activity) {
    level1WeightSum += activity.adjustedWeight;
  });
  
  debugOutput += "一级活动权重总和: " + (level1WeightSum * 100).toFixed(2) + "% (应为100%)\n\n";
  
  // 检查每个父活动的子活动权重总和
  level1Activities.forEach(function(activity) {
    if (activity.children && activity.children.length > 0) {
      var childWeightSum = 0;
      activity.children.forEach(function(child) {
        childWeightSum += child.adjustedWeight;
      });
      
      debugOutput += activity.name + " 的二级活动权重总和: " + (childWeightSum * 100).toFixed(2) + "% (应为100%)\n";
      
      // 检查每个二级活动的子活动权重总和
      activity.children.forEach(function(child) {
        if (child.children && child.children.length > 0) {
          var grandchildWeightSum = 0;
          child.children.forEach(function(grandchild) {
            grandchildWeightSum += grandchild.adjustedWeight;
          });
          
          debugOutput += "  " + child.name + " 的三级活动权重总和: " + (grandchildWeightSum * 100).toFixed(2) + "% (应为100%)\n";
        }
      });
    }
  });
  
  // 显示结果
  Logger.log(debugOutput);
  
  return debugOutput;
}


/**
 * Recursively distribute points to children activities
 * @param {Object} activity - The parent activity
 * @param {Object} activityMap - Map of all activities
 */
function distributePointsToChildren(activity, activityMap) {
  if (!activity.children || activity.children.length === 0) {
    return;
  }
  
  // Calculate total weight of children
  var totalChildWeight = 0;
  for (var i = 0; i < activity.children.length; i++) {
    totalChildWeight += activity.children[i].baseWeight;
  }
  
  // If total weight is 0, avoid division by zero
  if (totalChildWeight === 0) totalChildWeight = 1;
  
  // Distribute parent's points to children
  for (var i = 0; i < activity.children.length; i++) {
    var child = activity.children[i];
    child.basePoints = activity.basePoints * (child.baseWeight / totalChildWeight);
    
    // Update in the map
    activityMap[child.id].basePoints = child.basePoints;
    
    // Recursively process this child's children
    distributePointsToChildren(child, activityMap);
  }
}

/**
 * Save participation records with updated scoring logic
 * Replace this function in AllianceManagementSystem.gs
 */
/**
 * 修改活动记录保存函数，仅计算单次活动分数，不需要考虑权重分配
 */
function saveParticipationRecords(records) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var timestamp = new Date();
  
  var activitiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var activityIdIdx = activitiesHeaders.indexOf('ActivityID');
  var typeIdx = activitiesHeaders.indexOf('Type');
  var minThresholdIdx = activitiesHeaders.indexOf('MinThreshold');
  var maxThresholdIdx = activitiesHeaders.indexOf('MaxThreshold');
  var lowScoreFactorIdx = activitiesHeaders.indexOf('LowScoreFactor');
  var highScoreFactorIdx = activitiesHeaders.indexOf('HighScoreFactor');
  
  // 创建活动属性映射
  var activityMap = {};
  for (var i = 0; i < activitiesData.length; i++) {
    var row = activitiesData[i];
    activityMap[row[activityIdIdx]] = {
      type: row[typeIdx],
      minThreshold: row[minThresholdIdx],
      maxThreshold: row[maxThresholdIdx],
      lowScoreFactor: row[lowScoreFactorIdx],
      highScoreFactor: row[highScoreFactorIdx]
    };
  }
  
  // 获取出席系数配置
  var weightConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WeightConfig');
  var weightData = weightConfigSheet.getDataRange().getValues();
  var attendanceMult = 1.0; // 默认
  var absentExcusedMult = 0.5; // 默认
  var absentUnexcusedMult = 0.0; // 默认
  
  // 查找出席系数配置
  for (var i = 1; i < weightData.length; i++) {
    var configType = weightData[i][0];
    if (configType === "ActivityAttendance") {
      attendanceMult = weightData[i][1];
    } else if (configType === "AbsentExcused") {
      absentExcusedMult = weightData[i][1];
    } else if (configType === "AbsentUnexcused") {
      absentUnexcusedMult = weightData[i][1];
    }
  }
  
  var newRecords = [];
  
  records.forEach(function(record, index) {
    var recordId = "REC" + (lastRow + index);
    var activity = activityMap[record.activityId];
    
    if (!activity) {
      console.error("Activity not found: " + record.activityId);
      return;
    }
    
    var milestoneRating = "";
    var actualScore = 0;
    
    // 计算里程碑评级和实际得分
    if (activity.type === "Score") {
      var score = parseFloat(record.score) || 0;
      
      // 修改: 从0到最高里程碑计算分数
      if (score < activity.minThreshold) {
        milestoneRating = "Below";
        // 计算低于最低里程碑的分数：按照与最低里程碑的比例计算
        // 例如：最低里程碑是50，得分是25，则得到25/50=50%的基础分
        var proportionOfMin = (activity.minThreshold > 0) ? 
                              (score / activity.minThreshold) : 0;
        
        // 基础分数是按比例计算的，然后应用低分系数
        var baseScore = 1000 * (activity.minThreshold / activity.maxThreshold) * proportionOfMin;
        actualScore = baseScore * (1 + activity.lowScoreFactor);
      } else if (score > activity.maxThreshold) {
        milestoneRating = "Above";
        // 超过最高里程碑的分数：基础分为1000，再应用高分系数
        actualScore = 1000 * (1 + activity.highScoreFactor);
      } else {
        milestoneRating = "Within";
        // 在里程碑范围内：0到最高里程碑的线性比例
        actualScore = 1000 * (score / activity.maxThreshold);
      }
    } else { // 出席型活动
      var status = record.status;
      milestoneRating = "N/A";
      
      // 根据出席状态计算分数
      if (status === "Present") {
        actualScore = 1000 * attendanceMult;
      } else if (status === "Absent-Excused") {
        actualScore = 1000 * absentExcusedMult;
      } else { // Absent-Unexcused或其他
        actualScore = 1000 * absentUnexcusedMult;
      }
    }
    
    // 创建新记录
    newRecords.push([
      recordId,                    // RecordID
      record.activityId,           // ActivityID
      new Date(record.date),       // Date
      record.memberId,             // MemberID
      record.status,               // ParticipationStatus
      record.score || "",          // Score
      milestoneRating,             // MilestoneRating
      actualScore,                 // ActualScore
      timestamp,                   // RecordTime
      record.notes || ""           // Notes
    ]);
  });
  
  // 一次性写入所有新记录
  if (newRecords.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRecords.length, 10).setValues(newRecords);
  }
  
  return newRecords.length;
}

// Keep the original generateAllRankSuggestions function
// 修改後的generateAllRankSuggestions函數 - 移除出席率檢查
function generateAllRankSuggestions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var membersSheet = ss.getSheetByName('Members');
  var membersData = membersSheet.getDataRange().getValues();
  var headers = membersData.shift(); // Remove headers
  
  // Find column indices
  var memberIdColIndex = headers.indexOf('MemberID');
  var rankColIndex = headers.indexOf('Rank');
  var totalScoreColIndex = headers.indexOf('TotalScore');
  var rankSuggestionColIndex = headers.indexOf('RankSuggestion');
  
  // Skip if required columns not found
  if (memberIdColIndex === -1 || rankColIndex === -1 || 
      totalScoreColIndex === -1 || rankSuggestionColIndex === -1) {
    Browser.msgBox('Error', 'Required columns not found in Members sheet.', Browser.Buttons.OK);
    return;
  }
  
  // Get evaluation criteria
  var criteriaSheet = ss.getSheetByName('EvaluationCriteria');
  var criteriaData = criteriaSheet.getDataRange().getValues();
  var criteriaHeaders = criteriaData.shift();
  
  var evalTypeIdx = criteriaHeaders.indexOf('EvaluationType');
  var fromRankIdx = criteriaHeaders.indexOf('FromRank');
  var toRankIdx = criteriaHeaders.indexOf('ToRank');
  var totalScoreReqIdx = criteriaHeaders.indexOf('TotalScoreRequired');
  
  // Process each member
  var suggestionCount = 0;
  
  for (var i = 0; i < membersData.length; i++) {
    var row = membersData[i];
    var memberId = row[memberIdColIndex];
    var currentRank = row[rankColIndex];
    var totalScore = row[totalScoreColIndex];
    
    // Skip inactive members
    if (currentRank === 'X') continue;
    
    // Skip R4 and R5 members (management, manually assigned)
    if (currentRank === 'R4' || currentRank === 'R5') continue;
    
    // 如需調試特定成員，可取消下面注釋
    /*
    if (memberId === "M0092" || memberId === "M0100") {
      Logger.log("評估成員: " + memberId + ", 等級: " + currentRank + ", 總分: " + totalScore);
    }
    */
    
    // Find applicable criteria
    var suggestion = '';
    
    // Check promotion criteria
    for (var j = 0; j < criteriaData.length; j++) {
      if (criteriaData[j][evalTypeIdx] === 'Promotion' && 
          criteriaData[j][fromRankIdx] === currentRank) {
        
        var requiredScore = criteriaData[j][totalScoreReqIdx];
        
        // 移除出席率檢查，僅根據總分判斷晉升
        if (totalScore >= requiredScore) {
          suggestion = '晉升至 ' + criteriaData[j][toRankIdx];
          break;
        }
      }
    }
    
    // If no promotion, check demotion criteria
    if (suggestion === '') {
      for (var j = 0; j < criteriaData.length; j++) {
        if (criteriaData[j][evalTypeIdx] === 'Demotion' && 
            criteriaData[j][fromRankIdx] === currentRank) {
          
          var requiredScore = criteriaData[j][totalScoreReqIdx];
          
          // 移除出席率檢查，僅根據總分判斷降級
          if (totalScore < requiredScore) {
            suggestion = '降級至 ' + criteriaData[j][toRankIdx];
            
            // 如需調試特定成員，可取消下面注釋
            /*
            if (memberId === "M0092" || memberId === "M0100") {
              Logger.log("符合降級條件: 總分" + totalScore + " < 要求分數" + requiredScore);
            }
            */
            
            break;
          }
          // 如需調試特定成員，可取消下面注釋
          /*
          else if (memberId === "M0092" || memberId === "M0100") {
            Logger.log("不符合降級條件: 總分" + totalScore + " >= 要求分數" + requiredScore);
          }
          */
        }
      }
    }
    
    // If R1 and meets removal criteria
    if (currentRank === 'R1' && suggestion === '') {
      for (var j = 0; j < criteriaData.length; j++) {
        if (criteriaData[j][evalTypeIdx] === 'Removal' && 
            criteriaData[j][fromRankIdx] === 'R1') {
          
          var requiredScore = criteriaData[j][totalScoreReqIdx];
          
          // 移除出席率檢查，僅根據總分判斷移除
          if (totalScore < requiredScore) {
            suggestion = '從聯盟中移除';
            break;
          }
        }
      }
    }
    
    // Update suggestion in Members sheet
    if (suggestion !== '') {
      membersSheet.getRange(i + 2, rankSuggestionColIndex + 1).setValue(suggestion);
      suggestionCount++;
    } else {
      membersSheet.getRange(i + 2, rankSuggestionColIndex + 1).setValue('無變更');
    }
  }
  
  Browser.msgBox('成功', '已生成等級建議。找到 ' + suggestionCount + ' 位成員有變更建議。', Browser.Buttons.OK);
  
  // Update dashboard with new suggestions
  updateDashboard();
}

// Keep the original calculateAttendanceRate function
function calculateAttendanceRate(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var participationSheet = ss.getSheetByName('Participation');
  var participationData = participationSheet.getDataRange().getValues();
  var headers = participationData.shift();
  
  var memberIdIdx = headers.indexOf('MemberID');
  var activityIdIdx = headers.indexOf('ActivityID');
  var statusIdx = headers.indexOf('ParticipationStatus');
  var dateIdx = headers.indexOf('Date');
  
  // Get activities to filter only attendance type activities
  var activitiesSheet = ss.getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  var activityIdColIdx = activitiesHeaders.indexOf('ActivityID');
  var typeColIdx = activitiesHeaders.indexOf('Type');
  
  var attendanceActivities = {};
  activitiesData.forEach(function(row) {
    if (row[typeColIdx] === 'Attendance') {
      attendanceActivities[row[activityIdColIdx]] = true;
    }
  });
  
  // Filter and process last 30 days of attendance records
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  var attendanceRecords = participationData.filter(function(row) {
    return row[memberIdIdx] === memberId && 
           attendanceActivities[row[activityIdIdx]] &&
           new Date(row[dateIdx]) >= thirtyDaysAgo;
  });
  
  if (attendanceRecords.length === 0) {
    return 0; // No attendance records in the last 30 days
  }
  
  var presentCount = 0;
  attendanceRecords.forEach(function(record) {
    if (record[statusIdx] === 'Present') {
      presentCount++;
    }
  });
  
  return (presentCount / attendanceRecords.length) * 100;
}

// Keep the original updateDashboard function
// Modified updateDashboard function with low score alerts (in English)
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName('Dashboard');
  var membersSheet = ss.getSheetByName('Members');
  var membersData = membersSheet.getDataRange().getValues();
  var headers = membersData.shift();
  
  var rankColIndex = headers.indexOf('Rank');
  var lastActiveColIndex = headers.indexOf('LastActiveDate');
  var totalScoreColIndex = headers.indexOf('TotalScore');
  var gameNameColIndex = headers.indexOf('GameName');
  var rankSuggestionColIndex = headers.indexOf('RankSuggestion');
  
  // Count members by rank
  var counts = {
    total: 0,
    R1: 0,
    R2: 0,
    R3: 0,
    R4: 0,
    R5: 0,
    inactive: 0,
    lowScore: 0 // New: track members with low scores
  };
  
  var today = new Date();
  var fourteenDaysAgo = new Date();
  fourteenDaysAgo.setDate(today.getDate() - 14);
  
  // List of low-scoring members
  var lowScoreMembers = [];
  
  membersData.forEach(function(row) {
    var rank = row[rankColIndex];
    var lastActive = row[lastActiveColIndex] ? new Date(row[lastActiveColIndex]) : null;
    var totalScore = row[totalScoreColIndex];
    var memberName = row[gameNameColIndex];
    
    // Skip inactive members
    if (rank === 'X') return;
    
    counts.total++;
    
    // Count by rank
    if (rank === 'R1') counts.R1++;
    else if (rank === 'R2') counts.R2++;
    else if (rank === 'R3') counts.R3++;
    else if (rank === 'R4') counts.R4++;
    else if (rank === 'R5') counts.R5++;
    
    // Count inactive members (not active in last 14 days)
    if (lastActive && lastActive < fourteenDaysAgo) {
      counts.inactive++;
    }
    
    // New: Check for members with scores below 700
    if (totalScore < 700) {
      counts.lowScore++;
      lowScoreMembers.push(memberName + ' (' + rank + ', ' + totalScore + ')');
    }
  });
  
  // Update dashboard with counts
  dashboardSheet.getRange('B4').setValue(counts.total);
  dashboardSheet.getRange('B5').setValue(counts.R1);
  dashboardSheet.getRange('B6').setValue(counts.R2);
  dashboardSheet.getRange('B7').setValue(counts.R3);
  dashboardSheet.getRange('B8').setValue(counts.R4);
  dashboardSheet.getRange('B9').setValue(counts.R5);
  dashboardSheet.getRange('B10').setValue(counts.inactive);
  
  // Update alerts section
  var alerts = [];
  if (counts.inactive > 0) {
    alerts.push(counts.inactive + ' 位成員超過14天未活躍');
  }
  
  // New: Add low score members alert
  if (counts.lowScore > 0) {
    alerts.push(counts.lowScore + ' 位成員總分低於700');
    
    // If there are 5 or fewer low-scoring members, list their names
    if (lowScoreMembers.length <= 5) {
      alerts.push('低分成員: ' + lowScoreMembers.join(', '));
    }
  }
  
  // Display alerts
  if (alerts.length > 0) {
    dashboardSheet.getRange('A13:G13').merge();
    dashboardSheet.getRange('A13').setValue(alerts.join('\n'));
    
    // If there are multiple alerts, increase the row height
    if (alerts.length > 1) {
      dashboardSheet.setRowHeight(13, 20 * alerts.length);
    }
  } else {
    dashboardSheet.getRange('A13:G13').merge();
    dashboardSheet.getRange('A13').setValue('No critical alerts at this time.');
    dashboardSheet.setRowHeight(13, 20); // Reset row height
  }
  
  // Display rank suggestions
  var suggestions = {
    promote: [],
    demote: [],
    remove: []
  };
  
  membersData.forEach(function(row) {
    var rank = row[rankColIndex];
    var suggestion = row[rankSuggestionColIndex];
    var memberName = row[gameNameColIndex];
    
    // Skip inactive members
    if (rank === 'X') return;
    
    // Process suggestions
    if (suggestion && suggestion.indexOf('Promote') === 0) {
      suggestions.promote.push(memberName + ' (' + rank + ' → ' + suggestion.split(' ').pop() + ')');
    } else if (suggestion && suggestion.indexOf('Demote') === 0) {
      suggestions.demote.push(memberName + ' (' + rank + ' → ' + suggestion.split(' ').pop() + ')');
    } else if (suggestion && suggestion.indexOf('Remove') === 0) {
      suggestions.remove.push(memberName + ' (' + rank + ')');
    }
  });
  
  // Display rank suggestions
  var suggestionText = '';
  
  if (suggestions.promote.length > 0) {
    suggestionText += '晉升建議:\n' + suggestions.promote.join('\n') + '\n\n';
  }
  
  if (suggestions.demote.length > 0) {
    suggestionText += '降級建議:\n' + suggestions.demote.join('\n') + '\n\n';
  }
  
  if (suggestions.remove.length > 0) {
    suggestionText += '移除建議:\n' + suggestions.remove.join('\n');
  }
  
  if (suggestionText === '') {
    suggestionText = '目前沒有等級變更建議。';
  }
  
  dashboardSheet.getRange('A16:G25').merge();
  dashboardSheet.getRange('A16:G25').setValue(suggestionText);
}

// Keep the original setupTriggers function
function setupTriggers() {
  // Delete all existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create daily trigger to calculate scores
  ScriptApp.newTrigger('calculateAllMemberScores')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  
  // Create weekly trigger to generate rank suggestions
  ScriptApp.newTrigger('generateAllRankSuggestions')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
  
  // Create daily trigger to update dashboard
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  
  Browser.msgBox('成功', '自動觸發器已設置:\n\n' +
                '每日 (凌晨1:00): 計算成員分數\n' +
                '每週 (週日, 凌晨2:00): 生成等級建議\n' +
                '每日 (凌晨3:00): 更新儀表板', 
                Browser.Buttons.OK);
}

/**
 * 修复计算成员分数函数 - 添加缺失的findTopParentId函数
 */
function findTopParentId(activityId, activityMap) {
  var activity = activityMap[activityId];
  if (!activity) return null;
  
  if (activity.level === 1) {
    return activity.id;
  }
  
  if (activity.parentId && activityMap[activity.parentId]) {
    return findTopParentId(activity.parentId, activityMap);
  }
  
  return activityId; // 如果找不到父活动，返回自身ID
}

/**
 * 显示活动结构和权重
 */
function showActivityStructure() {
  var structure = debugPrintActivityStructure();
  var ui = SpreadsheetApp.getUi();
  
  // 由于消息框有大小限制，可能需要分段显示
  var maxLength = 500; // 消息框大约能显示的字符数
  
  if (structure.length <= maxLength) {
    ui.alert('活动结构与权重', structure, ui.ButtonSet.OK);
  } else {
    // 分段显示
    var parts = Math.ceil(structure.length / maxLength);
    for (var i = 0; i < parts; i++) {
      var start = i * maxLength;
      var end = Math.min((i + 1) * maxLength, structure.length);
      var part = structure.substring(start, end);
      
      ui.alert('活动结构与权重 (部分 ' + (i+1) + '/' + parts + ')', part, ui.ButtonSet.OK);
    }
  }
}

/**
 * Alliance Management System - Member Management Module
 * Functions to add to AllianceManagementSystem.gs


// Update onOpen function to include member management
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alliance System')
    .addItem('Record Activity', 'showActivityRecordingSidebar')
    .addItem('Manage Members', 'showMemberManagementSidebar')
    .addSeparator()
    .addItem('Calculate Scores', 'calculateAllMemberScores')
    .addItem('Generate Rank Suggestions', 'generateAllRankSuggestions')
    .addItem('Update Dashboard', 'updateDashboard')
    .addSeparator()
    .addItem('Setup System', 'runSetup')
    .addItem('Setup Triggers', 'setupTriggers')
    .addToUi();
}
 */

/**
 * Opens the member management dialog
 */
function showMemberManagementDialog() {
  var html = HtmlService.createTemplateFromFile('MembersDialog')
    .evaluate()
    .setWidth(900)
    .setHeight(600)
    .setTitle('Member Management');
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Management');
}

/**
 * @deprecated 使用showMemberManagementDialog()取代
 * 保留此函數用於向後兼容
 */
function showMemberManagementSidebar() {
  // 轉向使用對話框版本
  showMemberManagementDialog();
  
  // 顯示棄用提示
//  SpreadsheetApp.getActive().toast(
//    '側邊欄版本已棄用。現在使用對話框版本。',
//    '提示',
//    5
//  );
}

/**
 * Get all members for the sidebar
 */
function getAllMembers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var members = [];
  
  data.forEach(function(row) {
    var member = {};
    headers.forEach(function(header, index) {
      // Format dates for proper display in the UI
      if (header === 'JoinDate' || header === 'LastActiveDate') {
        member[header] = row[index] ? Utilities.formatDate(new Date(row[index]), 
                                                          Session.getScriptTimeZone(), 
                                                          'yyyy-MM-dd') : '';
      } else {
        member[header] = row[index];
      }
    });
    members.push(member);
  });
  
  return {
    members: members,
    headers: headers
  };
}

/**
 * Add a new member
 */
function addNewMember(memberData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  // Generate a unique MemberID
  var lastRow = data.length > 0 ? data[data.length-1] : null;
  var lastMemberId = lastRow ? lastRow[headers.indexOf('MemberID')] : 'M0000';
  var numericPart = parseInt(lastMemberId.substring(1), 10);
  var newMemberId = 'M' + String(numericPart + 1).padStart(4, '0');
  
  // Set default values for some fields
  memberData.MemberID = newMemberId;
  memberData.JoinDate = memberData.JoinDate || new Date();
  memberData.LastActiveDate = memberData.LastActiveDate || new Date();
  memberData.TotalScore = memberData.TotalScore || 0;
  memberData.RankSuggestion = memberData.RankSuggestion || 'No change';
  
  // Create row array in the correct order
  var newRow = headers.map(function(header) {
    if (header === 'JoinDate' || header === 'LastActiveDate') {
      return new Date(memberData[header]);
    }
    return memberData[header] || '';
  });
  
  // Add new row
  sheet.appendRow(newRow);
  
  // Update dashboard
  updateDashboard();
  
  return newMemberId;
}

/**
 * Update a member's information
 */
function updateMemberInfo(memberData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var memberId = memberData.MemberID;
  
  // Find the row index for the member
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][memberIdIdx] === memberId) {
      rowIndex = i + 2; // +2 because we removed headers and sheet is 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Member not found: ' + memberId);
  }
  
  // Update the row with new data
  headers.forEach(function(header, colIndex) {
    if (memberData.hasOwnProperty(header)) {
      var value = memberData[header];
      
      // Convert date strings to Date objects
      if (header === 'JoinDate' || header === 'LastActiveDate') {
        value = value ? new Date(value) : '';
      }
      
      sheet.getRange(rowIndex, colIndex + 1).setValue(value);
    }
  });
  
  // Update dashboard
  updateDashboard();
  
  return true;
}

/**
 * Delete a member (set to inactive)
 */
function deleteMember(memberId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var rankIdx = headers.indexOf('Rank');
  
  // Find the row index for the member
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][memberIdIdx] === memberId) {
      rowIndex = i + 2; // +2 because we removed headers and sheet is 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Member not found: ' + memberId);
  }
  
  // Set the rank to 'X' (inactive) instead of deleting
  sheet.getRange(rowIndex, rankIdx + 1).setValue('X');
  
  // Update dashboard
  updateDashboard();
  
  return true;
}

/**
 * Get filtered and sorted members
 */
function getFilteredMembers(filterOptions) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var gameNameIdx = headers.indexOf('GameName');
  var rankIdx = headers.indexOf('Rank');
  var powerIdx = headers.indexOf('Power');
  var lastActiveIdx = headers.indexOf('LastActiveDate');
  
  var members = [];
  
  // Convert data to objects
  data.forEach(function(row) {
    var member = {};
    headers.forEach(function(header, index) {
      // Format dates
      if (header === 'JoinDate' || header === 'LastActiveDate') {
        member[header] = row[index] ? Utilities.formatDate(new Date(row[index]), 
                                                          Session.getScriptTimeZone(), 
                                                          'yyyy-MM-dd') : '';
      } else {
        member[header] = row[index];
      }
    });
    members.push(member);
  });
  
  // Apply filters
  if (filterOptions) {
    // Filter by rank
    if (filterOptions.rank && filterOptions.rank !== 'All') {
      members = members.filter(function(member) {
        return member.Rank === filterOptions.rank;
      });
    }
    
    // Filter by search term
    if (filterOptions.searchTerm) {
      var term = filterOptions.searchTerm.toLowerCase();
      members = members.filter(function(member) {
        return member.GameName.toLowerCase().includes(term) || 
               member.MemberID.toLowerCase().includes(term);
      });
    }
    
    // Filter inactive members
    if (filterOptions.hideInactive) {
      members = members.filter(function(member) {
        return member.Rank !== 'X';
      });
    }
    
    // Sort by specified field
    if (filterOptions.sortBy) {
      var sortField = filterOptions.sortBy;
      var sortDir = filterOptions.sortDirection || 'asc';
      
      members.sort(function(a, b) {
        var aValue = a[sortField];
        var bValue = b[sortField];
        
        // Handle numerical values
        if (sortField === 'Power' || sortField === 'TotalScore') {
          aValue = parseFloat(aValue) || 0;
          bValue = parseFloat(bValue) || 0;
        }
        
        // Handle dates
        if (sortField === 'JoinDate' || sortField === 'LastActiveDate') {
          aValue = aValue ? new Date(aValue) : new Date(0);
          bValue = bValue ? new Date(bValue) : new Date(0);
        }
        
        // Special handling for Rank
        if (sortField === 'Rank') {
          // Custom rank ordering (R5 > R4 > R3 > R2 > R1 > X)
          var rankOrder = { 'R5': 5, 'R4': 4, 'R3': 3, 'R2': 2, 'R1': 1, 'X': 0 };
          aValue = rankOrder[aValue] || 0;
          bValue = rankOrder[bValue] || 0;
        }
        
        if (sortDir === 'asc') {
          return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
        } else {
          return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
        }
      });
    }
  }
  
  return members;
}

/**
 * Update a member's power
 */
function updateMemberPower(memberId, power) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var powerIdx = headers.indexOf('Power');
  var lastActiveIdx = headers.indexOf('LastActiveDate');
  
  // Find the row index for the member
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][memberIdIdx] === memberId) {
      rowIndex = i + 2; // +2 because we removed headers and sheet is 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Member not found: ' + memberId);
  }
  
  // Update power
  sheet.getRange(rowIndex, powerIdx + 1).setValue(power);
  
  // Update last active date to today
  sheet.getRange(rowIndex, lastActiveIdx + 1).setValue(new Date());
  
  // Recalculate the member's score
  calculateMemberScore(memberId);
  
  return true;
}

/**
 * Set a member's rank (manual override)
 */
function setMemberRank(memberId, newRank) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var rankIdx = headers.indexOf('Rank');
  var rankSuggestionIdx = headers.indexOf('RankSuggestion');
  
  // Find the row index for the member
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][memberIdIdx] === memberId) {
      rowIndex = i + 2; // +2 because we removed headers and sheet is 1-indexed
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('Member not found: ' + memberId);
  }
  
  // Update rank
  sheet.getRange(rowIndex, rankIdx + 1).setValue(newRank);
  
  // Clear rank suggestion
  sheet.getRange(rowIndex, rankSuggestionIdx + 1).setValue('No change');
  
  // Update dashboard
  updateDashboard();
  
  return true;
}

/**
 * Bulk update members' power
 */
function bulkUpdatePower(powerUpdates) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var powerIdx = headers.indexOf('Power');
  var lastActiveIdx = headers.indexOf('LastActiveDate');
  
  var updatedCount = 0;
  
  for (var i = 0; i < powerUpdates.length; i++) {
    var memberId = powerUpdates[i].memberId;
    var power = powerUpdates[i].power;
    
    // Find the row index for the member
    var rowIndex = -1;
    for (var j = 0; j < data.length; j++) {
      if (data[j][memberIdIdx] === memberId) {
        rowIndex = j + 2; // +2 because we removed headers and sheet is 1-indexed
        break;
      }
    }
    
    if (rowIndex === -1) continue; // Skip if member not found
    
    // Update power
    sheet.getRange(rowIndex, powerIdx + 1).setValue(power);
    
    // Update last active date to today
    sheet.getRange(rowIndex, lastActiveIdx + 1).setValue(new Date());
    
    updatedCount++;
  }
  
  // Recalculate all scores
  calculateAllMemberScores();
  
  return updatedCount;
}

/**
 * Alliance Management System - Activity Recording Interface
 * Functions to add to AllianceManagementSystem.gs


// Create menu to access the system functions
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alliance System')
    .addItem('Record Activity', 'showActivityRecordingSidebar')
    .addItem('Setup System', 'runSetup')
    .addToUi();
}
 */

/**
 * Opens the activity recording sidebar
 */
function showActivityRecordingSidebar() {
  var html = HtmlService.createTemplateFromFile('ActivitySidebar')
    .evaluate()
    .setTitle('Record Activity')
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Include external HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get all activities for the cascading dropdowns
 */
function getActivities() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var activities = {
    level1: [],
    level2: {},
    level3: {}
  };
  
  // Extract data for the three-level hierarchy
  data.forEach(function(row) {
    var activityId = row[0];
    var activityName = row[1];
    var level = row[2];
    var parentId = row[3];
    var type = row[4];
    var baseWeight = row[5];
    var minThreshold = row[8] || 0;
    var maxThreshold = row[9] || 0;
    
    var activityInfo = {
      id: activityId,
      name: activityName,
      type: type,
      baseWeight: baseWeight,
      minThreshold: minThreshold,
      maxThreshold: maxThreshold
    };
    
    if (level == 1) {
      activities.level1.push(activityInfo);
    } else if (level == 2) {
      if (!activities.level2[parentId]) {
        activities.level2[parentId] = [];
      }
      activities.level2[parentId].push(activityInfo);
    } else if (level == 3) {
      if (!activities.level3[parentId]) {
        activities.level3[parentId] = [];
      }
      activities.level3[parentId].push(activityInfo);
    }
  });
  
  return activities;
}

/**
 * Get all active members for the form
 */
function getActiveMembers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var members = [];
  
  data.forEach(function(row) {
    // Skip inactive members (rank 'X')
    if (row[2] !== 'X') {
      members.push({
        id: row[0],
        name: row[1],
        rank: row[2]
      });
    }
  });
  
  // Sort by rank then name
  members.sort(function(a, b) {
    if (a.rank !== b.rank) {
      // Custom rank sorting (R5 > R4 > R3 > R2 > R1)
      return "R54321".indexOf(a.rank) - "R54321".indexOf(b.rank);
    }
    return a.name.localeCompare(b.name);
  });
  
  return members;
}

/**
 * Save participation records to the Participation sheet
 */
function saveParticipationRecords(records) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participation');
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var timestamp = new Date();
  
  var activitiesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activityMap = {};
  
  // Create a map of activity IDs to their properties
  for (var i = 1; i < activitiesData.length; i++) {
    var row = activitiesData[i];
    activityMap[row[0]] = {
      type: row[4],
      baseWeight: row[5],
      decayRate: row[6],
      enableDecay: row[7],
      minThreshold: row[8],
      maxThreshold: row[9],
      lowScoreFactor: row[10],
      highScoreFactor: row[11]
    };
  }
  
  var newRecords = [];
  
  records.forEach(function(record, index) {
    var recordId = "REC" + (lastRow + index);
    var activityProperties = activityMap[record.activityId];
    var milestoneRating = "";
    var actualScore = 0;
    
    // Calculate milestone rating and actual score
    if (activityProperties.type === "Score") {
      var score = parseFloat(record.score) || 0;
      
      // Determine milestone rating
      if (score < activityProperties.minThreshold) {
        milestoneRating = "Below";
        // Apply low score factor
        actualScore = activityProperties.baseWeight * (1 + activityProperties.lowScoreFactor);
      } else if (score > activityProperties.maxThreshold) {
        milestoneRating = "Above";
        // Apply high score factor
        actualScore = activityProperties.baseWeight * (1 + activityProperties.highScoreFactor);
      } else {
        milestoneRating = "Within";
        actualScore = activityProperties.baseWeight;
      }
    } else { // Attendance type activity
      var status = record.status;
      milestoneRating = "N/A";
      
      // Get attendance multipliers from WeightConfig
      var weightConfigSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WeightConfig');
      var weightData = weightConfigSheet.getDataRange().getValues();
      var attendanceMult = 1.0; // Default
      var absentExcusedMult = 0.5; // Default
      var absentUnexcusedMult = 0.0; // Default
      
      // Find the multipliers in WeightConfig
      for (var i = 1; i < weightData.length; i++) {
        var configType = weightData[i][0];
        if (configType === "ActivityAttendance") {
          attendanceMult = weightData[i][1];
        } else if (configType === "AbsentExcused") {
          absentExcusedMult = weightData[i][1];
        } else if (configType === "AbsentUnexcused") {
          absentUnexcusedMult = weightData[i][1];
        }
      }
      
      // Calculate actual score based on attendance status
      if (status === "Present") {
        actualScore = activityProperties.baseWeight * attendanceMult;
      } else if (status === "Absent-Excused") {
        actualScore = activityProperties.baseWeight * absentExcusedMult;
      } else { // Absent-Unexcused or other
        actualScore = activityProperties.baseWeight * absentUnexcusedMult;
      }
    }
    
    // Create the new record
    newRecords.push([
      recordId,                    // RecordID
      record.activityId,           // ActivityID
      new Date(record.date),       // Date
      record.memberId,             // MemberID
      record.status,               // ParticipationStatus
      record.score || "",          // Score
      milestoneRating,             // MilestoneRating
      actualScore,                 // ActualScore
      timestamp,                   // RecordTime
      record.notes || ""           // Notes
    ]);
  });
  
  // Write all new records at once
  if (newRecords.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRecords.length, 10).setValues(newRecords);
  }
  
  return newRecords.length;
}

/**
 * Alliance Management System - Member Ranking Module
 * Functions to add to AllianceManagementSystem.gs
 */

/**
 * Shows a ranked member list in a popup window
 */
function showRankedMemberList() {
  var html = HtmlService.createTemplateFromFile('RankedMembers')
    .evaluate()
    .setWidth(800)
    .setHeight(1000)
    .setTitle('Member Rankings');
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Rankings');
}

/**
 * Get all members sorted by total score
 */
function getRankedMembers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // Remove headers
  
  var memberIdIdx = headers.indexOf('MemberID');
  var gameNameIdx = headers.indexOf('GameName');
  var rankIdx = headers.indexOf('Rank');
  var powerIdx = headers.indexOf('Power');
  var totalScoreIdx = headers.indexOf('TotalScore');
  var joinDateIdx = headers.indexOf('JoinDate');
  var lastActiveIdx = headers.indexOf('LastActiveDate');
  
  var members = [];
  
  data.forEach(function(row) {
    // Format dates for proper display
    var joinDate = row[joinDateIdx] ? Utilities.formatDate(new Date(row[joinDateIdx]), 
                                                      Session.getScriptTimeZone(), 
                                                      'yyyy-MM-dd') : '';
    var lastActive = row[lastActiveIdx] ? Utilities.formatDate(new Date(row[lastActiveIdx]), 
                                                          Session.getScriptTimeZone(), 
                                                          'yyyy-MM-dd') : '';
    
    members.push({
      memberId: row[memberIdIdx],
      name: row[gameNameIdx],
      rank: row[rankIdx],
      power: row[powerIdx],
      totalScore: row[totalScoreIdx],
      joinDate: joinDate,
      lastActive: lastActive
    });
  });
  
  // Sort by total score (descending)
  members.sort(function(a, b) {
    return b.totalScore - a.totalScore;
  });
  
  return members;
}

/**
 * 改进版 - 重新计算所有活动记录的ActualScore，提供实时进度显示
 */
function recalculateActualScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var participationSheet = ss.getSheetByName('Participation');
  
  // 获取总行数，不包括表头
  var totalRows = participationSheet.getLastRow() - 1;
  
  // 如果没有数据，直接返回
  if (totalRows <= 0) {
    Browser.msgBox('提示', '没有找到活动记录数据', Browser.Buttons.OK);
    return;
  }
  
  // 获取表头
  var headers = participationSheet.getRange(1, 1, 1, participationSheet.getLastColumn()).getValues()[0];
  
  // 查找必要列的索引
  var activityIdIdx = headers.indexOf('ActivityID');
  var statusIdx = headers.indexOf('ParticipationStatus');
  var scoreIdx = headers.indexOf('Score');
  var actualScoreIdx = headers.indexOf('ActualScore');
  var milestoneRatingIdx = headers.indexOf('MilestoneRating');
  
  // 确保所需列存在
  if (activityIdIdx === -1 || statusIdx === -1 || scoreIdx === -1 || actualScoreIdx === -1 || milestoneRatingIdx === -1) {
    Browser.msgBox('错误', '找不到必要的列：ActivityID, ParticipationStatus, Score, ActualScore 或 MilestoneRating', Browser.Buttons.OK);
    return;
  }
  
  // 获取活动信息
  var activitiesSheet = ss.getSheetByName('Activities');
  var activitiesData = activitiesSheet.getDataRange().getValues();
  var activitiesHeaders = activitiesData.shift();
  
  // 查找活动表中必要的列
  var activityIdColIdx = activitiesHeaders.indexOf('ActivityID');
  var typeColIdx = activitiesHeaders.indexOf('Type');
  var minThresholdColIdx = activitiesHeaders.indexOf('MinThreshold');
  var maxThresholdColIdx = activitiesHeaders.indexOf('MaxThreshold');
  var lowScoreFactorColIdx = activitiesHeaders.indexOf('LowScoreFactor');
  var highScoreFactorColIdx = activitiesHeaders.indexOf('HighScoreFactor');
  
  // 确保活动表中有需要的列
  if (activityIdColIdx === -1 || typeColIdx === -1) {
    Browser.msgBox('错误', '活动表结构不正确，找不到ActivityID或Type列', Browser.Buttons.OK);
    return;
  }
  
  // 构建活动映射表
  var activityMap = {};
  for (var i = 0; i < activitiesData.length; i++) {
    var row = activitiesData[i];
    activityMap[row[activityIdColIdx]] = {
      type: row[typeColIdx],
      minThreshold: row[minThresholdColIdx] || 0,
      maxThreshold: row[maxThresholdColIdx] || 0,
      lowScoreFactor: row[lowScoreFactorColIdx] || -0.2,
      highScoreFactor: row[highScoreFactorColIdx] || 0.2
    };
  }
  
  // 获取出席状态系数
  var weightConfigSheet = ss.getSheetByName('WeightConfig');
  var weightData = weightConfigSheet.getDataRange().getValues();
  var attendanceMult = 1.0;
  var absentExcusedMult = 0.5;
  var absentUnexcusedMult = 0.0;
  
  for (var i = 1; i < weightData.length; i++) {
    var configType = weightData[i][0];
    if (configType === "ActivityAttendance") {
      attendanceMult = weightData[i][1];
    } else if (configType === "AbsentExcused") {
      absentExcusedMult = weightData[i][1];
    } else if (configType === "AbsentUnexcused") {
      absentUnexcusedMult = weightData[i][1];
    }
  }
  
  // 确认操作
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '重新计算活动分数',
    '将使用新的计分方式重新计算 ' + totalRows + ' 条记录的ActualScore。\n' +
    '新计分方式：从0到最高里程碑\n\n继续吗？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 创建进度跟踪表格
  var progressSheet = null;
  try {
    // 先检查是否已存在进度表，如果存在则删除
    if (ss.getSheetByName('RecalculationProgress')) {
      ss.deleteSheet(ss.getSheetByName('RecalculationProgress'));
    }
    
    // 创建新的进度表
    progressSheet = ss.insertSheet('RecalculationProgress');
    progressSheet.getRange('A1:D1').setValues([['处理进度', '已处理', '总计', '百分比']]);
    progressSheet.getRange('B2').setValue(0);
    progressSheet.getRange('C2').setValue(totalRows);
    progressSheet.getRange('D2').setValue('0%');
    progressSheet.getRange('A4:E4').setValues([['最近处理的记录', 'ActivityID', '原分数', '新分数', '状态']]);
    
    // 设置格式
    progressSheet.getRange('A1:D1').setFontWeight('bold');
    progressSheet.getRange('A4:E4').setFontWeight('bold');
    progressSheet.setColumnWidth(1, 150);
    progressSheet.setColumnWidth(2, 100);
    progressSheet.setColumnWidth(3, 100);
    progressSheet.setColumnWidth(4, 100);
    progressSheet.setColumnWidth(5, 150);
  } catch (e) {
    console.error('创建进度表失败: ' + e);
    // 继续执行，即使没有进度表
  }
  
  // 开始处理
  var processed = 0;
  var skipped = 0;
  var updateInterval = Math.max(1, Math.floor(totalRows / 100)); // 每处理1%的记录更新一次进度
  
  // 逐行处理数据
  for (var rowIndex = 2; rowIndex <= totalRows + 1; rowIndex++) {
    // 获取当前行数据
    var rowData = participationSheet.getRange(rowIndex, 1, 1, participationSheet.getLastColumn()).getValues()[0];
    var activityId = rowData[activityIdIdx];
    var status = rowData[statusIdx];
    var scoreValue = rowData[scoreIdx];
    var oldActualScore = rowData[actualScoreIdx];
    
    // 检查活动是否存在
    if (!activityMap[activityId]) {
      skipped++;
      // 更新进度表中的最近处理记录
      if (progressSheet) {
        progressSheet.getRange('B5').setValue(activityId);
        progressSheet.getRange('C5').setValue('N/A');
        progressSheet.getRange('D5').setValue('N/A');
        progressSheet.getRange('E5').setValue('跳过 - 未找到活动');
      }
      continue;
    }
    
    var activity = activityMap[activityId];
    var milestoneRating = '';
    var actualScore = 0;
    
    // 根据活动类型计算分数 - 使用与saveParticipationRecords相同的逻辑
    if (activity.type === 'Score') {
      var score = parseFloat(scoreValue) || 0;
      
      // 使用新的从0到最高里程碑的评分方式
      if (score < activity.minThreshold) {
        milestoneRating = "Below";
        // 计算低于最低里程碑的分数
        var proportionOfMin = (activity.minThreshold > 0) ? 
                              (score / activity.minThreshold) : 0;
        
        var baseScore = 1000 * (activity.minThreshold / activity.maxThreshold) * proportionOfMin;
        actualScore = baseScore * (1 + activity.lowScoreFactor);
      } else if (score > activity.maxThreshold) {
        milestoneRating = "Above";
        actualScore = 1000 * (1 + activity.highScoreFactor);
      } else {
        milestoneRating = "Within";
        actualScore = 1000 * (score / activity.maxThreshold);
      }
    } else { // 出席型活动
      milestoneRating = 'N/A';
      
      if (status === 'Present') {
        actualScore = 1000 * attendanceMult;
      } else if (status === 'Absent-Excused') {
        actualScore = 1000 * absentExcusedMult;
      } else { // Absent-Unexcused或其他
        actualScore = 1000 * absentUnexcusedMult;
      }
    }
    
    // 更新进度表中的最近处理记录
    if (progressSheet) {
      progressSheet.getRange('B5').setValue(activityId);
      progressSheet.getRange('C5').setValue(oldActualScore);
      progressSheet.getRange('D5').setValue(actualScore);
      progressSheet.getRange('E5').setValue('已更新');
    }
    
    // 直接更新值
    participationSheet.getRange(rowIndex, milestoneRatingIdx + 1).setValue(milestoneRating);
    participationSheet.getRange(rowIndex, actualScoreIdx + 1).setValue(actualScore);
    
    processed++;
    
    // 更新进度
    if (processed % updateInterval === 0 || processed === totalRows) {
      var percentComplete = Math.round((processed / totalRows) * 100);
      
      if (progressSheet) {
        progressSheet.getRange('B2').setValue(processed);
        progressSheet.getRange('D2').setValue(percentComplete + '%');
      }
      
      // 确保用户能看到进度表
      SpreadsheetApp.flush();
      
      // 使用toast作为辅助进度指示
      SpreadsheetApp.getActive().toast('已处理: ' + processed + ' / ' + totalRows + ' (' + percentComplete + '%)', '进度', 3);
    }
  }
  
  // 显示结果
  var resultMessage = '已重新计算 ' + processed + ' 条记录的分数\n' + 
                      (skipped > 0 ? '跳过了 ' + skipped + ' 条无效记录\n' : '') +
                      '使用新的从0到最高里程碑的评分方式计算完成！';
  
  // 在进度表中显示完成消息
  if (progressSheet) {
    progressSheet.getRange('A7:E7').merge();
    progressSheet.getRange('A7').setValue('计算完成! ' + resultMessage);
    progressSheet.getRange('A7').setFontWeight('bold');
    progressSheet.getRange('A7').setHorizontalAlignment('center');
    progressSheet.getRange('A7').setBackgroundRGB(220, 255, 220);
  }
  
  Browser.msgBox('完成', resultMessage, Browser.Buttons.OK);
  
  // 5秒后删除进度表
  Utilities.sleep(5000);
  if (progressSheet) {
    try {
      ss.deleteSheet(progressSheet);
    } catch (e) {
      // 忽略删除失败的错误
    }
  }
}
