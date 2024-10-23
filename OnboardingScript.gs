function sendOnboardingEmails(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName('FormResponses');
  var taskSheet = ss.getSheetByName('TaskList');

  if (!formSheet || !taskSheet) {
    Logger.log('Required sheets not found');
    return;
  }

  var lastRow = formSheet.getLastRow();
  if (lastRow < 1) {
    Logger.log('No data in "FormResponses" sheet');
    return;
  }

  var data = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
  if (!data) {
    Logger.log('No data retrieved from "FormResponses" sheet');
    return;
  }

  var name = data[1];
  var email = data[2];
  var assignedLab = data[4];
  var assignedProject = data[5];

  var taskSheet;
  var supervisorEmail;
  // Choose task sheet and supervisor based on lab assignment
  if (assignedLab === 'Animal lab') {
    taskSheet = animalLabSheet;
    supervisorEmail = 'sample@email.com';  // Replace with actual email
  } else if (assignedLab === 'Human lab') {
    taskSheet = humanLabSheet;
    supervisorEmail = 'sample@gmail.com';  // Replace with actual email
  } else {
    Logger.log('Invalid lab assignment');
    return;
  }

  // Task categories for new personnel
  var taskCategoriesForPersonnel = assignedLab == 'Animal Lab' ? {
    'Citi Requirements': [],
    'BioRaft Requirements': [],
    'Animal Handling': ["Receive training from MICV staff on animal handling (coordinate with supervisor and Lab Manager)"],
    'Card Access': [],
    'Work Alone': [],
    'Shared Drive Access': [],
    'Manual': []
  } : {
    'Citi Requirements': [],
    'BioRaft Requirements' : [],
    'Card Access' : [],
    'Work Alone' : [],
    'Shared Drive Access' : [], 
    'Manual' : []
  };

  // Task categories for hiring manager
  var taskCategoriesForManager = assignedLab == 'Animal Lab' ? {
    'Manager Task - Shared Drive Access': [],
    'Manager Task - IACUC': [],
    'Manager Task - BioRaft': ["Add new member and assign trainings"],
    'Manager Task - Introductions': []
  } : {
    'Manager Task - Shared Drive Access' : [],
    'Manager Task - Admin' : [],
    'Manager Task - BioRaft' : ["Add new member and assign trainings"]
  };

  // Retrieve all tasks from TaskList
  var tasks = taskSheet.getDataRange().getValues();
  var headers = tasks[0]; // Assuming first row contains headers

  Logger.log('Headers: ' + headers.join(', ')); // Log headers for debugging

  // Categorize tasks for both personnel and manager across all rows
  for (var i = 1; i < tasks.length; i++) { // Start at 1 to skip header row

    // Get tasks for new personnel (Columns B-G)
    for (var j = 1; j <= 7; j++) { // Columns B-G are indices 1-7
      var header = headers[j];
      if (header in taskCategoriesForPersonnel) {
        if (tasks[i][j]) {
          taskCategoriesForPersonnel[header].push(tasks[i][j]);
        }
      }
    }

    // Get tasks for hiring manager (Columns K-N)
    for (var k = 10; k <= 13; k++) { // Columns K-N are indices 10-13
      var managerHeader = headers[k];
      if (managerHeader in taskCategoriesForManager) {
        if (tasks[i][k]) {
          taskCategoriesForManager[managerHeader].push(tasks[i][k]);
        }
      }
    }
  }

  Logger.log('Personnel Tasks: ' + JSON.stringify(taskCategoriesForPersonnel)); // Log personnel tasks for debugging
  Logger.log('Manager Tasks: ' + JSON.stringify(taskCategoriesForManager)); // Log manager tasks for debugging

  // Prepare the email content for personnel
  var emailSubject = "Your Onboarding Tasks for blank";
  var emailBody = "Hi " + name + ",\n\nHere are the onboarding tasks for the project you've been assigned: " + assignedProject + ":\n\n";

  // Add tasks to the email body categorized by headers
  for (var header in taskCategoriesForPersonnel) {
    if (taskCategoriesForPersonnel[header].length > 0) {
      emailBody += header + ":\n";
      for (var k = 0; k < taskCategoriesForPersonnel[header].length; k++) {
        emailBody += "- " + taskCategoriesForPersonnel[header][k] + "\n";
      }
      emailBody += "\n"; // Add extra newline for separation
    }
  }

  emailBody += "Welcome! \n Onboarding Robot";

  // Send email to the new personnel
  MailApp.sendEmail(email, emailSubject, emailBody);

  // Notify Hiring Manager
  var managerEmail = 'sample@email.com'; // Replace with actual email or fetch dynamically

  var managerEmailSubject = "New Personnel Added";
  var managerEmailBody = "Hi name,\n\nThe new personnel " + name + " has been onboarded.\n\nPlease ensure the following tasks are completed:\n";

  // Add tasks to the email body categorized by columns K to N
  for (var managerHeader in taskCategoriesForManager) {
    if (taskCategoriesForManager[managerHeader].length > 0) {
      managerEmailBody += managerHeader + ":\n";
      for (var l = 0; l < taskCategoriesForManager[managerHeader].length; l++) {
        managerEmailBody += "- " + taskCategoriesForManager[managerHeader][l] + "\n";
      }
      managerEmailBody += "\n"; // Add extra newline for separation
    }
  }

  managerEmailBody += "\nBest regards,\nNML Onboarding Robot";

  // Send email to Hiring Manager
  MailApp.sendEmail(managerEmail, managerEmailSubject, managerEmailBody);
  // Notify Supervisor
  var supervisorEmail = 'sample@email.com'; // Replace with actual supervisor's email address

  var supervisorEmailSubject = managerEmailSubject;
  var supervisorEmailBody = managerEmailBody.replace('Hi name', 'Hi name'); // Replace manager's name with supervisor's

  // Send email to Supervisor
  MailApp.sendEmail(supervisorEmail, supervisorEmailSubject, supervisorEmailBody);
}

function onFormSubmit(e) {
  sendOnboardingEmails(e);
}

function deleteTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function setUpTrigger() {
  deleteTriggers(); // Clear existing triggers first
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}
