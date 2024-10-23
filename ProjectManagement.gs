function submitProject() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName('Ongoing Projects');
  
  // Extract data from projectData object
  var projectName = projectData.projectName;
  var teamMembers = projectData.teamMembers;
  var description = projectData.description;
  var trainings = projectData.trainings;
  var protocolNumber = projectData.protocolNumber;
  var date = projectData.date;
  
  // Log all values to ensure they're being read correctly
  Logger.log('Project Name: ' + projectName);
  Logger.log('Team Members: ' + teamMembers);
  Logger.log('Description: ' + description);
  Logger.log('Trainings: ' + trainings);
  Logger.log('Protocol Number: ' + protocolNumber);
  Logger.log('Date (Before Processing): ' + date);
  
  // Convert date string to Date object if necessary
  if (date) {
    var parsedDate = new Date(date);
    if (!isNaN(parsedDate.getTime())) {
      date = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      date = ""; // Set to empty string if date cannot be parsed
      Logger.log('Date could not be parsed: ' + date);
    }
  }

  Logger.log('Date (After Processing): ' + date); // Log the final date value

  // Append the data to the Ongoing Projects sheet
  dataSheet.appendRow([
    projectName,
    teamMembers,
    description,
    trainings,
    protocolNumber,
    date
  ]);
  
  // Email notification
  var emailBody = 'New project added:\n\n' +
                  'Project Name: ' + projectName + '\n' +
                  'Team Members: ' + teamMembers + '\n' +
                  'Description: ' + description + '\n' +
                  'Relevant Trainings: ' + trainings + '\n' +
                  'Protocol Number: ' + protocolNumber + '\n' +
                  'Date: ' + date + '\n\n' +
                  'Best regards,\n';
  
  MailApp.sendEmail({
    to: 'sample@email.com', 
    subject: 'New Project Added: ' + projectName,
    body: emailBody
  });
}

// entry form in html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
</head>
<body>
  <h1>Project Entry Form</h1>
  <form id="projectForm">
    <label for="cellB3">Project Name:</label><br>
    <input type="text" id="cellB3" name="projectName" required><br><br>
    <label for="cellD3">Team Members:</label><br>
    <input type="text" id="cellD3" name="teamMembers" required><br><br>
    <label for="cellB6">Description:</label><br>
    <textarea id="cellB6" name="description" required></textarea><br><br>
    <label for="cellD6">Relevant Trainings:</label><br>
    <input type="text" id="cellD6" name="trainings"><br><br>
    <label for="cellB9">Protocol Number:</label><br>
    <input type="text" id="cellB9" name="protocolNumber" required><br><br>
    <label for="cellD9">Date:</label><br>
    <input type="date" id="cellD9" name="date" required><br><br>
    <input type="button" value="Submit" onclick="submitProject()">
  </form>
  <script>
    function submitProject() {
      var projectData = {
        projectName: document.getElementById('cellB3').value,
        teamMembers: document.getElementById('cellD3').value,
        description: document.getElementById('cellB6').value,
        trainings: document.getElementById('cellD6').value,
        protocolNumber: document.getElementById('cellB9').value,
        date: document.getElementById('cellD9').value
      };
    }
  </script>
</body>
</html>




