function SendResultsEmails() {

  // SHEET SETUP
  const Sheet = SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results')); // Setup your sheet here
  const EndCell = 'D' + 100; // Number of rows to process
  const DataRange = Sheet.getRange('A2:'+ EndCell); // Cell Reference

  // EMAIL CONSTANTS
  const Subject = 'UNIT1000 - Final Exam Results';
  const CCEmails = ['examiner@university.edu', 'lecturer@university.edu', 'tutor@university.edu'];
  const SentTime = new Date().toLocaleString().replace(',','').replace(/:.. /,' '); // Format time
  
  // LIST OF NAMES
  const PassList = [];
  const FailList = [];

  // SEND EMAILS TO STUDENTS
  const Data = DataRange.getValues();
  // Go through Sheet
  for (var i = 0; i < Data.length; i++) {
      var Row = Data[i];
      // Student Info
      var EmailAddress = Row[0];  
      var StudentName = Row[1];
      var Score = Row[2];
      var EmailStatus = Row[3];

      // Check if Email Sent
      if (EmailStatus !== 'Email Sent') {
        // Approved Message
        if (Score >= 50){
          var Message = '<b>CONGRATULATIONS, YOU HAVE PASSED UNIT1000!</b>';
          var Recommendation = "We suggest that you take a look at the units you might want to take next semester to fulfill your major's requirements";
          PassList.push(StudentName + ' - ' + EmailAddress + ' (' + Score + ')')
          // Add to Pass List
        } else {
          var Message = "Unfortunately, you did not meet this unit's minimum requirements";
          var Recommendation = 'We suggest that you retake this unit next semester.';
          FailList.push(StudentName + ' - ' + EmailAddress + ' (' + Score + ')');
        }
        // Send Email
        MailApp.sendEmail({  to:        EmailAddress
                            ,cc:        CCEmails.join()
                            ,subject:   Subject
                            ,htmlBody:  'Dear ' + StudentName + ',<br><br>'+
                                         Message + '<br>' +
                                        'Your final score for UNIT1000 was ' + Score + '.<br>' +
                                         Recommendation + '<br>'
                                        'Please email us if you have any queries.<br><br>' + 
                                        'Thanks<br>' +
                                        'UNIT1000 Teaching Team'
                          });
        // Update Sheet
        Sheet.getRange('D' + (i + 2)).setValue('Email Sent');
        Sheet.getRange('E' + (i + 2)).setValue(SentTime);
        SpreadsheetApp.flush();
      }
  }
  
  // SEND COMPILED PASS/FAIL LIST
  // Join lists and order ascending by name
  const HTMLPassList = '<ol><li>' + PassList.sort().join('</li><li>') + '</li></ol>';
  const HTMLFailList = '<ol><li>' + FailList.sort().join('</li><li>') + '</li></ol>';
  // Send Email to HoD
  MailApp.sendEmail({  to:        'headofschool@university.edu'
                      ,cc:        CCEmails.join()
                      ,subject:   'UNIT1000 - Final Exam Results'
                      ,htmlBody:  'Hi Team <br><br>' + 
                                  'Below is a list of passes and failures this semester for UNIT1000: <br>' + 
                                  '<b>PASS<\b>' +
                                   HTMLPassList +
                                  '<b>FAIL</b>' + 
                                   HTMLFailList +
                                  'Emails sent at ' + SentTime + '.<br>' + 
                                  'Let me know if you have any questions. <br><br>' + 
                                  'Have a nice day!<br>' + 
                                  'Brice'
                    });
}
