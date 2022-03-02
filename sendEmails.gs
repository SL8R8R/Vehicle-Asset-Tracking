function sendEmail() {
  // Fetch the value of vehicle status
  var vehicleStatusRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Vehicles").getRange("B3:B"); 
  var vehicleStatus = vehicleStatusRange.getValues();

  var vehicleReasonsRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Vehicles").getRange("C3:G"); 
  var vehicleReasons = vehicleReasonsRange.getValues();

  var vehicleNumberRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Vehicles").getRange("A3:A"); 
  var vehicleNumber = vehicleNumberRange.getValues();

  // Check vehicle status
  if (vehicleStatus = "Maintenance Status"){
    // Fetch the email address
    var emailRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails").getRange("A1");
    var emailAddress = emailRange.getValue();
  
    // Send Alert Email.
    var message = 'The following vehicle needs service: Unit ' + vehicleNumber + "\n" + "Reasons for service: " + vehicleReasons; 
    var subject = 'Vehicle Service Request - Nordic Security Services';
    MailApp.sendEmail(emailAddress, subject, message);
    }
}
