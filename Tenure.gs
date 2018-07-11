/**
 * Returns the date difference in Years, Months, and Days. termDate is optional, if not used the current date will be used.
 *
 * @param {startDate} input first date.
 * @param {termDate} (Optional) input end date.
 * @return The difference in years, months, and days.
 * @customfunction
 */

function TENURE(startDate, termDate) {
  
  var sheet = SpreadsheetApp.getActive();
  var endDate;
  
  if (termDate == "" | termDate == null) {endDate = new Date();}
    else {endDate = termDate;}
  
// valueOf gives milliseconds since midnight on 1st of Jan 1970
// Convert milliseconds to seconds then minutes then hours then days
// Averaging years to 356.25 to include leap years and months to 30.5
  var dateDiffInDays = (endDate.valueOf()-startDate.valueOf())/(1000*60*60*24);
  var years = Math.floor(dateDiffInDays/365.25);
  var months = Math.floor((dateDiffInDays%365.25)/30.5);
  var days = Math.floor(((dateDiffInDays%365.25)%30.5));
  
  Logger.log("Years: " + years);
  Logger.log("Months: " + months);
  Logger.log("Days : " + days);
  
  var tenure = years + " Years, " + months + " Months, " + days + " Days";

  return(tenure);
  
}
