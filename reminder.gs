function locale(x) { return Utilities.formatDate(new Date(x), 'PST', 'EEE, MMM d, yyyy') }

function email(recipent, content, date) {
  return GmailApp.sendEmail(recipent.replace(/;/g, ','), 'Reminder to return keys by '+locale(date), content, {
    name: 'Automatic Emailer',
    cc: ''
  })
}


function reminder() {
  var sheet = 'Log'
  var col1 = 'A'
  var colEnd = 'E'
  var startRow = 2
  var spreadsheet = SpreadsheetApp.openById('1cpZaQq1EseAwz6MK66-tjAO6LXgOVtMgGW0bTajUaE4')
  var len = spreadsheet.getRange(sheet+'!B1:B').getValues().filter(String).length
  var range = spreadsheet.getRange(sheet+'!'+col1+startRow+':'+colEnd+len)
  var data = range.getValues()
  var content = '' // email content
  
  for (var i in data) {
    var returnCol = data[i][2]
    var emailCol = data[i][3]
    
    // if return date and valid email exist
    if (returnCol && /@/.test(emailCol)) {
      var row = data[i]
      var nameCol = row[1]  
      var reminderCol = row[4]
      
      var cellDate = (new Date(returnCol)).valueOf() 
      var today = (new Date()).valueOf() 
      var daysInMilliSeconds = 86400000
      var isWithinOneWeek = cellDate<(today+(7*daysInMilliSeconds))
      var isWithinTwoWeeks = cellDate<(today+(14*daysInMilliSeconds))
      
      // get reminder value
      var currentRowNum = (Number(i)+startRow).toFixed(0)
      var reminderRange = spreadsheet.getRange("E"+currentRowNum)
      var currentVal = +(reminderRange.getValues()) // convert to Number
        
      if (reminderCol == 0 && isWithinTwoWeeks) {
        reminderRange.setValue(currentVal+1)
        content = 'reminder (within 2 weeks) - closed holidays'
        email(emailCol, content, returnCol)
      } else if (reminderCol == 1 && isWithinOneWeek) {
        reminderRange.setValue(currentVal+1)
        content = 'reminder (within 1 week) - closed on holidays'
        email(emailCol, content, returnCol)
      }
    }
  }

}
