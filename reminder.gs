function locale(x) { return Utilities.formatDate(new Date(x), 'PST', 'EEE, MMM d, yyyy') }

function email(recipent, content, date, copy) {
  return GmailApp.sendEmail(recipent.replace(/;/g, ','), 'Reminder to return keys by '+(copy&&'today: ')+locale(date), content, {
    name: 'Automatic Emailer',
    cc: copy
  })
}

function getSheet() {
  var x = {}
  x.sheet = 'Log'
  x.col1 = 'A'
  x.colEnd = 'E'
  x.startRow = 2
  x.spreadsheet = SpreadsheetApp.openById('1cpZaQq1EseAwz6MK66-tjAO6LXgOVtMgGW0bTajUaE4')
  x.len = x.spreadsheet.getRange(x.sheet+'!B1:B').getValues().filter(String).length
  x.range = x.spreadsheet.getRange(x.sheet+'!'+x.col1+x.startRow+':'+x.colEnd+x.len)
  x.data = x.range.getValues()
  
  return x
}

function weeklyReminder() {
  var sheet = getSheet()
  var startRow = sheet.startRow
  var spreadsheet = sheet.spreadsheet
  var data = sheet.data
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
        email(emailCol, content, returnCol, '')
      } else if (reminderCol == 1 && isWithinOneWeek) {
        reminderRange.setValue(currentVal+1)
        content = 'reminder (within 1 week) - closed on holidays'
        email(emailCol, content, returnCol, '')
      }
    }
  }

}


function dailyReminder() {
  var sheet = getSheet()
  var startRow = sheet.startRow
  var spreadsheet = sheet.spreadsheet
  var data = sheet.data
  var content = '' // email content
  var copy = '' // admin email
  
  for (var i in data) {
    var returnCol = data[i][2]
    var emailCol = data[i][3]
    
    // if return date and valid email exist
    if (returnCol && /@/.test(emailCol)) {
      var row = data[i]
      var nameCol = row[1]  
      var reminderCol = row[4]
      
      var isToday = (new Date(returnCol).toDateString()==(new Date()).toDateString())
      
      // get reminder value
      var currentRowNum = (Number(i)+startRow).toFixed(0)
      var reminderRange = spreadsheet.getRange("E"+currentRowNum)
        
      if (isToday) {
        reminderRange.setValue(-1)
        content = 'reminder (today) - closed holidays'
        email(emailCol, content, returnCol, copy)
      }
    }
  }

}
