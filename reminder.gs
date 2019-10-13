function locale(x) { return Utilities.formatDate(new Date(x), 'PST', 'EEE MMM d, yyyy') }
function days(x) {
  var daysInMilliSeconds = 86400000
  return x * daysInMilliSeconds
}
function isToday(x) { return (new Date(x).toDateString()==(new Date()).toDateString()) }
function email(recipent, content, date, copy, due) {
  
  return GmailApp.sendEmail(recipent.replace(/;/g, ','), 'Reminder: '+(due?due+' key'+(due>1?'s':'')+' due -- ':'')+'return keys by '+(copy?'today -- ':'')+locale(date), content, {
    name: 'BEES Admin',
    cc: copy
  })
}
function content(name, date, keys) {
  var firstName = name.replace(/.+,\s/g, '')
  return 'Hi '+firstName+', your key return schedule is on '+locale(date)+'.'
  +'\n\nPlease plan to return keys @ 2460A Geology or reply to update your schedule.'
  +'\n\nOffice hours: 8:00 am - 12:00 pm and 1:00 pm - 4:00 pm (closed on Academic Holidays)'
  +(keys ? '\n\n----------KEYS-DUE------\n*'+keys.join('\n*'): '')
  +'\n\n-UCR BEES Administrative Unit (Automatic Emailer)'
}
function getSheet() {
  var x = {}
  x.sheet = 'Log'
  x.col1 = 'A'
  x.reminderColLetter = 'G'
  x.colEnd = 'H'
  x.startRow = 2
  x.spreadsheet = SpreadsheetApp.openById('1cpZaQq1EseAwz6MK66-tjAO6LXgOVtMgGW0bTajUaE4')
  x.len = x.spreadsheet.getRange(x.sheet+'!C1:C').getValues().filter(String).length
  x.range = x.spreadsheet.getRange(x.sheet+'!'+x.col1+x.startRow+':'+x.colEnd+x.len)
  x.data = x.range.getValues()
  
  return x
}

function weeklyReminder() {
  var sheet = getSheet()
  var startRow = sheet.startRow
  var reminderColLetter = sheet.reminderColLetter
  var spreadsheet = sheet.spreadsheet
  var data = sheet.data
  
  for (var i in data) {
    var scheduleCol = data[i][3]
    var emailCol = data[i][5]
    
    // if return date and valid email exist
    if (scheduleCol && /@/.test(emailCol)) {
      var row = data[i]
      var nameCol = row[2]
      var reminderCol = row[6]
      
      var cellDate = (new Date(scheduleCol)).valueOf() 
      var today = (new Date()).valueOf() 
      var isWithinOneWeek = cellDate<(today+days(7))
      var isWithinOneMonth = cellDate<(today+days(30))
      
      // get reminder value
      var currentRowNum = (Number(i)+startRow).toFixed(0)
      var reminderRange = spreadsheet.getRange(reminderColLetter+currentRowNum)
      var currentVal = +(reminderRange.getValues()) // convert to Number
        
      if (reminderCol == 0 && isWithinOneMonth) {
        reminderRange.setValue(currentVal+1)
        email(emailCol, content(nameCol, scheduleCol), scheduleCol)
      } else if (!isToday(scheduleCol) && reminderCol == 1 && isWithinOneWeek) {
        reminderRange.setValue(currentVal+1)
        email(emailCol, content(nameCol, scheduleCol), scheduleCol)
      }
    }
  }

}


function dailyReminder() {
  var sheet = getSheet()
  var startRow = sheet.startRow
  var reminderColLetter = sheet.reminderColLetter
  var spreadsheet = sheet.spreadsheet
  var data = sheet.data
  var copy = '' // admin email
  
  for (var i in data) {
    var scheduleCol = data[i][3]
    var emailCol = data[i][5]
    
    // if return date and valid email exist
    if (scheduleCol && /@/.test(emailCol) && isToday(scheduleCol)) {
      var row = data[i]
      var nameCol = row[2]  
      var reminderCol = row[6]
      
      // get reminder value
      var currentRowNum = (Number(i)+startRow).toFixed(0)
      var reminderRange = spreadsheet.getRange(reminderColLetter+currentRowNum)
      
      // compile list of unreturned keys
      var keys = []
      for (var x in data) {
        var thisRow = data[x]
        var thisKeyCol = thisRow[0]
        var thisAccessCol = thisRow[1]
        var thisNameCol = thisRow[2]
        var returnedCol = thisRow[4]
        if (new RegExp(nameCol).test(thisNameCol) && !returnedCol) {
          keys.push(thisKeyCol+' - '+thisAccessCol)
        }
      }
      
      reminderRange.setValue(-1)
      email(emailCol, content(nameCol, scheduleCol, keys), scheduleCol, copy, keys.length)
      
    }
  }

}
