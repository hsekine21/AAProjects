// Yay it's a new attendance sheet!!!

// Thingies that are needed 

var tutoringDragonsEmail = 'tutoringdragons@ssis.edu.vn'

// SET UP 
var sheet = SpreadsheetApp.getActiveSheet();

// FUNCTIONS 
function writeSheetCell(sheetname, cellPlace, replaceText){
  var dataMake = sheetname + '!' + cellPlace
  sheet.getRange(dataMake).setValue(replaceText)
}

/*
function makeWeekDateCalendars(){
  // function makes the calendar dates for each day of the week 
  var calendarDates = []
  var weekday = ['B1', 'C1', 'D1', 'E1', 'F1']
  
  for each (var col in weekday){
    // gets the lists from the spreadsheet [[], [], [], ... , []]
    dataMake = "Calendar!" + col
    var weekDates = sheet.getRange(dataMake).getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues()
   
    var weekCorrection = []
    // make the weekdates non lists 
    for each (var thing in weekDates){
      weekCorrection.push(thing[0])
    }
    
    calendarDates.push(weekCorrection) 
  }
  return calendarDates 
}

function addDatesToAttendance(dayNum, sheetname, calendarDates){ // day num is M = 0, T = 1, W = 2 ... 
  // adds the dates to the attendance sheet & makes them black if they are 'NONE'
  var rangeName = sheetname + "!B1:Q1"
  
  var colHead = 'BCDEFGHIJKLMNOPQ' 
  var datesData = calendarDates[dayNum]
  
  var i;
  for (i = 0; i < 16; i++) {
    var position = colHead[i]
    var weekDate = datesData[i]
    
    if (weekDate == 'NONE'){
      // makes the whole colum black if the day doesn't have tutoring 
      sheet.getRange(sheetname + '!' + position + '1' + ':' + position + '100').setBackground('gray')
    }
    else{
      writeSheetCell(sheetname, position + '1', weekDate)
    }
  }
  
  sheet.getRange(rangeName).setFontStyle('bold')
}
*/

function tutorAttendanceNames(){
  // literally just put names in the 'TutorAttendance' sheet from the 'TutorNames' sheet
  var tutorNames = sheet.getRange('MASTER!B2:B100').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues(); // gets all the tutor names 
  
  for (i = 1; i < (tutorNames.length); i++) {
    var cellPlace = 'A' + (i + 1).toString()
    writeSheetCell('TutorAttendance', cellPlace, tutorNames[i])
  }
}

function makeTutorStudentPairs(){
  // this function makes the WeekTutorStudentPairs
  
  // data is all [[ignore 0], [], [], ...]
  var tutorNames = sheet.getRange('MASTER!B1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  var monday = sheet.getRange('MASTER!C1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  var tuesday = sheet.getRange('MASTER!D2').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  var wednesday = sheet.getRange('MASTER!E2').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  var thursday = sheet.getRange('MASTER!F2').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  var friday = sheet.getRange('MASTER!G2').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues();
  
  var schedule = [monday, tuesday, wednesday, thursday, friday]
  
  var masterNames = [[], []] // [[[tutor monday], [tutor tuesday], .....], [[students monday], [students tuesday], ....]
  
  // looping through the student names for each week (makes master names)
  for each (var weekday in schedule){ // looping through each week 
    var tutor = []
    var student = []
    for each (i = 1; i < (weekday.length); i++){ // looping through each name (y, n, name)
      var name = weekday[i]
      if (name != 'Y' && name != 'N'){ // if there is a name 
        tutor.push(tutorNames[i][0])
        student.push(name[0])
        }
      }
    masterNames[0].push(tutor)
    masterNames[1].push(student)
    }
  
  // FIRST DELETE everything on the sheet 
  var colDelete = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
  for each (var col in colDelete){
    for (i = 2; i < 100; i++){
      writeSheetCell('WeekTutorStudentPairs', col + i.toString(), ' ')
    }
  }
  
  // now add them to the sheet
  var colWeekdays = [['A', 'B'], ['C', 'D'], ['E', 'F'], ['G', 'H'], ['I', 'J']]
  for (i = 0; i < 5; i++){
    var studentNamesList = masterNames[1][i] // takes the name for the day of the week 
    var tutorNamesList = masterNames[0][i]
    var colTutor = colWeekdays[i][0] // takes the colums thingies for the day of the week 
    var colStudent = colWeekdays[i][1]
    
    // now add them 
    var row = 2 
    for (k = 0; k < (studentNamesList.length); k++){
      // replace for the tutor 
      writeSheetCell('WeekTutorStudentPairs', colTutor + row.toString(), tutorNamesList[k])
      // replace for the sudent 
      writeSheetCell('WeekTutorStudentPairs', colStudent + row.toString(), studentNamesList[k])
      
      row = row + 1 
    }
  }
  }

function attendanceEvaluation(){
  // this function stores the Present, Unexcused Absence, Excused Absence from Monday - Friday sheets for each person into TutorAttendance
  // could be more efficient by using date and only updating that 
  
  var weekSheetNames = ['MondayTutor', 'TuesdayTutor', 'WednesdayTutor', 'ThursdayTutor', 'FridayTutor'] 
  
  var masterAttendance = []
  
  for each (var sheetName in weekSheetNames){ // WILL NOT WORK UNTIL ALL THE SHEETS ARE MADE 
    
  // for each week sheet names 
  
  // get all the names from the top of the sheet
  var dataMake = sheetName + '!B1'
  var tutorNames = sheet.getRange(dataMake).getDataRegion(SpreadsheetApp.Dimension.COLUMNS).getDisplayValues() // [[timestamp, name, name ....]]
  tutorNames = tutorNames[0] // removes outer brackets 
  
  // get the Present, Excused Absence, Unexcused Absence lists as a parllell list 
  var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  var attendanceLinkNames = []
  for (i = 0; i < tutorNames.length; i++){ 
    dataMake = sheetName + '!' + alphabet[i] + '1'
    var personAttendance = sheet.getRange(dataMake).getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues() // [abigail, present, unexcused...]
    attendanceLinkNames.push(personAttendance)
  }
  
  // now make a new list with present, unexcused, excused counts 
  // tutorNames = [abigail, billy, cassandra]
  // new one = [['abigail', [5, 6, 7]], ['billy', [6, 1, 1]], ['cassandra', [1, 1, 1]]...]
  // [Present, Unexcused Absense, Excused Absense] each one up top 
  var namesAttendanceCounts = []
  for (s = 1; s < attendanceLinkNames.length; s++){
    var markAttendance = attendanceLinkNames[s]
    var name = markAttendance[0]
    
    var attendanceCounts = [0, 0, 0] // [Present, Unexcused, Excused]
    for (m = 1; m < markAttendance.length; m++){
      var mark = markAttendance[m]
      if (mark == 'Present'){
        attendanceCounts[0] = attendanceCounts[0] + 1
      }
      else if (mark == 'Unexcused Absence'){
        attendanceCounts[1] = attendanceCounts[1] + 1
      }
      else if (mark == 'Excused Absence'){
        attendanceCounts[2] = attendanceCounts[2] + 1
      }
    }
    masterAttendance.push([name[0], attendanceCounts])    
  }}
  
  //now put them into the spreadsheet 
  var tutorAttendanceList = sheet.getRange('TutorAttendance!A1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues() // get the names of all the kids [[Names], [Name1], [Name2], [Name3], [Name4], ]
  var masterTutorNames = []
  for each (var hsTutorName in tutorAttendanceList){
    masterTutorNames.push(hsTutorName[0])
  }
  
  for each (var person in masterAttendance){
    var personName = person[0]
    var plusPres = person[1][0]
    var plusUnex = person[1][1]
    var plusEx = person[1][2]
    
    // finds the index of the name needed using the masterTutorNames
    var personRow = masterTutorNames.indexOf(personName) + 1
    
    var presCellAdress = 'TutorAttendance!B' + personRow.toString()
    var unexCellAdress = 'TutorAttendance!C' + personRow.toString()
    var exCellAdress = 'TutorAttendance!D' + personRow.toString()
    
    // gets the current attendance values from the sheet 
    var currPres = sheet.getRange(presCellAdress).getDisplayValue()
    var currUnex = sheet.getRange(unexCellAdress).getDisplayValue()
    var currEx = sheet.getRange(exCellAdress).getDisplayValue()
    
    // now adds them to the sheet 
    sheet.getRange(presCellAdress).setValue(Number(currPres) + Number(plusPres))
    sheet.getRange(unexCellAdress).setValue(Number(currUnex) + Number(plusUnex))
    sheet.getRange(exCellAdress).setValue(Number(currEx) + Number(plusEx))  
  } 
}

/*
function tooManyUnexcusedAbsences(tutoringDragonsEmail){
  // sends emails to kids when they have more than 3 absenses  
  var unexcusedCounts = sheet.getRange('TutorAttendance!C1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues() // this gets the attendance counts from TutorAttendance = [[Unexcused Absenses], [num]. [num]...]
  var tutorEmails = sheet.getRange('MASTER!A1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues() // this gets the emails of the tutors from MASTER = [[Email Adress], [yay@ssis.edu.vn], [yay1@ssis.edu.vn]...]
  // testing 
  Logger.log(unexcusedCounts)
  Logger.log(tutorEmails)
  // now it's time to figure out who the late kids are 

}
*/ 


// THIS WHOLE SET IS GOING TO SEND EMAILS, you little toaster 
function tutoringToday(){
  // this function tells you if there is tutoring today via tech writing on the 'info' sheet 
  // returns today's date if there is tutoring today, if not, returns false 
  var todayDate = sheet.getRange('info!B2').getDisplayValue()
  if (todayDate == 'NONE'){
    return false 
  }
  else{
    return todayDate
  }
}

function setAttendanceOfficer(todayDate){
  // this function figures out the attendance officer for that day with the date passed in, assuming this function is only used when there is a tutoring dragons date
  
  var attendanceOfficers = sheet.getRange('info!F1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getDisplayValues() // this gets the emails for all the attendance officiers [['Emails '], [yay@ssis.edu.vn], ...]
  var dayOfWeek = todayDate.getday()
  
  Logger.log(dayOfWeek)
  
}

/*
function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
  var range = e.range;
  range.setNote('Last modified: ' + new Date());
}
*/

// MAIN ----------------------------------------------------

function mainOnce(){  
  // these are functions that only need to be run once :D 
  
  /*
  // Adds calendar dates to all of the weekly attendance sheets 
  calendarDates = makeWeekDateCalendars() // makes the calendar in the form [[MONDAY DATES...], [TUESDAY DATES...]]
  var weekSheetNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
  for (i = 0; i < 5; i++){
      addDatesToAttendance(i, weekSheetNames[i], calendarDates)
  }
  */
  
  //Logger.clear()
  //Logger.log(calendarDates)
}

function repeatDaily(){
  // these are functions that need to be run every day :) 
  attendanceEvaluation()
}

function repeatForNewTutorStudent(){
  // these are functions that need to be run every time a new tutor is added (basically every day)
  tutorAttendanceNames()
  makeTutorStudentPairs()
}

function repeatForScheduleChange(){
  makeTutorStudentPairs()
}


function test(){
  setAttendanceOfficer('January 30 2020')
}


















