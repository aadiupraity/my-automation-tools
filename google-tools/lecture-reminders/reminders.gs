function createRecurringLectureReminders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; 
  var numRows = sheet.getLastRow() - 1; 
  
  var dataRange = sheet.getRange(startRow, 1, numRows, 5);
  var data = dataRange.getValues();
  
  var cal = CalendarApp.getDefaultCalendar();

  
  var dayMap = {
    'monday': 1,
    'tuesday': 2,
    'wednesday': 3,
    'thursday': 4,
    'friday': 5,
    'saturday': 6,
    'sunday': 0
  };

  // End date set to November 1st, 2026, since that is the approx. date when my semester ends, edit accordingly
  var recurrenceEndDate = new Date(2026, 10, 1); 

  for (var i = 0; i < data.length; i++) {
    var row = data[i]; 
    var title = row[0]; 
    var dayName = String(row[1]).toLowerCase().trim(); 
    var startTime = row[2]; 
    var endTime = row[3]; 
    var location = row[4];
    
    if (title && dayName && startTime && endTime) {
      var targetDayNum = dayMap[dayName];
      if (targetDayNum === undefined) continue;

      var today = new Date();
      var firstOccurrenceDate = new Date(today.getTime());
      while (firstOccurrenceDate.getDay() !== targetDayNum) {
        firstOccurrenceDate.setDate(firstOccurrenceDate.getDate() + 1);
      }
      
      // Construct start and end times for the first instance of an event
      var start = new Date(firstOccurrenceDate.getFullYear(), firstOccurrenceDate.getMonth(), firstOccurrenceDate.getDate(), startTime.getHours(), startTime.getMinutes(), 0);
      var end = new Date(firstOccurrenceDate.getFullYear(), firstOccurrenceDate.getMonth(), firstOccurrenceDate.getDate(), endTime.getHours(), endTime.getMinutes(), 0);
      
      // Build the Weekly Recurrence
      var recurrencePattern;
      if (targetDayNum === 1) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.MONDAY).until(recurrenceEndDate);
      else if (targetDayNum === 2) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.TUESDAY).until(recurrenceEndDate);
      else if (targetDayNum === 3) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.WEDNESDAY).until(recurrenceEndDate);
      else if (targetDayNum === 4) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.THURSDAY).until(recurrenceEndDate);
      else if (targetDayNum === 5) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.FRIDAY).until(recurrenceEndDate);
      else if (targetDayNum === 6) recurrencePattern = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.SATURDAY).until(recurrenceEndDate);

      var series = cal.createEventSeries(title, start, end, recurrencePattern, {location: location});
      
      // Add a reminder for 15 minutes before
      series.addPopupReminder(15);
      
      Logger.log('Created recurring tasks for: ' + title);
    }
  }
  SpreadsheetApp.getUi().alert('All recurring lecture reminders have been added to your Google Calendar until the end of October!'); 
  // Again, you may edit this according to your end of semester, or not. 
}
