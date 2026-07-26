// ============================================
// TAT CALCULATOR - Turnaround Time as per Office Hours
// ============================================
// 
// HOW TO USE:
// ==========
// Call: calculateTAT(submissionDateTime, tatHours, officeStartTime, officeEndTime, weekOffDays, holidaySheetName)
//
// Parameters:
// - submissionDateTime: Date object or string "2025-01-20 17:30" (when work was submitted)
// - tatHours: Number (TAT in hours, e.g. 2, 4, 8, 24)
// - officeStartTime: String "10:00" (office start time in HH:MM)
// - officeEndTime: String "18:30" (office end time in HH:MM)
// - weekOffDays: Array of day numbers [0] = Sunday, [6] = Saturday, e.g. [0] for Sunday off
// - holidaySheetName: String - name of sheet containing holidays (Col A = Date)
// - sheetUrl: String - Google Sheet URL containing holiday sheet
//
// RETURNS: Date object (the due date/time as per office hours)
//
// HOLIDAY SHEET FORMAT:
// Col A = Date (dd/mm/yyyy or any date format)
// Col B = Holiday Name (optional, for reference)
//
// EXAMPLE:
// var dueDate = calculateTAT("2025-01-20 17:30", 2, "10:00", "18:30", [0], "HOLIDAYS", sheetUrl);
// Result: 2025-01-21 11:00 (because only 1hr left today, remaining 1hr carries to next working day)
//
// ============================================

/**
 * MAIN TAT CALCULATOR FUNCTION
 */
function calculateTAT(submissionDateTime, tatHours, officeStartTime, officeEndTime, weekOffDays, holidaySheetName, sheetUrl) {
  
  // Parse inputs
  var submission = parseDateTime(submissionDateTime);
  var officeStart = parseTime(officeStartTime); // {hours, minutes}
  var officeEnd = parseTime(officeEndTime);     // {hours, minutes}
  
  // Get holidays list
  var holidays = getHolidaysList(holidaySheetName, sheetUrl);
  
  // Calculate office hours per day (in minutes)
  var officeMinutesPerDay = (officeEnd.hours * 60 + officeEnd.minutes) - (officeStart.hours * 60 + officeStart.minutes);
  
  // Convert TAT to minutes
  var tatMinutesRemaining = tatHours * 60;
  
  // Start calculating from submission time
  var currentDate = new Date(submission.getTime());
  
  // STEP 1: Check if submission is on a working day and within office hours
  if (!isWorkingDay(currentDate, weekOffDays, holidays)) {
    // If submitted on holiday/weekoff, move to next working day start
    currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
  } else {
    // Check if submission is within office hours
    var currentMinutes = currentDate.getHours() * 60 + currentDate.getMinutes();
    var startMinutes = officeStart.hours * 60 + officeStart.minutes;
    var endMinutes = officeEnd.hours * 60 + officeEnd.minutes;
    
    if (currentMinutes >= endMinutes) {
      // After office hours - move to next working day start
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    } else if (currentMinutes < startMinutes) {
      // Before office hours - set to office start
      currentDate.setHours(officeStart.hours, officeStart.minutes, 0, 0);
    }
    // If within office hours, keep current time
  }
  
  // STEP 2: Now consume TAT minutes across working days
  while (tatMinutesRemaining > 0) {
    
    // Ensure we're on a working day
    if (!isWorkingDay(currentDate, weekOffDays, holidays)) {
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    }
    
    // Calculate remaining office minutes today
    var currentMinutes = currentDate.getHours() * 60 + currentDate.getMinutes();
    var endMinutes = officeEnd.hours * 60 + officeEnd.minutes;
    var startMinutes = officeStart.hours * 60 + officeStart.minutes;
    
    // If before office start, move to office start
    if (currentMinutes < startMinutes) {
      currentDate.setHours(officeStart.hours, officeStart.minutes, 0, 0);
      currentMinutes = startMinutes;
    }
    
    var remainingTodayMinutes = endMinutes - currentMinutes;
    
    if (remainingTodayMinutes <= 0) {
      // No time left today, move to next working day
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
      continue;
    }
    
    if (tatMinutesRemaining <= remainingTodayMinutes) {
      // TAT finishes today!
      currentDate.setMinutes(currentDate.getMinutes() + tatMinutesRemaining);
      tatMinutesRemaining = 0;
    } else {
      // Consume today's remaining hours and move to next day
      tatMinutesRemaining -= remainingTodayMinutes;
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    }
  }
  
  return currentDate;
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Parse date/time string or Date object
 */
function parseDateTime(input) {
  if (input instanceof Date) return new Date(input.getTime());
  
  // Try parsing "YYYY-MM-DD HH:MM" or "DD/MM/YYYY HH:MM"
  if (typeof input === 'string') {
    // Try standard format first
    var d = new Date(input);
    if (!isNaN(d.getTime())) return d;
    
    // Try DD/MM/YYYY HH:MM
    var parts = input.split(' ');
    var dateParts = parts[0].split('/');
    var timeParts = parts[1] ? parts[1].split(':') : ['0', '0'];
    
    if (dateParts.length === 3) {
      d = new Date(
        parseInt(dateParts[2]), // year
        parseInt(dateParts[1]) - 1, // month (0-indexed)
        parseInt(dateParts[0]), // day
        parseInt(timeParts[0]), // hours
        parseInt(timeParts[1])  // minutes
      );
      if (!isNaN(d.getTime())) return d;
    }
  }
  
  throw new Error('Invalid date format: ' + input);
}

/**
 * Parse time string "HH:MM" to {hours, minutes}
 */
function parseTime(timeStr) {
  var parts = timeStr.split(':');
  return {
    hours: parseInt(parts[0]),
    minutes: parseInt(parts[1]) || 0
  };
}

/**
 * Check if a date is a working day (not weekend, not holiday)
 */
function isWorkingDay(date, weekOffDays, holidays) {
  // Check week off
  var dayOfWeek = date.getDay(); // 0=Sunday, 6=Saturday
  if (weekOffDays.indexOf(dayOfWeek) !== -1) return false;
  
  // Check holidays
  var dateStr = formatDateKey(date);
  if (holidays.indexOf(dateStr) !== -1) return false;
  
  return true;
}

/**
 * Get next working day at office start time
 */
function getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays) {
  var nextDay = new Date(currentDate.getTime());
  nextDay.setDate(nextDay.getDate() + 1);
  nextDay.setHours(officeStart.hours, officeStart.minutes, 0, 0);
  
  // Keep moving forward until we find a working day
  var safety = 0;
  while (!isWorkingDay(nextDay, weekOffDays, holidays) && safety < 365) {
    nextDay.setDate(nextDay.getDate() + 1);
    safety++;
  }
  
  return nextDay;
}

/**
 * Format date as "YYYY-MM-DD" for comparison
 */
function formatDateKey(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * Get holidays from Google Sheet
 * Holiday sheet: Col A = Date
 */
function getHolidaysList(holidaySheetName, sheetUrl) {
  var holidays = [];
  
  try {
    if (!sheetUrl || !holidaySheetName) return holidays;
    
    var ss = SpreadsheetApp.openByUrl(sheetUrl);
    var sheet = ss.getSheetByName(holidaySheetName);
    
    if (!sheet) return holidays;
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) { // Skip header row
      var cellValue = data[i][0];
      if (cellValue) {
        var d;
        if (cellValue instanceof Date) {
          d = cellValue;
        } else {
          d = new Date(cellValue);
        }
        if (!isNaN(d.getTime())) {
          holidays.push(formatDateKey(d));
        }
      }
    }
  } catch (e) {
    Logger.log('Holiday fetch error: ' + e.message);
  }
  
  return holidays;
}

// ============================================
// TEST FUNCTION - Run this to test!
// ============================================
function testTATCalculator() {
  var sheetUrl = ''; // PUT YOUR SHEET URL HERE FOR HOLIDAY TEST
  
  // Test scenarios WITHOUT holidays (pass empty array manually)
  Logger.log('======= TAT CALCULATOR TESTS =======');
  Logger.log('');
  
  // TEST 1: Submit at 5:30 PM, TAT = 2 hours, Office 10:00-18:30
  // Remaining today = 1 hour (5:30 to 6:30), Carry 1 hour to next day
  // Expected: Next day 11:00 AM
  var test1 = calculateTATLocal(
    new Date(2025, 0, 20, 17, 30), // Jan 20, 2025, 5:30 PM (Monday)
    2,          // 2 hours TAT
    '10:00',    // Office start
    '18:30',    // Office end
    [0],        // Sunday off
    []          // No holidays
  );
  Logger.log('TEST 1: Submit 5:30 PM, TAT 2 hrs');
  Logger.log('Expected: Jan 21, 2025 11:00 AM');
  Logger.log('Got:      ' + formatResult(test1));
  Logger.log('');
  
  // TEST 2: Submit at 6:30 PM (after hours), TAT = 3 hours
  // Should start from next day 10:00 AM, finish at 1:00 PM
  var test2 = calculateTATLocal(
    new Date(2025, 0, 20, 18, 30), // Jan 20, 2025, 6:30 PM
    3,
    '10:00',
    '18:30',
    [0],
    []
  );
  Logger.log('TEST 2: Submit 6:30 PM (after hours), TAT 3 hrs');
  Logger.log('Expected: Jan 21, 2025 01:00 PM');
  Logger.log('Got:      ' + formatResult(test2));
  Logger.log('');
  
  // TEST 3: Submit Friday 5:00 PM, TAT = 4 hours, Sunday off
  // Remaining Friday = 1.5 hrs (5:00 to 6:30), Carry 2.5 hrs
  // Saturday is working (only Sunday off), so Saturday 10:00 + 2.5 hrs = 12:30
  var test3 = calculateTATLocal(
    new Date(2025, 0, 24, 17, 0), // Jan 24, 2025, 5:00 PM (Friday)
    4,
    '10:00',
    '18:30',
    [0],       // Only Sunday off
    []
  );
  Logger.log('TEST 3: Submit Friday 5 PM, TAT 4 hrs, Sunday off');
  Logger.log('Expected: Jan 25, 2025 12:30 PM (Saturday)');
  Logger.log('Got:      ' + formatResult(test3));
  Logger.log('');
  
  // TEST 4: Submit Saturday 3 PM, TAT = 2 hours, Sat+Sun off
  // Saturday is off, so move to Monday 10:00 AM + 2 hrs = 12:00 PM
  var test4 = calculateTATLocal(
    new Date(2025, 0, 25, 15, 0), // Jan 25, 2025 (Saturday)
    2,
    '10:00',
    '18:30',
    [0, 6],    // Sunday + Saturday off
    []
  );
  Logger.log('TEST 4: Submit Saturday, TAT 2 hrs, Sat+Sun off');
  Logger.log('Expected: Jan 27, 2025 12:00 PM (Monday)');
  Logger.log('Got:      ' + formatResult(test4));
  Logger.log('');
  
  // TEST 5: Submit Monday 4 PM, TAT = 2 hrs, Tuesday is holiday
  // Remaining Monday = 2.5 hrs (4:00 to 6:30), TAT = 2 hrs < 2.5 hrs
  // Finishes today at 6:00 PM
  var test5 = calculateTATLocal(
    new Date(2025, 0, 20, 16, 0), // Jan 20, 2025 (Monday) 4:00 PM
    2,
    '10:00',
    '18:30',
    [0],
    ['2025-01-21'] // Tuesday is holiday
  );
  Logger.log('TEST 5: Submit Mon 4 PM, TAT 2 hrs, Tue holiday');
  Logger.log('Expected: Jan 20, 2025 06:00 PM (same day - enough time)');
  Logger.log('Got:      ' + formatResult(test5));
  Logger.log('');
  
  // TEST 6: Submit Monday 5:30 PM, TAT = 2 hrs, Tuesday is holiday
  // Remaining Monday = 1 hr (5:30 to 6:30), Carry 1 hr
  // Tuesday is holiday, skip to Wednesday 10:00 + 1 hr = 11:00
  var test6 = calculateTATLocal(
    new Date(2025, 0, 20, 17, 30), // Jan 20, 2025 (Monday) 5:30 PM
    2,
    '10:00',
    '18:30',
    [0],
    ['2025-01-21'] // Tuesday is holiday
  );
  Logger.log('TEST 6: Submit Mon 5:30 PM, TAT 2 hrs, Tue holiday');
  Logger.log('Expected: Jan 22, 2025 11:00 AM (Wednesday)');
  Logger.log('Got:      ' + formatResult(test6));
  Logger.log('');
  
  // TEST 7: Large TAT - 24 hours (3 full working days of 8.5 hrs)
  // Office = 8.5 hrs/day, 24 hrs = 2 full days + 7 hrs
  // Submit Mon 10 AM: Mon 8.5 + Tue 8.5 = 17 hrs, remaining 7 hrs
  // Wed 10:00 + 7 hrs = 5:00 PM
  var test7 = calculateTATLocal(
    new Date(2025, 0, 20, 10, 0), // Jan 20, 2025 (Monday) 10:00 AM
    24,
    '10:00',
    '18:30',
    [0],
    []
  );
  Logger.log('TEST 7: Submit Mon 10 AM, TAT 24 hrs (office 8.5 hrs/day)');
  Logger.log('Expected: Jan 22, 2025 05:00 PM (Wednesday)');
  Logger.log('Got:      ' + formatResult(test7));
  Logger.log('');
  
  Logger.log('======= TESTS COMPLETE =======');
}

/**
 * Local version of calculateTAT that accepts holidays array directly (for testing)
 */
function calculateTATLocal(submissionDateTime, tatHours, officeStartTime, officeEndTime, weekOffDays, holidaysArray) {
  
  var submission = (submissionDateTime instanceof Date) ? new Date(submissionDateTime.getTime()) : parseDateTime(submissionDateTime);
  var officeStart = parseTime(officeStartTime);
  var officeEnd = parseTime(officeEndTime);
  var holidays = holidaysArray || [];
  
  var tatMinutesRemaining = tatHours * 60;
  var currentDate = new Date(submission.getTime());
  
  // STEP 1: Position to valid starting point
  if (!isWorkingDay(currentDate, weekOffDays, holidays)) {
    currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
  } else {
    var currentMinutes = currentDate.getHours() * 60 + currentDate.getMinutes();
    var startMinutes = officeStart.hours * 60 + officeStart.minutes;
    var endMinutes = officeEnd.hours * 60 + officeEnd.minutes;
    
    if (currentMinutes >= endMinutes) {
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    } else if (currentMinutes < startMinutes) {
      currentDate.setHours(officeStart.hours, officeStart.minutes, 0, 0);
    }
  }
  
  // STEP 2: Consume TAT minutes across working days
  while (tatMinutesRemaining > 0) {
    if (!isWorkingDay(currentDate, weekOffDays, holidays)) {
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    }
    
    var currentMinutes = currentDate.getHours() * 60 + currentDate.getMinutes();
    var endMinutes = officeEnd.hours * 60 + officeEnd.minutes;
    var startMinutes = officeStart.hours * 60 + officeStart.minutes;
    
    if (currentMinutes < startMinutes) {
      currentDate.setHours(officeStart.hours, officeStart.minutes, 0, 0);
      currentMinutes = startMinutes;
    }
    
    var remainingTodayMinutes = endMinutes - currentMinutes;
    
    if (remainingTodayMinutes <= 0) {
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
      continue;
    }
    
    if (tatMinutesRemaining <= remainingTodayMinutes) {
      currentDate.setMinutes(currentDate.getMinutes() + tatMinutesRemaining);
      tatMinutesRemaining = 0;
    } else {
      tatMinutesRemaining -= remainingTodayMinutes;
      currentDate = getNextWorkingDayStart(currentDate, officeStart, weekOffDays, holidays);
    }
  }
  
  return currentDate;
}

/**
 * Format a Date for display EVERYWHERE in FMS (table cells, popup form's
 * read-only Planned value, pending-list "Planned:" caption, "Last Updated"
 * timestamp, etc).
 *
 * Format: "DD Mon YY HH:MM:SS" (24-hour time), e.g. "24 Jul 26 17:43:32"
 *
 * EXCEPTION: if the time portion is exactly 00:00:00 (i.e. the value is a
 * pure date with no meaningful time - e.g. a plain date entered without a
 * time), the time is omitted entirely and only "DD Mon YY" is shown.
 */
function formatResult(date) {
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var pad2 = function (n) { return ('0' + n).slice(-2); };

  var dd = pad2(date.getDate());
  var mon = months[date.getMonth()];
  var yy = pad2(date.getFullYear() % 100);
  var hh = pad2(date.getHours());
  var mm = pad2(date.getMinutes());
  var ss = pad2(date.getSeconds());

  var isMidnight = (date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0);
  if (isMidnight) {
    return dd + ' ' + mon + ' ' + yy;
  }

  return dd + ' ' + mon + ' ' + yy + ' ' + hh + ':' + mm + ':' + ss;
}
