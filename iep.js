//globals
var hex_light_red = '#FFCCCC';
var hex_red = '#E66666';
var hex_white = '#FFFFFF';
var hex_green = '#6AA84F';

//validate entry made by user in entry form
var ssheet = SpreadsheetApp.getActiveSpreadsheet();
var form = ssheet.getSheetByName('Form');
var db = ssheet.getSheetByName('DB');
var students = ssheet.getSheetByName("Students");
var query = ssheet.getSheetByName("Query");
var ui = SpreadsheetApp.getUi(); 

//dates
var settings = ssheet.getSheetByName('Settings');
var today = new Date();
var year_start = SpreadsheetApp.getActive().getRangeByName('year_start').getValue();
var q2_start = SpreadsheetApp.getActive().getRangeByName('q2_start').getValue();
var q3_start = SpreadsheetApp.getActive().getRangeByName('q3_start').getValue();
var q4_start = SpreadsheetApp.getActive().getRangeByName('q4_start').getValue();
var year_end = SpreadsheetApp.getActive().getRangeByName('year_end').getValue();

//get form ranges
var input_name = form.getRange('C7');
var input_date = form.getRange('C15');
var input_type = form.getRange('G15');
var input_time = form.getRange('E15');
var input_notes = form.getRange('C19');
var input_image = form.getRange('C38');


function onEdit(e) {
  validateEntry(false);
  const range = e.range;
  if (range.getA1Notation() === 'C7') {
    updateStudentCards();
  }
  if (range.getA1Notation() === 'C38') {
    range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }
}

function resetColors() {
  //default field colors
  input_name.setBackground(hex_white);
  input_date.setBackground(hex_white);
  input_type.setBackground(hex_white);
  input_time.setBackground(hex_white);
  input_notes.setBackground(hex_white);
  input_image.setBackground(hex_white);
}

function addName() {
  var name_to_add = form.getRange('D5').getValue();

  var input_names = input_name.getValue();
  var save_names = input_names.split(',');
  
  // Trim any whitespace around the array elements
  save_names = save_names.map(function(item) {
    return item.trim();
  });
    
  for (var i = 0; i < save_names.length; i++) {
    if (save_names[i] === name_to_add) {
      ui.alert(name_to_add + ' has already been selected.');
      return false;
    }
  }
  
  if (input_name.getValue().trim().length === 0) {
    input_name.setValue(name_to_add)
  }
  else {
    input_name.setValue(input_name.getValue() + ', ' + name_to_add)
  }

  updateStudentCards();
  validateEntry(false);
}

function validateEntry(alert, save_attempted=false) { //show alerts

  resetColors();

  //required fields
  var req_data = [input_name, input_date, input_type, input_time]
  var valid_input = true;

  if (save_attempted) {

    for (var i = 0; i < req_data.length; i++) {
      if (req_data[i].isBlank()) {
        if (alert) {
          ui.alert('Please complete the form.')
        }
        valid_input = false;
        //highlight incomplete required field
        switch (i) {
          case 0:
            input_name.setBackground(hex_light_red);
            break;
          case 1:
            input_date.setBackground(hex_light_red);
            break;
          case 2:
            input_type.setBackground(hex_light_red);
            break;
          case 3:
            input_time.setBackground(hex_light_red);
            break;
          default:
            break;
        }

      }

      //field validations
    
      if (!valid_input) {
        return false;
      }
    }

  }

  return true;
}


function saveEntry(){
  if (!validateEntry(true, true)) {
    //display error if necessary and stop
    return false;
  }

  //TODO: check for exact duplicate entry on a given date and have user confirm they want both

  //handle multiple names
  var input_names = input_name.getValue();
  var save_names = input_names.split(',');
  
  // Trim any whitespace around the array elements
  save_names = save_names.map(function(item) {
    return item.trim();
  });

  for (var i = 0; i < save_names.length; i++) {

    //check if name exists in Students table
    if (save_names[i] === '' || !studentExists(save_names[i])) {
      ui.alert('"' + save_names[i] + '"' + ' does not exist in the Students table');
    }
    else {
      var input_data = [[save_names[i], input_date.getValue(), input_type.getValue(), input_time.getValue(), input_notes.getValue()]];
      
      //get next empty row in DB
      var row = nextEmptyRow('DB');
      var rng = db.getRange('A' + row + ":E" + row);
      rng.setValues(input_data);

      //add image to DB - only way is through copy/paste as of 8/7/2024
      input_image.copyTo(db.getRange('F' + row), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      db.getRange('F' + row).clearFormat();
      
      

      //Confirm Save in alert
      ui.alert('Record saved for ' + save_names[i]);
    }
  }

  //refresh Student Card
    updateStudentCards();
}

function studentExists(check_name) {
  var student_names = students.getRange('A2:A').getValues();
  for (var i = 0; i < student_names.length; i++) {
    if (student_names[i][0] === check_name) {
      return true;
    }
  }
  return false;
}


function clearForm() {
  input_name.clearContent();
  input_date.clearContent();
  input_type.clearContent();
  input_time.clearContent();
  input_notes.clearContent();
  input_image.clearContent();
}


function resetNames() {
  input_name.clearContent();
  updateStudentCards();
  validateEntry(false);
}


//search
card_index = 0;

function updateStudentCards() {
  //reset cards
  form.getRange('J:Q').clear().setBackground('#B7B7B7');


  var input_names = input_name.getValue(); //form.getRange('C3').getValue();

  if (input_names === "") {
    return
  }

  //handle multiple names
  var search_names = input_names.split(',');
  
  // Trim any whitespace around the array elements
  search_names = search_names.map(function(item) {
    return item.trim();
  });
  

  for (var i = 0; i < search_names.length; i++) {
    if (search_names[i] === "") {
      continue;
    }
    var student_config = getStudentConfig(search_names[i]);
    try {
      student_config.shift(); //removes name from returned values
    } catch (error) {
      //this means the name input does not exist in the Students table
      Logger.log(error);
      continue;
    }
    

    //get period start and end values
    var direct_time = student_config[0];
    var direct_unit = student_config[1];
    var indirect_time = student_config[2];
    var indirect_unit = student_config[3];

    var direct_period = getDateBounds(direct_unit);
    var indirect_period = getDateBounds(indirect_unit);

    var direct_completed_time = sumTimeByType(search_names[i], 'direct', direct_period[0], direct_period[1]);
    var indirect_completed_time = sumTimeByType(search_names[i], 'indirect', indirect_period[0], indirect_period[1]);

    var card_first_row = card_index * 12 + 2;
    var card_name_row = card_index * 12 + 3;
    var card_title_row = card_index * 12 + 5;
    var card_total_row = card_index * 12 + 7;
    var card_completed_row = card_index * 12 + 9;
    var card_remaining_row = card_index * 12 + 11;
    var card_last_row = card_index * 12 + 12;
    var card_first_col = 'J';
    var card_label_col = 'K';
    var card_direct_col = 'L';
    var card_direct_unit_col = 'M';
    var card_indirect_col = 'O';
    var card_indirect_unit_col = 'P';
    var card_last_col = 'Q';
    
    //format card
    form.getRange(card_first_col + card_first_row + ':' + card_last_col + card_last_row)
      .setBackground(null)
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setFontSize(12)
    form.getRange(card_first_col + card_name_row + ':' + card_last_col + card_name_row)
      .mergeAcross()
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
    form.getRange(card_direct_unit_col + card_title_row).setFontSize(10);
    form.getRange(card_indirect_unit_col + card_title_row).setFontSize(10);
    form.setRowHeight(card_first_row - 1, 10);
    form.setRowHeight(card_title_row + 1, 10);

    //populate card
    form.getRange(card_first_col + card_name_row).setValue(search_names[i]);
    form.getRange(card_direct_col + card_title_row).setValue('Direct');  
    form.getRange(card_direct_unit_col + card_title_row).setValue('(min/' + direct_unit + ')');  
    form.getRange(card_indirect_col + card_title_row).setValue('Indirect');
    form.getRange(card_indirect_unit_col + card_title_row).setValue('(min/' + indirect_unit + ')'); 
    form.getRange(card_label_col + card_total_row).setValue('Total');
    form.getRange(card_label_col + card_completed_row).setValue('Completed');
    form.getRange(card_label_col + card_remaining_row).setValue('Remaining');
    form.getRange(card_direct_col + card_total_row).setValue(direct_time); 
    form.getRange(card_indirect_col + card_total_row).setValue(indirect_time); 
    form.getRange(card_direct_col + card_completed_row).setValue(direct_completed_time); 
    form.getRange(card_indirect_col + card_completed_row).setValue(indirect_completed_time); 
    form.getRange(card_direct_col + card_remaining_row).setValue(direct_time - direct_completed_time); 
    form.getRange(card_indirect_col + card_remaining_row).setValue(indirect_time - indirect_completed_time); 

    if (direct_time - direct_completed_time < 0) {
      form.getRange(card_direct_col + card_remaining_row).setFontColor(hex_red);
    }
    if (direct_time - direct_completed_time === 0) {
      form.getRange(card_direct_col + card_remaining_row).setFontColor(hex_green);
    }

    if (indirect_time - indirect_completed_time < 0) {
      form.getRange(card_indirect_col + card_remaining_row).setFontColor(hex_red);
    }
    if (indirect_time - indirect_completed_time === 0) {
      form.getRange(card_indirect_col + card_remaining_row).setFontColor(hex_green);
    }

    card_index += 1;
  }

  

}

function getStudentConfig(name) {
  var data = students.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === name) {  
      return data[i];  // Return the entire row data
    }
  }
}

function subtractOneDay(date) {
  if (!(date instanceof Date)) {
    throw new Error("Input must be a valid Date object.");
  }

  var newDate = new Date(date);
  newDate.setDate(newDate.getDate() - 1);
  return newDate;
}




function getQuarterBounds() {
  if (today >= q2_start) {
    if (today >= q3_start) {
      if (today >= q4_start) {
        return [q4_start, year_end];
      }
      else {
        return [q3_start, subtractOneDay(q4_start)]
      }
    }
    else {
      return [q2_start, subtractOneDay(q3_start)]
    }
  }
  else {
    return [year_start, subtractOneDay(q2_start)]
  }
}

function getSemesterBounds() {
    if (today >= q3_start) {
      return [q3_start, year_end];
    }
    else {
      return [year_start, subtractOneDay(q3_start)];
    }
}

function isInDateRange(checkDate, startDate, endDate) { //confirm that Settings have the correct year configured
  if (!(checkDate instanceof Date) || !(startDate instanceof Date) || !(endDate instanceof Date)) {
    throw new Error("All inputs must be valid Date objects.");
  }

  // Ensure the dates are in the correct range
  var date = new Date(checkDate);
  var start = new Date(startDate);
  var end = new Date(endDate);
  
  // Adjust time to be at the start of the day for comparisons
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);
  date.setHours(0, 0, 0, 0);

  return date >= start && date <= end;
}


function getDateBounds(period) { //period = month, quarter, semester, or year
  
  if (!isInDateRange(today, year_start, year_end)) {
    ui.alert("Today is not between the School Year Start and End Dates on the Settings tab.");
    throw new Error("Dates outside year bounds");
  }

  switch (period) {
    case 'month':
      var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
      var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      return [firstDay, lastDay];
    case 'quarter':
      var bounds = getQuarterBounds();
      return [bounds[0], bounds[1]];
    case 'semester':
      var bounds = getSemesterBounds();
      return [bounds[0], bounds[1]];
    case 'year':
      return [year_start, year_end];
    case 'N/A':
      return ['N/A', 'N/A'];
    default:
      break;
  }
}

function sumTimeByType(name, type, startDate, endDate) {
  
  // Get the data range (assuming data starts from A1)
  var range = db.getDataRange();
  var values = range.getValues();
  
  // Initialize sums
  var directSum = 0;
  var indirectSum = 0;
  
  // Iterate through the rows
  for (var i = 1; i < values.length; i++) { // Start from 1 to skip header row
    var rowName = values[i][0];
    var rowDate = new Date(values[i][1]);
    var rowType = values[i][2];
    var rowTime = parseFloat(values[i][3]);
    
    // Check if the row matches the given name
    if (rowName === name) {
      // Check if the date is within the specified range
      if (rowDate >= new Date(startDate) && rowDate <= new Date(endDate)) {
        // Sum the time based on type
        if (rowType === 'Direct') {
          directSum += rowTime;
        } else if (rowType === 'Indirect') {
          indirectSum += rowTime;
        }
      }
    }
  }

  // Return results as an object
  if (type === 'direct') {
    return directSum
  }
  else {
    return indirectSum
  }
}



















