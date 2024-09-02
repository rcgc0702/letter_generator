var studentLetter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LevelsLetter');
var collection = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Collection');
var examLetter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExamLetter');
var infoLetter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('InfoLetter');
var staffChannels1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StaffChannels');
var numberOfAssigners = staffChannels1.getRange(5,6).getValue()
var listOfChannels = staffChannels1.getRange(4,6).getValue()

/* Latest update 2023.05.19 */

function onEdit() {
  var user = Session.getActiveUser().getEmail();
  var today = new Date();
  infoLetter.getRange(1, 3).setValue("Last user: " + user.substring(0,3) + "... Last Activity: " + todayEdit());
  studentLetter.getRange(1, 3).setValue("Last user: " + user.substring(0,3) + "... Last Activity: " + todayEdit());
  examLetter.getRange(1, 3).setValue("Last user: " + user.substring(0,3) + "... Last Activity: " + todayEdit());
}

function onOpen(e) {

  examLetter.getRange(15,2,2).clearContent()
  infoLetter.getRange(5,2,2).clearContent()
  studentLetter.getRange(10,2,2).clearContent()
  studentLetter.getRange(10,4,1).clearContent()
}

function createLetter() {
  
  /*
  * This is the function for Level 1 to 5
  */
  
  studentLetter.getRange(10, 2).clearContent();
  var wordsArray = new Array();
  var lcol = collection.getLastRow();
  var crazyArray = new Array();

  for (var i = 0; i < lcol; i++) {
    wordsArray[i] = collection.getRange(i+1, 2).getValue();
  }

  var link = studentLetter.getRange(4, 2).getValue();
  var urilink = encodeURIComponent(link);
  var emptyFields = 0;
  
  for (var y = 0; y < 8; y++) {
    crazyArray[y] = studentLetter.getRange(y+1, 2).getValue();
    emptyFields += checkIfEmpty(crazyArray[y]);
  }

  if (crazyArray[6] == crazyArray[7]) {
      crazyArray[8] = 'me'
      crazyArray[9] = 'I'
  } else {
      crazyArray[8] = 'your sensei'
      crazyArray[9] = 'Your sensei'
  }

  if (abortCode(emptyFields)) return;
  var lettersArray = new Array();
  var lengthOfArray;
  var concatLetter = "";
  var z = 0;
  
  switch(studentLetter.getRange(2, 2).getValue()) {
    case 'Level 1':
      Logger.log(true);
      lettersArray = [0,5,6,13]; 
      break;
    case 'Level 2':
      lettersArray = [1,5,6,13];
      break;
    case 'Level 3':
      lettersArray = [2,7,13];
      break;
    case 'Level 4':
      lettersArray = [3,5,6,13];
      break;
    case 'Level 5':
      lettersArray = [12];
      break;
  }
  
  lengthOfArray = wordsArray.length;
  for (var g = 0; g < lettersArray.length; g++) {
    concatLetter += wordsArray[lettersArray[g]];
  }
  
  var replacors = ["*studentID*","*level*","*channel*","*link*",
                   "*videoNo*","*partNo*","*sensei*","*assigner*", '*ref1*','*ref2*'];
  var changedLetter = "";

  for(var w = 0; w < replacors.length; w++) {
    changedLetter = concatLetter;
    concatLetter = replaceMultiples(changedLetter,replacors[w],crazyArray[w]);
  }

  studentLetter.getRange(10,2).setValue('NSSA Assignment: ' + studentLetter.getRange(2,2).getValue())
  studentLetter.getRange(10,4).setValue(studentLetter.getRange(5,2).getValue() + "." + studentLetter.getRange(6,2).getValue() + " = " + studentLetter.getRange(1,2).getValue())

  concatLetter = replaceMultiples(concatLetter,"*urilink*",urilink);
  display(studentLetter,concatLetter,11,2);
  studentLetter.getRange(1, 2, 8).clearContent();

}

function createGradLetter() {
  
  /*
  * This is the function for the grad letter
  */
  
  var gradLetterTemplate = collection.getRange(9, 2).getValue();
  var emptyFields = 0;
  var fieldValues = new Array();
  var urilink = encodeURIComponent(examLetter.getRange(3, 2).getValue());
  var urilink2 = encodeURIComponent(examLetter.getRange(7, 2).getValue());

  for (var r = 0; r < 13; r++) {
    fieldValues[r] = examLetter.getRange(r+1, 2).getValue();
    emptyFields += checkIfEmpty(fieldValues[r]); 
  }

  if (fieldValues[10] == fieldValues[12] || fieldValues[11] == fieldValues[12]) {
    fieldValues[13] = ''
    fieldValues[14] = ''
    fieldValues[15] = 'notify both panelists immediately'
  } else {
    fieldValues[13] = ' and me'
    fieldValues[14] = 'www.viki.com/users/' + fieldValues[12]
    fieldValues[15] = 'notify me and your final panelist immediately'
  }

  if (abortCode(emptyFields)) return;
  var replacors = ["*studentID*","*channel*","*link*","*videoNo*",
                  "*partNo*","*segmentRange*","*channel2*","*link2*","*videoNo2*",
                  "*partNo2*","*sensei1*","*sensei2*","*assigner*",'*addTheMe*','*finalLink*','*messageUs*'];
  var populateInfo = "";

  for(var w = 0; w < replacors.length; w++) {
    populateInfo = gradLetterTemplate;
    gradLetterTemplate = replaceMultiples(populateInfo,replacors[w],fieldValues[w]);
  }
  gradLetterTemplate = replaceMultiples(gradLetterTemplate,"*twoWeeksLater*",date14DaysLater());
  gradLetterTemplate = replaceMultiples(gradLetterTemplate,"*urilink*",urilink);
  gradLetterTemplate = replaceMultiples(gradLetterTemplate,"*urilink2*",urilink2);
  display(examLetter,gradLetterTemplate,16,2);
  examLetter.getRange(15,2).setValue('NSSA Video Exam: ' + examLetter.getRange(1,2).getValue())
  examLetter.getRange(1,2,13).clearContent();
}

function createInfoLetter() {
  
  /*
  * This is the function for the Information Letters
  */
  
  var answer = infoLetter.getRange(2, 2).getValue();
  var studentID = infoLetter.getRange(1, 2).getValue();
  //var sensei = infoLetter.getRange(3, 2).getValue(); // removed 3/21/2023
  var assigner = infoLetter.getRange(3, 2).getValue(); //changed to a differenct cell 3/21/2023 from (4, 2)
  var letter = "";
  var emptyFields = 0;
  
  emptyFields += checkIfEmpty(studentID);
  emptyFields += checkIfEmpty(assigner);
  
  switch(answer) {
  case 'Exam information':
      letter = collection.getRange(11, 2).getValue();
      infoLetter.getRange(5,2).setValue('NSSA Exam Information')
      break;
  case 'Sandbox':
      letter = collection.getRange(10, 2).getValue();
      infoLetter.getRange(5,2).setValue('NSSA Information Letter')
      break;
    default:
      emptyFields += 1;
      
  }
  
  if (abortCode(emptyFields)) return;
  var newLetter = replaceMultiples(letter,"*studentID*",studentID);
  // newLetter = replaceMultiples(newLetter,"*sensei*",sensei); //removed
  newLetter = replaceMultiples(newLetter,"*assigner*",assigner);
  display(infoLetter,newLetter,6,2); // changed from (infoLetter,newLetter,6,2) 3/21/2023
  infoLetter.getRange(1,2,3).clearContent(); // changed from (1,2,4) 3/21/2023

  infoLetter.getRange(4,6).setValue("")

  /*infoLetter.getRange(6,2).setFontFamily('Cambria') */
}

function replaceMultiples(source, lookupString, replacement) { 
  // var stillGo = true;
  // var banjo;
  // while (stillGo) {
  //   banjo = source;
  //   source = source.replace(lookupString,replacement);
  //   if (banjo == source) {
  //     stillGo = false;
  //   }
  // }

  source = source.replaceAll(lookupString,replacement)

  return source;
}

function checkIfEmpty(stringToCheck) {
  if(stringToCheck == "") {
    return 1;
  }
  return 0;
}

function date14DaysLater() {
  var month = createMonthsArray();
  var today = new Date();
  today.setDate(today.getDate() + 14); 
  return boldDate(month[today.getMonth()]) +" "+ boldNumbers(today.getDate().toString()) + ", " + boldNumbers(today.getFullYear().toString());
}

function todayEdit() {
  var month = createMonthsArray();
  var today = new Date();
  var hours = today.getHours() > 12 ? today.getHours()-12 : today.getHours();
  var ampm = today.getHours() > 12 ? " pm " : " am ";
  var minutes = today.getMinutes() < 10 ? "0" + today.getMinutes() : today.getMinutes();
  today.setDate(today.getDate()); 
  return month[today.getMonth()] +" "+today.getDate() + " @ " + hours + ":" + minutes + ampm + Session.getScriptTimeZone();

}

function createMonthsArray() {
  var month = new Array();
  month[0] = "January";
  month[1] = "February";
  month[2] = "March";
  month[3] = "April";
  month[4] = "May";
  month[5] = "June";
  month[6] = "July";
  month[7] = "August";
  month[8] = "September";
  month[9] = "October";
  month[10] = "November";
  month[11] = "December";
  return month;
}

function display(sheet,value_1,row,col) {
  sheet.getRange(row, col)
  .setWrap(true)
  .activateAsCurrentCell()
  .setValue(value_1);
}

function abortCode(emptyFields) {
  if(emptyFields > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("You have "+ emptyFields +" empty field(s). Operation aborted.", ui.ButtonSet.OK);
    return true;
  }
  return false;
}

function openTheSideBar() {
    var html = HtmlService.createHtmlOutputFromFile('SideBar')
      .setTitle('Input Letter Parameters');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

function openGradSideBar() {
    var html = HtmlService.createHtmlOutputFromFile('graduation')
      .setTitle('Grad Letter Parameters');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

function openInfoLetter() {
    var html = HtmlService.createHtmlOutputFromFile('info_letter')
      .setTitle('Info Letter Parameters');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


function processData(theDataFromInput) {

  studentLetter.getRange(1,2).setValue(theDataFromInput.name)
  studentLetter.getRange(2,2).setValue(theDataFromInput.theLevels)
  studentLetter.getRange(3,2).setValue(theDataFromInput.theChannel)
  studentLetter.getRange(4,2).setValue(theDataFromInput.link1)
  studentLetter.getRange(5,2).setValue(theDataFromInput.videoNo)
  studentLetter.getRange(6,2).setValue(theDataFromInput.partNo)
  studentLetter.getRange(7,2).setValue(theDataFromInput.sensei)
  studentLetter.getRange(8,2).setValue(theDataFromInput.assigner)

  createLetter()
}

function processGradData(theDataFromInput) {

  examLetter.getRange(1,2).setValue(theDataFromInput.name)
  examLetter.getRange(2,2).setValue(theDataFromInput.theChannel)
  examLetter.getRange(3,2).setValue(theDataFromInput.link1)
  examLetter.getRange(4,2).setValue(theDataFromInput.videoNo)
  examLetter.getRange(5,2).setValue(theDataFromInput.partNo)
  examLetter.getRange(6,2).setValue(theDataFromInput.segmentRange)
  examLetter.getRange(7,2).setValue(theDataFromInput.theChannel2)
  examLetter.getRange(8,2).setValue(theDataFromInput.link2)
  examLetter.getRange(9,2).setValue(theDataFromInput.videoNo2)
  examLetter.getRange(10,2).setValue(theDataFromInput.partNo2)
  examLetter.getRange(11,2).setValue(theDataFromInput.sensei)
  examLetter.getRange(12,2).setValue(theDataFromInput.sensei2)
  examLetter.getRange(13,2).setValue(theDataFromInput.assigner)

  createGradLetter()
}

function processTheInfoLetter(theDataFromInput) {

  infoLetter.getRange(1,2).setValue(theDataFromInput.name)
  infoLetter.getRange(2,2).setValue(theDataFromInput.theLevel)
  infoLetter.getRange(3,2).setValue(theDataFromInput.assigner)

  createInfoLetter()
}

function createTeachersArray(nameOfSelect) {

  //var itemToReturn;

  switch(nameOfSelect) {
    case 'sensei':
      var lcol = staffChannels1.getLastRow();
      var itemToReturn = staffChannels1.getRange(2,1,lcol).getValues();
      //return theTeachersArray;
      break;
    case 'assigner':
      var itemToReturn = staffChannels1.getRange(2,3,numberOfAssigners).getValues();
      //return assignerArray;
      break;
    case 'theChannel':
      var itemToReturn = staffChannels1.getRange(2,2,listOfChannels).getValues();
      //return channelArray;
      break;
  }

  return itemToReturn;

}

function abortOpenMessage() {

    var emptyFields = 0;
    var heyArray = new Array();
  
  for (var y = 0; y < 12; y++) {
    heyArray[y] = examLetter.getRange(y+1, 2).getValue();
    emptyFields += checkIfEmpty(heyArray[y]);
  }

  return emptyFields;
}

function boldDate(theDate) {

  let text = theDate;
  
    function translate (char)
    {
        let diff;
        if (/[A-Z]/.test (char))
        {
            diff = "𝗔".codePointAt (0) - "A".codePointAt (0);
        }
        else
        {
            diff = "𝗮".codePointAt (0) - "a".codePointAt (0);
        }
        return String.fromCodePoint (char.codePointAt (0) + diff);
    }

    let newText = text.replace (/[A-Za-z0-1]/g, translate);

  return newText;
}

function boldNumbers(theDate) {

  let text = theDate;
  
    function translate (char)
    {
        let diff;
        if (/[A-Z]/.test (char))
        {
            diff = "𝟎".codePointAt (0) - "0".codePointAt (0);
        }
        else
        {
            diff = "𝟎".codePointAt (0) - "0".codePointAt (0);
        }
        return String.fromCodePoint (char.codePointAt (0) + diff);
    }

    let newText = text.replace (/[0-9]/g, translate);

  return newText;
}
