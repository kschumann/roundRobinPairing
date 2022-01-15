var ss = function(){
 return SpreadsheetApp.getActive();
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi(); 
  ui.createAddonMenu()
  .addItem('Create Pairs', 'openPrompt')
  .addItem('Print PDF','printablePdf')  
  .addToUi(); 
}

function onInstall(e) {
  onOpen(e);
}


function openPrompt(){
  try{
    /*create warning if sheets already exist*/
    var participantsSheet = ss().getSheetByName('Participants');
    var warning = "";
    participantsSheet ? warning = "CAREFUL! You already have a Participants sheet created.  Pressing OK will erase previously entered data on this sheet." : warning = "";
    
    /*Open Prompt Dialogue box*/
    var ui = SpreadsheetApp.getUi(); 
    var response = ui.prompt("How Many Participants will be Pairing Up?", "Enter any number up to 40. " + warning, ui.ButtonSet.OK_CANCEL);
    
    /****VALIDATIONS****/
    if (response.getSelectedButton() == ui.Button.OK) {
      var responseInt = parseInt(response.getResponseText());
      if(isNaN(responseInt)){//Check to ensure that response is a number
        ui.alert("Ooops!  There was an issue with your entry. You entered: " + response.getResponseText() + ", which is not a number. Please try again."); 
        return
        
      } else{//Cleanse Entry
        responseInt % 2 != 0 ? responseInt = responseInt +1: responseInt = responseInt + 0;
        responseInt < 0 ? responseInt = responseInt * -1 : responseInt = responseInt * 1;
        if(responseInt > 40 || responseInt == 0){
          ui.alert("Ooops! There was an issue with your entry. You entered: " + response.getResponseText() + ".  The number must be between 1 and 40. Please try again.");     
          return
        }
      }
      setUpRoundRobin(responseInt);
    } 
  } catch(e){
    catchError("An unexpected error occurred.  If this error persists, please contact the Add-On administrator. Details: " + e);
    console.error("An unexpected error occurred.  If this error persists, please contact the Add-On administrator. Details: " + e);

  }
}


function setUpRoundRobin(numbParticipants){
  try{
    /*Set Up Participant Sheet*/
    var participantSheet = ss().getSheetByName('Participants');
    if(!participantSheet){
      var participantSheet = ss().insertSheet(0).setName('Participants').activate();
    };
    participantSheet.clear();
    var participants = [['Participant List']];
    for(var i=1; i<=numbParticipants; i++){
      participants.push(['Participant'+i]);
    }
    participantSheet.getRange(1, 1, numbParticipants+1).setValues(participants);
    
    /*Format Participant Sheet*/
    participantSheet.getRange(1,1).setFontWeight('bold').setFontSize(24).setBackground('#000000').setFontColor('white');
    participantSheet.setColumnWidth(1, 250).setRowHeights(1, participants.length, 40);
    participantSheet.getRange(2, 1, participants.length-1, 1).setFontSize(18)
    
    /*Set up Pairings*/
    var pairingSheet = ss().getSheetByName('Pairings');
    if(!pairingSheet){
      var pairingSheet = ss().insertSheet(1).setName('Pairings');  
    };  
    
    pairingSheet.clear();
    pairingSheet.setHiddenGridlines(true);
    for(var k = 0; k<numbParticipants-1; k++){
      var col1 = [['=Participants!A2']];
      var col2 = [];
      for(var i=0; i<numbParticipants/2;i++){
        var row1= i+2-k > 2 ? i+2-k : i-k + numbParticipants+1;
        var rowN = numbParticipants+1-i-k > 2 ? numbParticipants-i-k+1 : numbParticipants*2-i-k;
        if(i != 0){col1.push(['=Participants!A' + row1])};
        col2.push(['=Participants!A' + rowN]);
      }
      pairingSheet.getRange(2, 1+k*2, numbParticipants/2).setValues(col1);
      pairingSheet.getRange(2, 2+k*2, numbParticipants/2).setValues(col2);
    }
    
    /*Format Pairing sheet so that each pairing fits into single letter sized sheet*/
    var lastCol = pairingSheet.getLastColumn();
    var lastRow = pairingSheet.getLastRow();
    
    //Adjust Font size depending on number of pairs
    if(numbParticipants<=22){
      pairingSheet.getRange(1, 1, lastRow, lastCol).setFontSize(36);
      pairingSheet.setRowHeights(1, lastRow, 75);
      pairingSheet.setColumnWidths(1, lastCol, 300);
    } else if(numbParticipants<=32){
      pairingSheet.getRange(1, 1, lastRow, lastCol).setFontSize(28);
      pairingSheet.setRowHeights(1, lastRow, 50);
      pairingSheet.setColumnWidths(1, lastCol, 300);
    } else  {
      pairingSheet.getRange(1, 1, lastRow, lastCol).setFontSize(24);
      pairingSheet.setRowHeights(1, lastRow, 30);
      pairingSheet.setColumnWidths(1, lastCol, 300);
    }
    
    //Shade alternate rows
    for(var j=0; j<numbParticipants/2; j++){
      if(j % 2 == 0){
        pairingSheet.getRange(j+2,1,1,lastCol).setBackground('#efefef');
      }
    }
    
    //Format Headers and add placeholder text
    var groupTag = "Pairing ";
    for(var m=0; m<lastCol; m++){
      if(m % 2 != 0){
        var number = String((m+1)/2);
        pairingSheet.getRange(1,m).setValue(groupTag + number);
        pairingSheet.getRange(1,m,1,2)
        .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        .merge()
        .setHorizontalAlignment('center')
        .setFontWeight('bold');
      }
    }
    participantSheet.activate();
    console.info("RoundRobin Setup Completed. Number of Participants: " + numbParticipants);
  } catch(e){
    catchError("An unexpected error occurred.  If this error persists, please contact the Add-On administrator. Details: " + e);
    console.error("An unexpected error occurred.  If this error persists, please contact the Add-On administrator. Details: " + e);

  }
}


function printablePdf() {
  try{
  var ss = SpreadsheetApp.getActive();
  var sheetId = ss.getId();
  var sheet = ss.getSheetByName('Pairings');    
  var gid = sheet.getSheetId();
  var pdfOpts = '&size=Letter&fzr=true&portrait=true&fitw=false&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=true&attachment=false&gid='+gid;
  //var url = "https://docs.google.com/spreadsheets/d/" + sheetId + "/export?format=pdf" + pdfOpts;
  var url = ss.getUrl().replace(/edit[^]*/, '') + 'export?format=pdf' + pdfOpts;
    
  var html = "<script>window.open('" + url + "','_blank');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html).setHeight(1); 
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening PDF . . .');  
  console.info("PDF Created");
  }
    catch(e){
      catchError("Ooops!  The add-on ran into an issue.  The Pairings Sheet could not be printed.  Make sure that you have created the 'Pairings' sheet using the Create Pairs function. Do not rename the sheet.  Here are the error details: " + e);
      console.error("Ooops!  The add-on ran into an issue.  The Pairings Sheet could not be printed.  Make sure that you have created the 'Pairings' sheet using the Create Pairs function. Do not rename the sheet.  Here are the error details: " + e);

    }
  }

/****Error Handling*****/
  
function catchError(message){
  var ui = SpreadsheetApp.getUi(); 
  var response = ui.alert(message);
}

