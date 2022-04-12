function updatePM() {
  
  let ss = SpreadsheetApp.getActive();

  //Counts # of PMs on the new list in column A starting from cell A4
  let newPMs = ss.getRange("A7:A").getValues();                                               
  let newPMsNum = newPMs.filter(String).length;
  let newPMsRow = newPMsNum + 6
  let newPMsRange = ss.getRange("A7:A"+newPMsRow);
    newPMsRange.sort({column: 1, ascending: true});

  console.log("NewPMsRow = " + newPMsRow);

  //Counts # of PMs on the old list in column B starting from cell B4
  let oldPMs = ss.getRange("B7:B").getValues();                             
  let oldPMsNum = oldPMs.filter(String).length;
  let oldPMsRow = oldPMsNum + 6
  let oldPMsRange = ss.getRange("B7:B"+oldPMsRow);
    oldPMsRange.sort({column: 2, ascending: true});
                                               
  console.log("oldPMsRow = " + oldPMsRow);

  //Validation PM by PM. Who's new and who left the company. Adding and deleting pages.
  let arrayNewPMs = [];
  for (let i = 7; i <= newPMsRow; i++) {
    let currentNewPM = ss.getRange("A"+i+":A"+i).getValue();
    arrayNewPMs.push(currentNewPM); //Array for new list of PMs of Column A
  }
  // console.log(arrayNewPMs);

  let arrayOldPMs = [];
  for (let i = 7; i <= oldPMsRow; i++) {
    let currentOldPM = ss.getRange("B"+i+":B"+i).getValue();
    arrayOldPMs.push(currentOldPM); //Array for old list of PMs of Column B
  }
  // console.log(arrayOldPMs);

  let arrayNewEntries = []; 
  let menc = 0;
  for (let i = 0; i < arrayNewPMs.length; i++) {
    menc = 0
    for (let j = 0; j < arrayOldPMs.length; j++) {
      if (arrayNewPMs[i] === arrayOldPMs[j]) {          // Validates if a PM is new on the list
        menc = 1;
      }
    }
    if (menc === 0) {
      arrayNewEntries.push(arrayNewPMs[i]);             //If PM is new then pushes it into the array of arrayNewEntries
    }
  }

  console.log("New PMs to add: " + arrayNewEntries);

  let arrayOldEntries = [];
  for (let i = 0; i < arrayOldPMs.length; i++) {
    menc = 0
    for (let j = 0; j < arrayNewPMs.length; j++) {
      if (arrayOldPMs[i] === arrayNewPMs[j]) {          // Validates if a PM is not anymore on the list
        menc = 1;
      }
    }
    if (menc === 0) {
      arrayOldEntries.push(arrayOldPMs[i]);             //If PM is not on the list then pushes it into the array of arrayOldEntries
    }
  }

   console.log("Old PMs to delete: " + arrayOldEntries);
  
  let source = SpreadsheetApp.getActiveSpreadsheet();   //Add sheets for new PMs
  let sheet = source.getSheetByName("Example");
    for (let i = 0; i < arrayNewEntries.length; i++) {
      sheet.copyTo(source).setName(arrayNewEntries[i]);
      ss.getSheetByName(arrayNewEntries[i]).activate();
      ss.getRange("A3:H3").clearContent();
      ss.getRange("B1:B1").setValue("PM: "+arrayNewEntries[i])
    }

  //Delete spreadsheets of PMs that left BairesDev
  try {
    let sheetToDelete = ""
    for (let i = 0; i < arrayOldEntries.length; i++) {
      sheetToDelete = ss.getSheetByName(arrayOldEntries[i]);
      ss.deleteSheet(sheetToDelete);
    }
  console.log("Hojas Borradas: " + arrayOldEntries);
  } catch(error) {console.log("There's no PMs that left BairesDev on the list")};

  //Update directory list and update sheet link for each PM
  ss.getSheetByName('Information').activate();
  console.log(arrayNewPMs);
  ss.getRange("A7:A" + newPMsRow).clearContent();
  ss.getRange("B7:B" + oldPMsRow).clearContent();
  for (let i = 0; i < arrayNewPMs.length; i++) {
    let pmSheet = SpreadsheetApp.getActive().getSheetByName(arrayNewPMs[i]);
    console.log("i = " + i);
    console.log("arrayNewPMs: " + arrayNewPMs[i]);

    let rangeToAddLink = SpreadsheetApp.getActive().getSheetByName('Information').getRange('B' + (i + 7));
    let richText = SpreadsheetApp.newRichTextValue()
      .setText(arrayNewPMs[i])
      .setLinkUrl("#gid=" + pmSheet.getSheetId())
      .build();
    rangeToAddLink.setRichTextValue(richText);
    console.log("Link " + i + " added");
  }


}
