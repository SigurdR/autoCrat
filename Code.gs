var ui = SpreadsheetApp.getUi();

function onOpen(e){
  // Create menu options
  ui.createAddonMenu()
    .addSubMenu(ui.createMenu("Admin")
      .addItem("AutoCrat Setup", "autoCratSetup"))
    .addToUi();
};
                
function autoCratSetup() {

//***************************************************************************//
//autoCratSetup()                                                            //
//Initialize all the scriptProperties,                                       //
//create and open getTemplate html on client side which allow user to select //
//a document template to merge                                               //
//***************************************************************************//

  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.deleteAllProperties();

  var html = HtmlService.createTemplateFromFile('getTemplate.html')
    .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle('AutoCrat Setup - Step 1');
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function checkTaskExist(fileId) {

//***************************************************************************//
//checkTaskExist(fileId)                                                     //
//check if a merge task exist by looking for the template field ID & spread  //
//sheet name in the autoCrat_Parameters sheet. If an identical merge task is //
//found, ask if the user want to remove the task previously set and setup a  //
//new task -> getMappingString(), or leave it as it is.                      //
//***************************************************************************//

  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  var paraSheetName = "autoCrat_Parameters";
  var sParaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(paraSheetName);
  
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('fileId', fileId);
  scriptProperties.setProperty('sheetName', sheetName);
  
  if (sParaSheet) {
  
    var dataParameter = sParaSheet.getDataRange().getValues();
   
    dataParameter.splice(0, 1)
    
    var matchTask = dataParameter.filter(function (row) {
        
      return (row[1] === fileId && row[2] === sheetName);
    });
    
    if (matchTask.length > 0) {
    
      scriptProperties.setProperty('matchTaskId', matchTask[0][0]);
      
      var template = HtmlService.createTemplateFromFile("removeExistTask.html");
      
      var html = template
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('AutoCrat Setup - Step 1.5');
        
      SpreadsheetApp.getUi().showSidebar(html);
    }
    else {
    
      getOffset();
    }
  }
  else {
  
    getOffset();
  }
}

function getOffset() {

//***************************************************************************//
//  getOffset()                                                              //
//  Check the value of taskId, if it is not empty, then find in the sheet    //
//  auto_Crat_Parameters the task with that taskId and remove it. If it is   //
//  empty, skip the delete task Id procedure. Get the data range of the      //
//  target sheet and pass the number of rows to client side for defining the //
//  location of header row and data row                                      //  
//***************************************************************************//
  
  var paraSheetName = "autoCrat_Parameters"
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var taskId = scriptProperties.getProperty('matchTaskId');
  var fileId = scriptProperties.getProperty('fileId');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // remove task if the matchTaskId in scriptProperties is not empty
  if (taskId) {
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sParaSheet = ss.getSheetByName(paraSheetName);
    
    var dataParameter = sParaSheet.getDataRange().getValues();
    
    for (var rowNumber = 0; rowNumber < dataParameter.length; ++rowNumber) {
    
      if (dataParameter[rowNumber][0] === taskId) {
      
        var matchRow = rowNumber + 1;
        break;
      }
    }
    
    var rngList = [];
    
    var names = {};
   
    names['range'] = paraSheetName + "!A" + matchRow + ":H" + matchRow;
    
    rngList.push(names);
    
    var resource = {
   
      ranges: paraSheetName + "!A" + matchRow + ":H" + matchRow
    };
    
    Sheets.Spreadsheets.Values.batchClear(resource, ss.getId())
    sParaSheet.deleteRow(matchRow);
  };
  
  var sheetName = scriptProperties.getProperty('sheetName');
  
  var dataToMerge = ss.getSheetByName(sheetName).getDataRange().getValues();
  
  scriptProperties.setProperty('dataArray', JSON.stringify(dataToMerge));
  
  var dataRow = new Array();
  
  for (var i=0; i<=dataToMerge.length - 1; ++i) {
  
    dataRow.push(i+1);
  }
  
  var template = HtmlService.createTemplateFromFile("getOffset.html");
  
  template.dataRow = dataRow;
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 2');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

function getMappingString(headerRow, dataRow) {

//***************************************************************************//
//getMappingString(fileId)                                                   //
//write the value for header row and data row offset to scriptproperties. 
//Set fileId variable with the value get from     //
//getTemplate html on client side, get & validate the headers of the data to //
//be merged, get the merge fields in document, create an array for data      //
//types, and output to the 'data headers' and 'merge fields' and 'data type' //
//columns in getMappingString.html on client side for matching by users      //
//***************************************************************************//

  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty('headerRow', headerRow);
  scriptProperties.setProperty('dataRow', dataRow);
 
  var dataArray = JSON.parse(scriptProperties.getProperty('dataArray'));
  var sheetName = scriptProperties.getProperty('sheetName');
  var fileId = scriptProperties.getProperty('fileId');
  
  var headersPreRefine = new Array;
  
  for (var i=0; i<dataArray[0].length; ++i) {
  
    headersPreRefine.push(dataArray[headerRow-1][i]);
  }
  
  scriptProperties.setProperty('headersPreRefine', JSON.stringify(headersPreRefine));
  
  var headers = new Array();
  var headersIndex = new Array();
  
  //check if the heading is purely numeric
  for (var h=0; h<headersPreRefine.length; h++) {
    if (autoCrat_normalizeHeader(headersPreRefine[h])=="") {
      Browser.msgBox("Ooops! You must have an illegal header value in column " + (h+1) + ".  Headers cannot be blank or purely numeric.  Please fix.");
      return;
    }
    else {
      //skip headers with the "AutoCrat_" prefix"
      if ((headersPreRefine[h].indexOf("AutoCrat_")) < 0) {
        headers[h] = headersPreRefine[h];
        headersIndex[h] = headersPreRefine.indexOf(headersPreRefine[h]);
      }
    }
  }
  
  var mergeFields = autoCrat_fetchDocFields(fileId);
  
  // remove << >> from mergeFields
  var selectFields = autoCrat_prepareMergeFields(mergeFields);
  
  var dataTypes = new Array();
  
  dataTypes.push("Standard");
  dataTypes.push("Photos");
  dataTypes.push("Links");
  dataTypes.push("Checkboxes");
  
  var template = HtmlService.createTemplateFromFile("getMappingString.html");
  
  template.headers = headers;
  template.headersIndex = headersIndex;
  template.mergeFields = selectFields;
  template.dataTypes = dataTypes;
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 3');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSelectedHeaders() {
  
//***************************************************************************//
//getSelectHeaders()                                                         //
//retrieve frm the scriptProperties the value of mappingString created in    //
//step 2. The value of mappingString is a key-value pair with column number  //
//as the value. So we retrieve the value of headersPreRefine from script-    //
//Properties to get all the column header and match the value of mapping-    //
//String with headersPreRefine and stored the result in an array selected-   //
//Header. Return the selectedHeaders array back to getMergeConditions.html   //
//for further processing.                                                    //
//***************************************************************************//

  var scriptProperties = PropertiesService.getScriptProperties();
  
  var mappingString = scriptProperties.getProperty("mappingString");
  
  var mappingObject = JSON.parse(mappingString);
  var selectedHeaderIndex = new Array();
  
  var i = 0;
  
  for (var i = 0; i < mappingObject.length; ++i) {
  
    selectedHeaderIndex[i] = mappingObject[i].data;
  }
  
  var selectedHeader = new Array();
  var headersPreRefine = JSON.parse(scriptProperties.getProperty("headersPreRefine"));
  
  for (var i = 0; i < selectedHeaderIndex.length; ++i) {
  
    selectedHeader[i] = headersPreRefine[selectedHeaderIndex[i]];
  }
  
  return selectedHeader;
}

function getMergeConditions(mappingString) {

//***************************************************************************//
//getMergeConditions(mappingString)                                          //
//save the mappingString uploaded from getMappingString.html form client side//
//create and open getMergeConditions.html for user to define the conditions  //
//to run the task.                                                           //
//***************************************************************************//

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("mappingString", mappingString);
  
  var template = HtmlService.createTemplateFromFile("getMergeCondition.html");
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 4');

  SpreadsheetApp.getUi().showSidebar(html);
}

function getDestinationFolder(conditionString) {

//****************************************************************************//
// getDestinationFolder(conditionString)                                      //
// set conditionString variable with the value get from getMergeCondition     //
// .html from client side. Create and open getDestinationFolder.html for user // 
// to select a folder for placing the merged document.                        //
//****************************************************************************//

  if (conditionString != "NULL") {
  
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("mergeCondition", conditionString);
  }
  
  var template = HtmlService.createTemplateFromFile("getDestinationFolder.html");
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 5');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOutputFileType (destinationFolderId) {

//***************************************************************************//
//getOutputFileType(destinationFolderId)                                     //
//set destinationFolderId variable with the value get from                   //
//getDestinationFolder.html from client side,create and open                 //
//getOutputFileType.html for user to select the merged file format.          //
//***************************************************************************//

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("destinationFolderId", destinationFolderId);
  
  var template = HtmlService.createTemplateFromFile("getOutputFileType.html");
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 6');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

function getMergedFileName (fileType) {

//***************************************************************************//
//getMergedFileName(fileType)                                                //
//get the format of the merged file type input from getOutputFileType.html.  //
//from client side. If the file type is PDF simply set the value to          //
//scriptProperties, but if the value is anything other than PDF, get the     //
//file type of the template file from scriptProperties, and assign that value// 
//to mergedFileType. Create and open getOutputFileName.html for user to      //
//assign a file name for the merged file.                                    //
//***************************************************************************//

  var colNeedData = new Array();
  var scriptProperties = PropertiesService.getScriptProperties();
  var data = JSON.parse(scriptProperties.getProperty("mergeCondition"));
  
  if (fileType != "PDF") {
  
    var fileId = scriptProperties.getProperty("fileId");
    var fileTypeString = DriveApp.getFileById(fileId).getMimeType();
    scriptProperties.setProperty("mergedFileType", fileTypeString);
  }
  else {
  
    scriptProperties.setProperty("mergedFileType", fileType);
  }
  
  var nameToSelect = new Array();
  
  for (var i = 0; i < data.length; ++i) {
  
    nameToSelect.push(data[i].headerMap);
  }
  
  var template = HtmlService.createTemplateFromFile("getOutputFileName.html");
  
  template.nameToSelect = nameToSelect;
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('AutoCrat Setup - Step 7');
    
  SpreadsheetApp.getUi().showSidebar(html);
}

function setAllParameters(fileNameTypeString) {

//***************************************************************************//
//setAllParameters(fileNameTypeString)                                       //
//get the fileNameTypeString input from getOutputFileName.html. from client  //
//side and assign the value to mergedFileName. Save all the parameters       //
//obtained from step 1 to 5 in a worksheet as a task. Assign a task id for   //
//calling the autoCrat procedure. Check the value in the batchUpdate response//
//body to see if all the parameters are saved successfully and notify the    //
//user of the result in completeAutoCratSetup.html in client side.           //
//***************************************************************************//
  
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("mergedFileName", fileNameTypeString);
  
  var sheetName = scriptProperties.getProperty("sheetName");
  var destinationFoldId = scriptProperties.getProperty("destinationFolderId");
  var mergedFileName = scriptProperties.getProperty("mergedFileName");
  var mergeCondition = scriptProperties.getProperty("mergeCondition");
  var mergedFileType = scriptProperties.getProperty("mergedFileType");
  var mappingString = scriptProperties.getProperty("mappingString");
  var fileId = scriptProperties.getProperty("fileId");
  var headerRow = scriptProperties.getProperty("headerRow");
  var dataRow = scriptProperties.getProperty("dataRow");
  
  var uniqueTaskId = false;
  var taskId = ''
  var paraSheetName = "autoCrat_Parameters"
  
  var arrUpdate = [];
  var arrHeader = [];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var paraSheetExist = ss.getSheetByName(paraSheetName);
  var paraSheetId = '';
  
  if (!paraSheetExist) {
  
    var sParaSheet = ss.insertSheet(paraSheetName)
    paraSheetId = sParaSheet.getSheetId();
    sParaSheet.hideSheet();
    var newRow = 1;
    
    taskId = randomStr(15);  
  }
  else {
  
    paraSheetId = paraSheetExist.getSheetId();
    var matchId = '';
    var sParaSheet = ss.getSheetByName(paraSheetName);
    var fullRange = sParaSheet.getDataRange().getValues();
    
    if (fullRange) {
    
      // fullRange.splice(0, 1); splice does not work on 2D array
      var newRow = sParaSheet.getLastRow() + 1;
      
      while (!uniqueTaskId) {
    
        taskId = randomStr(15);
        
        matchId = fullRange.filter(function (row) {
        
          return row[0] === taskId;
        });
        
        if ( matchId.length == 0 ) {
        
          uniqueTaskId = true;
        }
      }
    }
    else {
    
      var newRow = 1;
    }
  }
  
  arrHeader.push("Task ID");
  arrHeader.push("File ID");
  arrHeader.push("Sheet Name");
  arrHeader.push("Header Row");
  arrHeader.push("Data Row");
  arrHeader.push("Destination Folder ID");
  arrHeader.push("Mapping String");
  arrHeader.push("Merge Conditions");
  arrHeader.push("Merged File Type");
  arrHeader.push("Merged File Name");
  
  arrUpdate.push(taskId);
  arrUpdate.push(fileId);
  arrUpdate.push(sheetName);
  arrUpdate.push(headerRow);
  arrUpdate.push(dataRow);
  arrUpdate.push(destinationFoldId);
  arrUpdate.push(mappingString);
  arrUpdate.push(mergeCondition);
  arrUpdate.push(mergedFileType);
  arrUpdate.push(mergedFileName);
  
  if (paraSheetExist) {
  
    var rngList = [];
    
    var names = {};
   
    names['range'] = paraSheetName + "!A" + newRow + ":J" + newRow;
    names['values'] = [arrUpdate];
    
    rngList.push(names);

    var resource = {
    
      valueInputOption: "USER_ENTERED",
      data: rngList
    };
    
    var response = Sheets.Spreadsheets.Values.batchUpdate(resource, spreadSheetId);
  }
  else {
  
    var rngList = [];
    
    var names = {};
    
    var arrResult = [];
    
    arrResult.push(arrHeader);
    arrResult.push(arrUpdate);
    
    names['range'] = paraSheetName + "!A" + newRow + ":J" + newRow + 1;
    names['values'] = arrResult;
    
    rngList.push(names);
    
    var resource = {
    
      valueInputOption: "USER_ENTERED",
      data: rngList
    };
    
    var response = Sheets.Spreadsheets.Values.batchUpdate(resource, spreadSheetId);
  }
  
  if (response.totalUpdatedSheets >= 1) {
    
    var responseString = "AutoCrat task parameters has been successfully set.";
    // delete all the global variables saved in property service
    scriptProperties.deleteAllProperties();
  }
  else {
    
    var responseString = "Unexpected error occured, task parameters cannot be saved.";
  }
  
  var template = HtmlService.createTemplateFromFile("completeAutoCratSetup");
  
  template.responseString = responseString;
  
  var html = template
  .evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle('AutoCrat Setup - Completed');
  
  SpreadsheetApp.getUi().showSidebar(html); 
}

function randomStr(m) {

    //Generate random string
    var m = m || 15; s = '', r = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (var i=0; i < m; i++) { 
      s += r.charAt(Math.floor(Math.random()*r.length)); 
    }
    return s;
};

function mergeDoc(taskArray) {
  
//***************************************************************************//
// mergeDoc(taskArray)                                                       //
  // accept the taskArray return from function searchPlan
//***************************************************************************//
  
  var successMerge = 0;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scriptProperties = PropertiesService.getScriptProperties();
  var conditionParam = JSON.parse(scriptProperties.getProperty("conditionParam"));
  var dataArray = JSON.parse(scriptProperties.getProperty("dataArray"));
  var formatArray = JSON.parse(scriptProperties.getProperty("formatArray"));
  
  var responseMsg = new Array();
  
  var destFolderId = DriveApp.getRootFolder().getId();
  var responseWriteRange = new Array();
  var docToTrash = new Array();
  var docToMerge = 0;
  
  for (var i=0; i<=taskArray.length-1; ++i) {
  
    var matchTask = conditionParam.filter(function (row) {
        
      return row[0] === taskArray[i][1];
    });
    
    if (matchTask.length > 0) {
    
      var fileId = matchTask[0][1];
      var sheetName = matchTask[0][2];
      var headerRowIndex = matchTask[0][3]-1;
      var dataRowIndex = matchTask[0][4]-1;
      var secondaryFolderId = matchTask[0][5];
      var mappingString = matchTask[0][6];
      var mergedFileType = matchTask[0][8];
      var mergedFileName = matchTask[0][9];
      
      var headers = dataArray[headerRowIndex];
      
      // get the merged file name
      var mergedFileNameObject = JSON.parse(mergedFileName);
      
      if (mergedFileNameObject[0].type == "preDefined") {
              
        var fileName = dataArray[taskArray[i][0]][headers.indexOf(mergedFileNameObject[0].value)];
      } else {
          // if the type property of mergedFileNameObject is not "preDefined", that means the user do not want to use any of the merge fields for the file name,
          // and use the string one specify as the file name. To avoid duplicated file names (even it is acceptable in google drive), we will use user sepecify
          // string followed by a string which equivalent to the time the file is generated as a suffix for the file name.
                
          var fileName = mergedFileNameObject[0].value + "-" + genTimeString();
      }
              
      var copyId = DriveApp.getFileById(fileId).makeCopy(fileName).getId();
              
      var dataRow = dataArray[taskArray[i][0]];
      var formatRow = formatArray[taskArray[i][0]];
              
      // replace tag in the document and save it to a selected folder
      copyId = autoCrat_makeMergeDoc(copyId, dataRow, formatRow, destFolderId, secondaryFolderId, JSON.parse(mappingString));
             
      if (mergedFileType = "PDF") {
              
        var response = JSON.parse(autoCrat_converToPdf(copyId, destFolderId, secondaryFolderId));
                
        if (response[0].fileId) {
                  
          var responseWriteBack = new Array();
         
          responseWriteBack.push(taskArray[i][0]+1);
          responseWriteBack.push(response[0].fileId);
          responseWriteBack.push('=HYPERLINK("' + response[0].url + '", "' + response[0].fileName + '")');
          responseWriteBack.push(response[0].url);
          responseWriteBack.push(response[0].message);
                  
          responseWriteRange.push(responseWriteBack);
          
          docToMerge += 1;
        }
        
        docToTrash.push(copyId);
      }
    }
  }
          
  if (responseWriteRange.length > 0) {
            
    autoCrat_updateLog(responseWriteRange, sheetName, ss.getId());
  }
          
  if (docToMerge) {
    
    responseMsg.push(1, "成功製作了 " + docToMerge + " 份文件。", docToTrash);
    return responseMsg;
  } else {
    
    responseMsg.push(0, "未能製作任何文件。");
    return responseMsg;
  }   
}
    
function autoCratMain(taskId) {
  
//***************************************************************************//
// autoCratMain(taskId)                                                      //
//                                                                           //
// - ver.0 retrieve task settings by querying the autoCrat_Parameters sheet  //
//         with a unique taskId. Check if the 4 columns with autoCrat prefix //
//         are empty, if they are not empty, it means the corresponding row  //
//         has been processed by autoCrat. Next check if the merge condition //
//         retrieved from autoCrat_Parameters sheet can be met. Clone the    //
//         template file with filename retrieved from autCrat_Parameters.    //
//         Replace the value in the template file according to the mapping   //
//         string retrieved from autoCrat_Parameters. Convert the file to    //
//         the file type specified in the autoCrat_Parameters. Batch write   //
//         the file id, file name, URL & message to the sheet.               //
//                                                                           //
//         This is another approach to merge file which require programmer   //
//         to hardcode the task Id to the function to generate merged        //
//         document. Then this function will check the merge condition       //
//         against the merge data to see if the all the conditions are       //
//         satisfied prior to run the job. This approach is workable but     //
//         not intuitive in terms of operation as we cannot expect one to    //
//         define the merge condition and ask a programmer to hardcode the   //
//         taskId.                                                           //
//***************************************************************************//
  
  //taskId ="bcTFqEIGzQikSk8";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sParaSheet = ss.getSheetByName("autoCrat_Parameters");
  
  if (sParaSheet) {
    
    var fullRange = sParaSheet.getDataRange().getValues();
    
    if (fullRange) {
        
        var matchTask = fullRange.filter(function (row) {
        
          return row[0] === taskId;
        });
        
        if ( matchTask.length > 0 ) {
        
          var fileId = matchTask[0][1];
          var sheetName = matchTask[0][2];
          var secondaryFolderId = matchTask[0][3];
          var mappingString = matchTask[0][4];
          var mergeCondition = matchTask[0][5];
          var mergedFileType = matchTask[0][6];
          var mergedFileName = matchTask[0][7];
          
          var mergeConditionObject = JSON.parse(mergeCondition);
          
          var dataSheet = ss.getSheetByName(sheetName);
          var dataRange = dataSheet.getDataRange().getValues();
          var formatRange = dataSheet.getDataRange().getNumberFormats();
          var spreadsheetId = ss.getId();
          
          var headers = dataRange.splice(0, 1)[0];
                    
          var statusCol = headers.indexOf("AutoCrat_Merged Doc ID");
          var linkCol = headers.indexOf("AutoCrat_Merged Doc URL");
          var urlCol = headers.indexOf("AutoCrat_Link to merged Doc");
          var docIdCol = headers.indexOf("AutoCrat_Document Merge Status");
          
          
          // check if the merge condition is met, if merge condition is set
          if (mergeConditionObject.length > 0) {
          
            var checkConditionResponse = JSON.parse(checkRunCondition(headers, dataRange, mergeConditionObject));
            
            if (checkConditionResponse[0].ruleViolate == true) {
            
              Browser.msgBox(checkConditionResponse[0].response);
              return;
            }
          }
          
          // get the merged file name
          var docToMerge = 0;
          var mergedFileNameObject = JSON.parse(mergedFileName);
          var destFolderId = DriveApp.getRootFolder().getId();
          var responseWriteRange = new Array();
          var docToTrash = new Array();
          
          for (var i = 0; i < dataRange.length; ++i) {
          
            if (dataRange[i][statusCol] == '' && dataRange[i][linkCol] == '' && dataRange[i][urlCol] == '' && dataRange[i][docIdCol] == '') {
            
              docToMerge += 1;
              
              if (mergedFileNameObject[0].type == "preDefined") {
              
                var fileName = dataRange[i][headers.indexOf(mergedFileNameObject[0].value)];
              }
              else {
                // if the type property of mergedFileNameObject is not "preDefined", that means the user do not want to use any of the merge fields for the file name,
                // and use the string one specify as the file name. To avoid duplicated file names (even it is acceptable in google drive), we will use user sepecify
                // string followed by a string which equivalent to the time the file is generated as a suffix for the file name.
                
                var fileName = mergedFileNameObject[0].value + "-" + genTimeString();
              }
              
              var copyId = DriveApp.getFileById(fileId).makeCopy(fileName).getId();
              
              var dataRow = dataRange[i];
              var formatRow = formatRange[i];
              
              // replace tag in the document and save it to a selected folder
              copyId = autoCrat_makeMergeDoc(copyId, dataRow, formatRange, destFolderId, secondaryFolderId, JSON.parse(mappingString));
             
              if (mergedFileType = "PDF") {
              
                var response = JSON.parse(autoCrat_converToPdf(copyId, destFolderId, secondaryFolderId));
                
                if (response[0].fileId) {
                  
                    var responseWriteBack = new Array();
                    responseWriteBack.push(i+2);
                    responseWriteBack.push(response[0].fileId);
                    responseWriteBack.push('=HYPERLINK("' + response[0].url + '", "' + response[0].fileName + '")');
                    responseWriteBack.push(response[0].url);
                    responseWriteBack.push(response[0].message);
                  
                  responseWriteRange.push(responseWriteBack);
                }
                
                //autoCrat_trashDoc(copyId);
                
                docToTrash.push(copyId);
              }
            }
          }
          
          if (responseWriteRange.length > 0) {
            
            autoCrat_updateLog(responseWriteRange, sheetName, ss.getId());
          }
          
          if (docToMerge) {
          
            Browser.msgBox("成功製件了 " + docToMerge + " 份文件。");
            autoCrat_trashDoc(docToTrash);
          } else {
          
            Browser.msgBox("未能製件任何文件");
          }
        }
      
    }
  }
}

function checkRunCondition(headers, dataRange, mergeConditionObject) {

//***************************************************************************//
// checkRunCondition(headers, dataRange, mergeConditionObject)               //
//                                                                           //
// - ver.0 called by autoCratMain retrieve task settings by querying the     //
//         autoCrat_Parameters sheet with a unique taskId. Check if the 4    //
//         columns with autoCrat prefix are empty, if they are not empty, it //
//         means the corresponding row has been processed by autoCrat. Next  //
//         check if the merge condition retrieved from autoCrat_Parameters   //
//         sheet can be met. Clone the template file with filename retrieved //
//         from autCrat_Parameters. Replace the value in the template file   //
//         according to the mapping string retrieved from autoCrat_          //
//         Parameters. Convert the file to the file type specified in the    //
//         autoCrat_Parameters. Batch write the file id, file name, URL &    // 
//         message to the sheet.                                             //  
//***************************************************************************//
  
  var conditionViolate = false;
  var errorMessage = '';

  for (var i = 0; i < dataRange.length; ++i) {
  
    for (var j = 0; j < mergeConditionObject.length; ++j) {
    
      if (mergeConditionObject[j].value == "NOT_NULL") {
          
        if (dataRange[i][headers.indexOf(mergeConditionObject[j].headerMap)] == '') {
      
          conditionViolate = true;
          errorMessage = "Field " + mergeConditionObject[j].headerMap + " cannot be blank.";
          break;
        }    
      }
      else {
    
        if (dataRange[headers.indexOf(mergeConditionObject[j].headerMap)] != mergeConditionObject[j].value) {
            
          conditionViolate = true;
          errorMessage = "Field " + mergeConditionObject[j].headerMap + " must equal to " + mergeConditionObject[j].value + ".";
          break;
        }
      }
    }
    
    if (conditionViolate) {
        
      break;    
    }
  }
    
  var responseString = '[{"ruleViolate" : "' + conditionViolate + '", "response" : "' + errorMessage + '"}]'; 
  
  return responseString;
}

function createDocument() {

  var template = HtmlService.createTemplateFromFile("mergeDocConsole.html");
  
  var html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(600)
    .setHeight(200);
    
  SpreadsheetApp.getUi().showModalDialog(html, "製作文件");
}

function searchPlan() {

  var paraSheetName = "autoCrat_Parameters";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var curSheetName = ss.getActiveSheet().getName();
  var sParaSheet = ss.getSheetByName(paraSheetName);
  var mergeDataRange = ss.getSheetByName(curSheetName).getDataRange();
  var dataToMerge = mergeDataRange.getValues();
  var formatToMerge = mergeDataRange.getNumberFormats();
  
  var responseMsg = new Array();
      
  if (sParaSheet) {
  
    var paraSheetData = sParaSheet.getDataRange().getValues();
    var sheetCol = paraSheetData[0].indexOf("Sheet Name");
    var conditionCol = paraSheetData[0].indexOf("Merge Conditions");
    var headerCol = paraSheetData[0].indexOf("Header Row");
    var dataCol = paraSheetData[0].indexOf("Data Row");
    var doneMark = dataToMerge[0].indexOf("AutoCrat_Merged Doc ID");
    
    var sheetCheck = new Array();
    var conditionCheck = new Array();
    var rowLocate = new Array();
    
    for (var i=1; i<=paraSheetData.length-1; ++i) {
    
      var getParam = JSON.parse(paraSheetData[i][conditionCol]);
      
      sheetCheck.push([paraSheetData[i][sheetCol]]);
      rowLocate.push([paraSheetData[i][headerCol], paraSheetData[i][dataCol]]);
      
      var conditionCheckInner = new Array();
      
      for (var j=0; j<=getParam.length-1; ++j) {
        
        conditionCheckInner.push(getParam[j]);
      }
      
      conditionCheck.push(conditionCheckInner);
    }
    
    var scriptProperties = PropertiesService.getScriptProperties();
    
    scriptProperties.setProperty("conditionParam", JSON.stringify(paraSheetData));
    scriptProperties.setProperty("dataArray", JSON.stringify(dataToMerge));
    scriptProperties.setProperty("formatArray", JSON.stringify(formatToMerge));
    
    var exeArray = new Array();
    
    for (var i=0; i<=sheetCheck.length-1; ++i) {
    
      if (curSheetName == sheetCheck[i]) {
      
        var conditionHeaderRow = rowLocate[i][0];
        var conditionDataRow = rowLocate[i][1];
        
        for (var j=conditionDataRow-1; j<=dataToMerge.length-1; ++j) {
        
          var matchCondition = true;
          
          if (dataToMerge[j][doneMark] == "") {
          
            for (var m=0; m<=conditionCheck.length-1; ++m) {
            
              var colToCheck = dataToMerge[conditionHeaderRow-1].indexOf(conditionCheck[i][m].headerMap);
              
              if (colToCheck >= 0) {
              
                if (conditionCheck[i][m].value == "NOT_NULL") {
                
                  if (dataToMerge[j][colToCheck] == "") {
                    
                    matchCondition = false;
                  }
                } else {
                
                  if (dataToMerge[j][colToCheck] != conditionCheck[i][m].value) {
                  
                    matchCondition = false;
                  }
                }
              } else {
              
                matchCondition = false;
              }
            }
            
            m += 1;
          } else {
          
            matchCondition = false;
          }
          
          if (matchCondition == true) {
          
            exeArray.push([j, paraSheetData[i+1][0]]);
          }
        }
      }
    }
    if (exeArray.length > 0) {
    
      //var responseMsg = mergeDoc(exeArray);
      
      responseMsg = mergeDoc(exeArray);
      return responseMsg;
    } else {
    
      // no job can be merged
      
      responseMsg.push(0, "沒有資料可以製作文件。");
      return responseMsg;
    }
  } else {
  
    // no parasheet
    
    responseMsg.push(0, "必須先設定工作參數才可以製作文件。");
    return responseMsg;
  }
}





