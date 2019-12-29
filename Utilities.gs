function autoCrat_fetchSheetHeaders(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var cols = sheet.getLastColumn();
  var range = sheet.getRange(1,1,1,cols);
  var data = range.getValues();
  var headers = new Array();
  for (i = 0; i<data[0].length; i++) {
  
     headers[i] = data[0][i];
  }
  return headers;
}

function autoCrat_normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = autoCrat_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

function autoCrat_normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!autoCrat_isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && autoCrat_isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function autoCrat_prepareMergeFields(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = autoCrat_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

function autoCrat_prepareMergeField(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    
    header[i].replaceText("<>", "");
  }
  return header[i];
}

function autoCrat_replaceStringFields(string, rowValues, rowFormats, headers, mergeTags) {
  var newString = string;
  var timeZone = Session.getTimeZone();
  for (var i=0; i<headers.length; i++) {
    var thisHeader = headers[i];
    var colNum = autoCrat_getColumnNumberFromHeader(thisHeader, headers);
    if (((rowFormats[colNum-1]=="M/d/yyyy")||(rowFormats[colNum-1]=="MMMM d, yyyy")||(rowFormats[colNum-1]=="M/d/yyyy H:mm:ss"))&&(rowValues[colNum-1]!="")) {
      try {
        var replacementValue = Utilities.formatDate(rowValues[colNum-1], timeZone, rowFormats[colNum-1]);
      }
      catch(err) {
        var date = new Date(rowValues[colNum-1]);
        var colVal = Utilities.formatDate(date, timeZone, rowFormats[colNum-1]);
      }
    } else {
      var replacementValue = rowValues[colNum-1];
    }
    var replaceTag = mergeTags[i];
    replaceTag = replaceTag.replace("$","\\$") + "\\b";
    var find = new RegExp(replaceTag, "g");
    newString = newString.replace(find,replacementValue);
  }
  var currentTime = new Date()
  var month = currentTime.getMonth() + 1;
  var day = currentTime.getDate();
  var year = currentTime.getFullYear();
  newString = newString.replace("$currDate", month+"/"+day+"/"+year);
  return newString;
}

function autoCrat_getColumnNumberFromHeader(header, headers) {
  var colFlag = headers.indexOf(header) + 1;
  return colFlag;
}

//Grabs any <<Merge Tags>> from a document 
function autoCrat_fetchDocFields(fileId) {
  //var fileId = '1KiIBQYFmQR7QXTfJ0iGl3e5QogUGq8kDOgi9zioHuyE';
  //var fileType = DocsList.getFileById(fileId).getFileType().toString();
  var fileType = DriveApp.getFileById(fileId).getMimeType();
  if (fileType == "application/vnd.google-apps.document") {
  //if ((fileType=="DOCUMENT")||(fileType=="document")) { 
    var template = DocumentApp.openById(fileId);
    var title = template.getName();
    var fieldExp = "[<]{2,}\\S[^,]*?[>]{2,}";
    var result;
    var matchResults = new Array();
    var headerFieldNames = new Array();
    var bodyFieldNames = new Array();
    var footerFieldNames = new Array();
    
    //get all tags in doc header
    var header = template.getHeader();
    if (header!=null) { matchResults[0] = header.findText(fieldExp);}
    if (matchResults[0]!=null){
      var element = matchResults[0].getElement().asText().getText();
      var start = matchResults[0].getStartOffset()
      var end = matchResults[0].getEndOffsetInclusive()+1;
      var length = end-start;
      headerFieldNames[0] = element.substr(start,length)
      var i = 0;
      while (headerFieldNames[i]) {
        matchResults[i+1] = template.getHeader().findText(fieldExp, matchResults[i]);
        if (matchResults[i+1]) {
          var element = matchResults[i+1].getElement().asText().getText();
          var start = matchResults[i+1].getStartOffset()
          var end = matchResults[i+1].getEndOffsetInclusive()+1;
          var length = end-start;
          headerFieldNames[i+1] = element.substr(start,length);
        }
        i++;
      }
    }
    
    //get all tags in doc body
    matchResults = [];
    var body = template.getActiveSection();
    if (body!=null) { matchResults[0] = body.findText(fieldExp);}
    if (matchResults[0]!=null){
      var element = matchResults[0].getElement().asText().getText();
      var start = matchResults[0].getStartOffset()
      var end = matchResults[0].getEndOffsetInclusive()+1;
      var length = end-start;
      bodyFieldNames[0] = element.substr(start,length)
      var i = 0;
      while (bodyFieldNames[i]) {
        matchResults[i+1] = template.getActiveSection().findText(fieldExp, matchResults[i]);
        if (matchResults[i+1]) {
          var element = matchResults[i+1].getElement().asText().getText();
          var start = matchResults[i+1].getStartOffset()
          var end = matchResults[i+1].getEndOffsetInclusive()+1;
          var length = end-start;
          bodyFieldNames[i+1] = element.substr(start,length);
        }
        i++;
      }
    }
    
    //get all tags in doc footer
    var matchResults = [];
    var footer = template.getFooter();
    if (footer!=null) { matchResults[0] = footer.findText(fieldExp);}
    if (matchResults[0]!=null){
      var element = matchResults[0].getElement().asText().getText();
      var start = matchResults[0].getStartOffset()
      var end = matchResults[0].getEndOffsetInclusive()+1;
      var length = end-start;
      footerFieldNames[0] = element.substr(start,length)
      var i = 0;
      while (footerFieldNames[i]) {
        matchResults[i+1] = template.getFooter().findText(fieldExp, matchResults[i]);
        if (matchResults[i+1]) {
          var element = matchResults[i+1].getElement().asText().getText();
          var start = matchResults[i+1].getStartOffset()
          var end = matchResults[i+1].getEndOffsetInclusive()+1;
          var length = end-start;
          footerFieldNames[i+1] = element.substr(start,length);
        }
        i++;
      }
    }
    var fieldNames = headerFieldNames.concat(bodyFieldNames, footerFieldNames);
    fieldNames = autoCrat_removeDuplicateElement(fieldNames);
    return fieldNames; 
  }
  if (fileType=="application/vnd.google-apps.spreadsheet") {
    var ss = SpreadsheetApp.openById(fileId);
    var sheets = ss.getSheets();
    var allTags = [];
    for (var i=0; i<sheets.length; i++) {
      var range = sheets[i].getDataRange();
      var values = range.getValues();
      for (var j=0; j<values.length; j++) {
        for (var k=0; k<values[j].length; k++) {
          var cellValue = values[j][k].toString();
          var exp = new RegExp(/[<]{2,}\S[^,]*?[>]{2,}/g);
          var tags = cellValue.match(exp);
          if (tags) {
            for (var l=0; l<tags.length; l++) {
              allTags.push(tags[l]);
            }
          }
        }
      }
    }
    allTags = autoCrat_removeDuplicateElement(allTags);
    return allTags;
  }
}

//Takes out any duplicates from an array of values
function autoCrat_removeDuplicateElement(arrayName)
{
  var newArray=new Array();
  label:for(var i=0; i<arrayName.length;i++ )
  {  
    for(var j=0; j<newArray.length;j++ )
    {
      if(newArray[j]==arrayName[i]) 
        continue label;
    }
    newArray[newArray.length] = arrayName[i];
  }
  return newArray;
}

function autoCrat_makeMergeDoc(copyId, dataRange, formatRange, folderId, secondaryFolderIds, mappingObject) {
  // Get document template, copy it as a new temp doc, and save the Doc’s id
  // check the file type of the document template
  var fileType = DriveApp.getFileById(copyId).getMimeType();
  if (fileType == "application/vnd.google-apps.document") { 
    // Open the temporary document
    var copyDoc = DocumentApp.openById(copyId);
    // Get the document’s body section
    var copyHeader = copyDoc.getHeader();
    var copyBody = copyDoc.getActiveSection();
    var copyFooter = copyDoc.getFooter();
    
    // replace field in document template with values of mappingObject
    
    for ( var i = 0; i< mappingObject.length; ++i) {
      
      var timeZone = Session.getScriptTimeZone();
      
      if ((formatRange[mappingObject[i].data]=="M/d/yyyy")||(formatRange[mappingObject[i].data]=="MMMM d, yyyy")||(formatRange[mappingObject[i].data]=="M/d/yyyy H:mm:ss")) {
        try {
          var colVal = Utilities.formatDate(dataRange[mappingObject[i].data], timeZone, formatRange[mappingObject[i].data]);
        }
        catch(err) {
          var date = new Date(dataRange[mappingObject[i].data]);
          var colVal = Utilities.formatDate(date, timeZone, formatRange[mappingObject[i].data]);
        }
      } else {
        var colVal = dataRange[mappingObject[i].data];
      }
    
      if (copyHeader) {
      
        copyHeader.replaceText(mappingObject[i].tag, colVal);
      }
      
      if (copyBody) {
      
        copyBody.replaceText(mappingObject[i].tag, colVal);
      }
      
      if (copyFooter) {
      
        copyFooter.replaceText(mappingObject[i].tag, colVal);
      }
    }
    
    copyDoc.saveAndClose();
  }
  
  if (fileType == "application/vnd.google-apps.spreadsheet") {
    var exp = new RegExp(/[<]{2,}\S[^,]*?[>]{2,}/g);
    var ss = SpreadsheetApp.openById(copyId);
    var sheets = ss.getSheets();
    for (var i=0; i<sheets.length; i++) {
      
      var range = sheets[i].getDataRange();
      var values = range.getValues();
      var formulas = range.getFormulas();
      var formats = range.getNumberFormats();
      
      for (var j=0; j<values.length; j++) {
        
        for (var k=0; k<values[j].length; k++) {
          
          var cellValue = values[j][k].toString();
          var cellFormula = formulas[j][k].toString();
          var tags = cellValue.match(exp);
          
          if (tags) {
            
            for (var n=0; n<tags.length; n++) {
              
              var normalizedFieldName = autoCrat_normalizeHeader(tags[n]);
              var colNum = mappingObject[normalizedFieldName];
              var timeZone = Session.getScriptTimeZone();
              
              if ((formatRange[mappingObject[i].data]=="M/d/yyyy")||(formatRange[mappingObject[i].data]=="MMMM d, yyyy")||(formatRange[mappingObject[i].data]=="M/d/yyyy H:mm:ss")) {
                
                try {
                  
                  var colVal = Utilities.formatDate(dataRange[mappingObject[i].data], timeZone, formatRange[mappingObject[i].data]);
                  formats[j][k] = formatRange[mappingObject[i].data];
                } catch(err) {
                  
                  var date = new Date(dataRange[mappingObject[i].data]);
                  var colVal = Utilities.formatDate(date, timeZone, formatRange[mappingObject[i].data]);
                  formats[j][k] = formatRange[mappingObject[i].data];
                }
              } else {
                
                var colVal = dataRange[mappingObject[i].data];
              }
              
              if ((cellFormula) && (cellFormula != '')) {
                
                cellFormula = cellFormula.replace(tags[n], colVal);
                formulas[j][k] = cellFormula;
              } else {
                
                cellValue = cellValue.replace(tags[n], colVal);
                values[j][k] = cellValue;
              }
            }
          }
        }
      }
      range.setValues(values);
      for (var j=0; j<formulas.length; j++) {
        
        for (var k=0; k<formulas[0].length; k++) {
          
          if (formulas[j][k]!='') {
            
            sheets[i].getRange(j+1, k+1).setFormula(formulas[j][k]);
          }
        }
      }
      sheets[i].activate();
    }
  }
  
  // move to folder
  var folder = DriveApp.getFolderById(folderId);
  var file = DriveApp.getFileById(copyId);
  
  if (secondaryFolderIds) {
    
    var secondaryFolder = DriveApp.getFolderById(secondaryFolderIds);
    secondaryFolder.addFile(file);
  }
  
  var rootFolder = DriveApp.getFolderById(folderId);
  rootFolder.removeFile(file);
  
  return copyId;
}

//Creates PDF in a designated folder and returns Id
function autoCrat_converToPdf (copyId, folderId, secondaryFolderIds) {
  var folder = DriveApp.getFolderById(folderId);
  var pdfBlob = DriveApp.getFileById(copyId).getAs("application/pdf"); 
  var pdfFile = DriveApp.createFile(pdfBlob);
  pdfFile.setName(DriveApp.getFileById(copyId).getName() + ".pdf");
  folder.addFile(pdfFile);
  
  if (secondaryFolderIds) {
  
    var secondaryFolder = DriveApp.getFolderById(secondaryFolderIds);
    secondaryFolder.addFile(pdfFile);
  }
  
  var rootFolder = DriveApp.getRootFolder();
  rootFolder.removeFile(pdfFile);
  var pdfId = pdfFile.getId();
  var pdfName = pdfFile.getName();
  var pdfUrl = pdfFile.getUrl();
  var message = "PDF File " + pdfFile.getName() + " created.";
  
  var response = '[{"fileId" : "' + pdfId + '", "fileName" : "' + pdfName + '", "url" : "' + pdfUrl + '", "message" : "' + message + '"}]';
  
  return response;
}

//Trashes given docIdFind and replace
function autoCrat_trashDoc (docToTrash) {
  
  for (var i=0; i<docToTrash.length; ++i) {
  
    DriveApp.getFileById(docToTrash[i]).setTrashed(true);
  }
}

// batch write the id, filename, url and message of the generated file to sheet
function autoCrat_updateLog(responseWriteRange, sheetName, spreadsheetId) {
  
  var rowGrouping = new Array();
  var updateData = [];
  var startIndex = 0;
  var endIndex = 0;
  var dataValue = [];

  for (var i=0; i<=responseWriteRange.length - 1; ++i) {
    
    if (i != (responseWriteRange.length -1)) {
  
      if (responseWriteRange[i][0] + 1 == responseWriteRange[i+1][0]) {
   
        endIndex = i+1;
      } else {
    
              if (endIndex == 0) {
          
                endIndex = startIndex
              }
               rowGrouping.push([startIndex, endIndex]);
               endIndex = 0;
               startIndex = i+1;
      }
    } else { 
    
      rowGrouping.push([startIndex, responseWriteRange.length-1]);
    }
  }
 
  var writeData = new Array();
  
  for (var i=0; i<=rowGrouping.length - 1; ++i) {
    
    var dataGrouping = new Array();
  
    for (var j=rowGrouping[i][0]; j<=rowGrouping[i][1]; ++j) {
    
      dataGrouping.push([responseWriteRange[j][1], responseWriteRange[j][2], responseWriteRange[j][3], responseWriteRange[j][4]]);
    }
    
    var rangeString = sheetName + '!G' + responseWriteRange[rowGrouping[i][0]][0] + ':J' + responseWriteRange[rowGrouping[i][1]][0];
    
    var dataObject = {
       'range' : rangeString,
       'values' : dataGrouping
    };
    
    writeData.push(dataObject);
  }
  
  var request = {
  
    'valueInputOption' : 'USER_ENTERED',
    'data' : writeData
  };
  
  var response = Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheetId);
}

// Returns true if the character char is alphabetical, false otherwise.
function autoCrat_isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
      autoCrat_isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function autoCrat_isDigit(char) {
  return char >= '0' && char <= '9';
  
}

// Return time string for the time this function is run

function genTimeString() {

  var timeString = "";
  
  timeString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  
  return timeString;
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}