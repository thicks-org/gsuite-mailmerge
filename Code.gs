//Simple gsuite doc and spreadsheet mail merge
//thicks@thicks.org

function getMergeData()
{
  var ui = DocumentApp.getUi();
  var sourceUrl = ui.prompt("Enter source sheet document URL").getResponseText();
  
  var mailData = SpreadsheetApp.openByUrl(sourceUrl).getActiveSheet();
  var ranges = mailData.getNamedRanges();
  var dataRange = null
  for (var i = 0; i < ranges.length; i++) {
    if (ranges[i].getName()==='merge_data') {
      dataRange = ranges[i].getRange();
      break;
    }
  }
  if (dataRange===null) {
    ui.alert("No range named data in source sheet");
    return null;
  }
  
  return dataRange.getDisplayValues();
}

function getSourceDoc()
{
  var sourceDoc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi();
 
  return sourceDoc;
}

function getMergedDoc(sourceDoc)
{
  var defaultName = sourceDoc.getName() + "-merged";
  var docName = DocumentApp.getUi().prompt("Name for merged document [default: " + defaultName + "]").getResponseText();
  if (docName==="") {
    docName = defaultName;
  }
  return DocumentApp.create(docName);
}

function mailMergeInstance(text,variableNames,variableValues)
{
  //Text in source document is {{value}}, where the variable name is simply value
  if(variableValues.length) {
    for(var i=0; i<variableNames.length; i++) {
      var replaceRegex = new RegExp("\{\{" + variableNames[i] + "\}\}",'g');
      text = text.replace(replaceRegex, variableValues.shift());
    }
    return text;
  } else {
    return null;
  }
}

function mailMergeDoc() {
  var ui = DocumentApp.getUi();

  var sourceDoc = getSourceDoc();
  var sourceText = sourceDoc.getBody().getText();
    
  var mergeDataValues = getMergeData();
  
  //Create a new document based upon current document name
  var mergedDoc = getMergedDoc(sourceDoc);
  
  //The first row from mergedDataValues contains the variable names to be replaced in the source text
  //Each subsequent row is then the values corresponding to each variable name
  var variableNames = mergeDataValues.shift();
  while(mergeDataValues.length) {
    var mergedText = mailMergeInstance(sourceDoc.getBody().getText(),
                                       variableNames,
                                       mergeDataValues.shift());
    if (mergedText===null) {
      //Something went wrong, delete the merged doc
      break;
    }
    mergedDoc.getBody().appendParagraph(mergedText);
    mergedDoc.getBody().appendPageBreak();
  }
}


function myFunction() {
 DocumentApp.getUi()
  .createMenu("Mail Merge")
  .addItem("Start Mail Merge",'mailMergeDoc')
  .addToUi();
}

function test() {
  mailMergeDoc();
}
