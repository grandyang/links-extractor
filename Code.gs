/****************************************************************************
*
* Copyright (C) 2019 - Current, Xiaochen Yang, All rights reserved.
* Contact: grandyxc@gmail.com
*
*
* "Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are
* met:
*   * Redistributions of source code must retain the above copyright
*     notice, this list of conditions and the following disclaimer.
*   * Redistributions in binary form must reproduce the above copyright
*     notice, this list of conditions and the following disclaimer in
*     the documentation and/or other materials provided with the
*     distribution.
*   * The names of its contributors may not be used to endorse or promote 
*      products derived from this software without specific prior written 
*     permission.
*
*
* THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
* "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
* LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
* A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
* OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
* SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
* LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
* DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
* THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
* OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."
*
****************************************************************************/


/**
 * Runs when a user installs an add-on.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when a user opens a spreadsheet, document, presentation, 
 * or form that the user has permission to edit.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Links Extractor')
  .addItem('Extract', 'extractLinks')
  .addToUi();
}

/** 
 * Extract the texts and links from the user input range, 
 * and save to next available two columns. 
 * 
 */
function extractLinks() {
  var ui = SpreadsheetApp.getUi();
  
  // Get user input range that have links needs to be extracted.
  var inputRangeLabel = getInputRangeLabel();
  if (inputRangeLabel == "") { // If empty, return.
    return;
  }
  
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = activeSpreadSheet.getActiveSheet();
  var activeRange = activeSheet.getRange(inputRangeLabel);
  
  // Get all formulas from the range.
  var allFormulas = activeRange.getFormulas();
  // Get all text values from the range.
  var allTexts = activeRange.getValues();
  // Output arrays for texts and links.
  var outputTexts = [];
  var outputLinks = [];
  var countLinks = 0;
  for (var i = 0; i < allFormulas.length; ++i) {
    var row = [];
    //ui.alert(allFormulas[0] + " " + allFormulas[0].length);
    for (var j = 0; j < allFormulas[0].length; ++j) {
      var url = allFormulas[i][j].match(/=hyperlink\("([^"]+)"/i);
      row.push(url ? url[1] : '');
    }
    if (row.length != 0 && row[0] !== '') {
      ++countLinks;
      //ui.alert(row);
    }
    //ui.alert(row.length);
    outputLinks.push(row);
    outputTexts.push(allTexts[i]);
  }
  //ui.alert(countLinks);
  
  // Find next available range and save texts.
  var outputRangeText = findOutputRange(activeRange);
  outputRangeText.setValues(outputTexts);
  
  // Find next available range and save links.  
  var outputRangeLink = findOutputRange(activeRange);
  outputRangeLink.setValues(outputLinks);
  
  // If fail to extract any links, show hints to user.
  if (countLinks === 0) {
    ui.alert("Fail to extract any links!\n\n",
             " Please make sure the data contains link. \n\n"
             + "The links must be wrapped in Hyperlink and located in Formula.\n\n"
             + "For example: \n"
             + "=HYPERLINK(\"https://www.google.com/\",\"Google\")\n\n"
             + "If the pasted rich text data contain links, please try the following steps:\n\n"
             + "1. Right click the cell which contains link, and hit \"Edit Link\"\n"
             + "2. Click \"Apply\"\n\n"
             + "Then you should see the HYPERLINK in the Formula.",
             ui.ButtonSet.OK);
  }
}

/** 
 * Pop up a dialog that allows user to input the range they want to extract the links.
 * 
 * @return (String)    return user input range label.
 */
function getInputRangeLabel() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Links Extractor',
    'Please enter the range that you want to extract links from (e.g. A1:A10):',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // May need to check valid range later.
    return text;
  } 
  return "";
}

/** 
 * Find next available output range, the height and width are equal to the input range.
 * Also, the output range row is same as the input range. 
 * Reference: 
 * https://developers.google.com/apps-script/reference/spreadsheet/sheet#getrangerow-column-numrows-numcolumns
 *
 * @param {Range}    inputRange    The input range.
 * @return {Range}    Return the output range.
 */
function findOutputRange(inputRange) { 
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = activeSpreadSheet.getActiveSheet();
  
  var inputHeight = inputRange.getHeight();
  var inputWidth = inputRange.getWidth();
  var outputRow = inputRange.getRow();
  var outputColumn = activeSheet.getDataRange().getWidth() + 1;

  return activeSheet.getRange(outputRow, outputColumn, inputHeight, inputWidth);
}


