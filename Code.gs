const PRESENTATION_ID = '1qLn8K7VHbxVx0oF46ETOGFUInneVMPPftmZ7GArQf_rbT5CBxcnCRIcF'

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('My Functions')
      .addSubMenu(ui.createMenu('Code Slides')
          .addItem('Update Code All Slides', 'codeColourAllSlides')
          .addItem('Update Code Selected Item', 'updateCodeTextColours'))
      .addToUi();
}

////////// UPDATE CODE TEXT COLOURS //////////
/* Contains: codeColourAllSlides(), updateCodeTextColours(), regexExtractWords(text, regex), colour(searchWord, replaceWord) */
function codeColourAllSlides() {
  /* Runs through entire presentation and changes colour of text withing shapes that contain 'Roboto Mono' text 

  Inputs:
      - n/a

  Outputs:
      - n/a */

  var ui = SlidesApp.getUi(); // Same variations.
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var theme = showPrompt();

  const fontSearch = "Roboto Mono";
  for (let i = 0; i < slides.length; i++) {
    var slideElements = slides[i].getPageElements()
    
    for (let j = 0; j < slideElements.length; j++) {
      if (slideElements[j].getPageElementType() == "SHAPE") {
        var shapeText = slideElements[j].asShape().getText();
        var textFont = shapeText.getTextStyle().getFontFamily();

        Logger.log("slide: " + i + ", element: " + j + ", " + textFont);
        if (textFont == fontSearch) {
          slideElements[j].select();
          updateCodeTextColours(theme);
        }
      }
    }
  }
}

function updateCodeTextColours(theme = 0) {
  /* Choose a page element that contains code, and this function will update all the colours,
  according to the CSinSC style guide.

  Inputs:
      - theme: the colour theme the user wants to use
          - if function called with no theme parameter, the fucntion will prompt user to choose one
          - can be called with or without theme parameter

  Outputs:
      - n/a */

  var selection = SlidesApp.getActivePresentation().getSelection();
  var selectionType = selection.getSelectionType();
  const regex = /"[a-zA-Z0-9 :!''’.?=$\[\]\(\)]+"/g;
  const text = ['print', 'if', 'elif', 'else', 'input'];

  if (theme == 0) {
    var theme = showPrompt();
  }

  if (selectionType == "PAGE_ELEMENT") {
    var currentPage = selection.getCurrentPage();
    var pageElements = selection.getPageElementRange().getPageElements(); // Gets all elements on a page

    if (pageElements.length == 1) {                                       // If you chose one element...
      var shape = pageElements[0].asShape();                              //    Turn it into a shape
      var textRange = shape.getText();
      textRange.getTextStyle().setForegroundColor('#000000');             //    Set all of the text in the shape to black
      var stringToSearch = textRange.asString();                          //    Get string inside shape
      var listOfWords = regexExtractWords(stringToSearch, regex);         //    Search for particular strings inside shape
      var textToChange = text.concat(listOfWords);                        //    Create mega list of all strings to replace

      for (let i = 0; i < textToChange.length; i++) {                     //    For each item in the list...
        var keyword = textRange.find(textToChange[i]);                    //        Find the words
        for (let j = 0; j < keyword.length; j++) {                        //        For each word...
          colour(text[i],keyword[j],theme);                                     //            Replace word
        }
      }
    } else {                                                              // If you chose too many elements
      Logger.log("Choose 1 textbox/shape.");
    }
  } else {                                                                // If you didnt choose a textbox or shape
    Logger.log("Choose a textbox/shape.");
  }
  SlidesApp.getUi();                                                      // Update menu items with this function
}

function regexExtractWords(text, regex) {
/* Provides a list of substrings that all match the regex passed in 

  Inputs:
      - text: a string
      - regex: the regular expression to match

  Outputs:
      - found: list of strings that match the regex */

  const found = text.match(regex);
  return found
}

function colour(searchWord, replaceWord, theme) {
/* Changes the replaceWord to the colour attached to the searchWord

  Inputs:
      - searchWord: the word you are searching for
      - replaceWord: the word you want to replace
      - theme: 'dark' or 'light' coding theme

  Outputs:
      - n/a */
  
  if (theme == 'light') {
    switch (searchWord) {
      case 'print':
        replaceWord.getTextStyle().setForegroundColor('#9900ff');
        break;
      case 'if':
        replaceWord.getTextStyle().setForegroundColor('#0000ff');
        break;
      case 'elif':
        replaceWord.getTextStyle().setForegroundColor('#0000ff');
        break;
      case 'else':
        replaceWord.getTextStyle().setForegroundColor('#0000ff');
        break;
      case 'input':
        replaceWord.getTextStyle().setForegroundColor('#9900ff');
        break;
      default:
        replaceWord.getTextStyle().setForegroundColor('#980000');
        break;
    } 
  } else if (theme == 'dark') {
    switch (searchWord) {
      case 'print':
        replaceWord.getTextStyle().setForegroundColor('#990000');
        break;
      case 'if':
        replaceWord.getTextStyle().setForegroundColor('#ff00ff');
        break;
      case 'elif':
        replaceWord.getTextStyle().setForegroundColor('#0340ff');
        break;
      case 'else':
        replaceWord.getTextStyle().setForegroundColor('#1750ff');
        break;
      case 'input':
        replaceWord.getTextStyle().setForegroundColor('#9a30ff');
        break;
      default:
        replaceWord.getTextStyle().setForegroundColor('#98f000');
        break;
    }
  }
}

////////// ARCHIVED //////////
/* Contains: showPrompt() */
function showPrompt() {
  /* Asks user what theme they want to change the code to

  Inputs:
      - n/a

  Outputs:
      - string stating what theme the user wants */

  var ui = SlidesApp.getUi(); // Same variations.

  var result = ui.prompt(
      'What would you like to edit?',
      'light, dark',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  /*if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('You will edit ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('You cancelled.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }*/
  return text;
}