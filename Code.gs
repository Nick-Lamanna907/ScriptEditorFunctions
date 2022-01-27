const PRESENTATION_ID = '1qLn8K7VHbxVx0oF46ETOGFUInneVMPPftmZ7GArQf_rbT5CBxcnCRIcF'

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('My Functions')
      .addSubMenu(ui.createMenu('Code Slides')
          .addItem('Format Code [All Slides]', 'formatCodeColourAllSlides')
          .addItem('Format Code [Selected Item]', 'formatCodeColour'))
      .addToUi();
}

////////// UPDATE CODE TEXT COLOURS //////////
/* Contains: formatCodeColourAllSlides(), formatCodeColour(), regexExtractWords(text, regex), colour(searchWord, replaceWord)
             showPrompt(), searchForExtras() */
function formatCodeColourAllSlides() {
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
          formatCodeColour(theme);
        }
      }
    }
  }
}

function formatCodeColour(theme = 0) {
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
  const regexNum = /[0-9]/g;
  const regexComment = /#[a-zA-Z0-9 :!''’.?=$\[\]\(\)]+/g;
  const regexLabel = /\.[a-zA-Z0-9 :!''’.?=$\[\]\(\)_]+/g;
  const text = [ 'int', 'str', 'print', 'input', 'if', 'elif', 'else', 'from', 'import'];

  if (theme == 0) {
    var theme = showPrompt();
  }

  if (selectionType == "PAGE_ELEMENT") {
    var currentPage = selection.getCurrentPage();
    var pageElements = selection.getPageElementRange().getPageElements(); // Gets all elements on a page

    if (pageElements.length == 1) {                                       // If you chose one element...
      var shape = pageElements[0].asShape();                              //    Turn it into a shape
      var textRange = shape.getText();                                    //    Get text
      
      if (theme == 'light') {textRange.getTextStyle().setForegroundColor('#000000');} //Set all of the text in the shape to black
      if (theme == 'dark')  {textRange.getTextStyle().setForegroundColor('#ffffff');} //Set all of the text in the shape to white
      
      var stringToSearch = textRange.asString();                          //    Get string inside shape

      searchForExtras(textRange, stringToSearch, regexNum, theme)         //    Replace all numbers first (later str's are handled)

      var listOfWords = regexExtractWords(stringToSearch, regex);         //    Search for particular strings inside shape
      var textToChange = text.concat(listOfWords);                        //    Create mega list of all strings to replace
      for (let i = 0; i < textToChange.length; i++) {                     //    For each item in the list...
        var keyword = textRange.find(textToChange[i]);                    //        Find the words
        for (let j = 0; j < keyword.length; j++) {                        //        For each word...
          colour(text[i],keyword[j],theme);                               //            Replace word
        }
      }

      searchForExtras(textRange, stringToSearch, regexLabel, theme);      //    Change colour of anything following a '#'
      searchForExtras(textRange, stringToSearch, regexComment, theme);    //    Change colour of anything following a '.'

    } else {                                                              // If you chose too many elements
      Logger.log("Choose 1 textbox/shape.");
    }
  } else {                                                                // If you didnt choose a textbox or shape
    Logger.log("Choose a textbox/shape.");
  }
  SlidesApp.getUi();                                                      // Update menu items with this function
}

function searchForExtras(textRange, stringToSearch, regex, theme){
/* Searches for and replaces anything matching the regex

  Inputs:
      - textRange: a text range object in which all strings that need to be replaced are contained
      - stringToSearch: the raw string to search
      - regex: the regular expression being used to dectect strings to be replaced
      - theme: the colour theme the user wants to use

  Outputs:
      - n/a */

  var listToReplace = regexExtractWords(stringToSearch, regex);           // Search for particular comments inside shape
  if (listToReplace != null) {
    for (let i = 0; i < listToReplace.length; i++) {
      var keyword = textRange.find(listToReplace[i]);                     //        Find the extras 
      for (let j = 0; j < keyword.length; j++) { 
        colour('number', keyword[j], theme);     
      }
    }
    Logger.log('Extras done: ' + listToReplace)
  }    
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
      case 'input':
        replaceWord.getTextStyle().setForegroundColor('#9900ff');
        break;
      case 'if':
      case 'elif':
      case 'else':
      case 'from':
      case 'import':
        replaceWord.getTextStyle().setForegroundColor('#0000ff');
        break;
      case 'int':
      case 'str':
        replaceWord.getTextStyle().setForegroundColor('#45818e');
        break;
      case 'comment':
        replaceWord.getTextStyle().setForegroundColor('#bf9000');
        break;
      case 'number':
        replaceWord.getTextStyle().setForegroundColor('#6aa84f');
        break
      default:
        replaceWord.getTextStyle().setForegroundColor('#980000');
        break;
    } 
  } else if (theme == 'dark') {
    switch (searchWord) {
      case 'print':
      case 'input':
        replaceWord.getTextStyle().setForegroundColor('#ffd966');
        break;
      case 'if':
      case 'elif':
      case 'else':
      case 'from':
      case 'import':
        replaceWord.getTextStyle().setForegroundColor('#6d9eeb');
        break;
      case 'int':
      case 'str':
        replaceWord.getTextStyle().setForegroundColor('#30ddae');
        break;
      case 'comment':
        replaceWord.getTextStyle().setForegroundColor('#f48fb1');
        break;
      case 'number':
        replaceWord.getTextStyle().setForegroundColor('#93c47d');
        break
      default:
        replaceWord.getTextStyle().setForegroundColor('#e06666');
        break;
    } 
  }
}

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
