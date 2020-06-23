function onOpen(e) {
  DocumentApp.getUi().
      createAddonMenu().
      addItem('Show Sidebar', 'showSidebar').
      addItem('AWS Credentials', 'showSettings').
      addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('addon/Sidebar').
      evaluate().
      setTitle('Speaking Docs');

  DocumentApp.getUi().showSidebar(ui);
}

function showSettings() {
  var ui = HtmlService.createTemplateFromFile('addon/Settings').
      evaluate().
      setHeight(250);

  DocumentApp.getUi().showModalDialog(ui, 'Edit AWS credentials');
}

function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var text = [];
  if (selection) {
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        //  skip images and other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText) {
            text.push(elementText);
          }
        }
      }
    }
  }
  return text;
}

function replaceSelection( newText ) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if ( selection ) {
    var elements = selection.getRangeElements();
    var replace = true;
    for ( var i = 0; i < elements.length; i ++ ) {
      if ( elements[i].isPartial() ) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        var text = element.getText().substring( startIndex, endIndex + 1 );
        element.deleteText( startIndex, endIndex );
        if ( replace ) {
          element.insertText( startIndex, newText );
          replace = false;
        }
      } else {
        var element = elements[i].getElement();
        if ( replace && element.editAsText ) {
          element.clear().asText().setText( newText );
          replace = false;
        } else {
          if ( replace && i === elements.length - 1 ) {
            var parent = element.getParent();
            parent[parent.insertText ? 'insertText' : 'insertParagraph']( parent.getChildIndex( element ), newText );
            replace = false; //not really necessary since it's the last one
          }
          element.removeFromParent();
        }
      }
    }
  } else {
    throw "Hey, select something so I can replace!";
  }
}

function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();

  return {
    awsRegion: userProperties.getProperty('awsRegion'),
    awsAccessKeyId: userProperties.getProperty('awsAccessKeyId'),
    awsSecretAccessKey: userProperties.getProperty('awsSecretAccessKey'),
  };
}

function setPreferences(options) {
  var userProperties = PropertiesService.getUserProperties();

  userProperties.setProperty('awsRegion', options.awsRegion);
  userProperties.setProperty('awsAccessKeyId', options.awsAccessKeyId);
  userProperties.setProperty('awsSecretAccessKey', options.awsSecretAccessKey);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
