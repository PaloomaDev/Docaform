/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  try {
    FormApp.getUi()
      .createAddonMenu()
      .addItem('Google Docs Template', 'displayDrivePicker')
      .addItem('Matching table', 'afficherTableCorrespondance')
      .addItem('Instructions', 'displayHelp')
      .addToUi();
  } catch (e) {
    // TODO (Developer) - Handle exception
    Logger.log('Failed with error: %s', e.error);
  }
}


/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
  adjustFormSubmitTrigger()
}

/**
 * Display the help 
 */
function displayHelp() {
  FormApp.getUi().showModelessDialog(HtmlService.createHtmlOutputFromFile("instructions").setWidth(1200).setHeight(1000), 'Help')
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
function adjustFormSubmitTrigger() {

  try {
    const form = FormApp.getActiveForm();
    const triggers = ScriptApp.getUserTriggers(form);
    // Create a new trigger if required; delete existing trigger
    // if it is not needed.
    let existingTrigger = null;
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
        existingTrigger = triggers[i];
        break;
      }
    }
    if (!existingTrigger) {
      const trigger = ScriptApp.newTrigger('updateDocWithAnswers')
        .forForm(form)
        .onFormSubmit()
        .create();
    } else {
      ScriptApp.deleteTrigger(existingTrigger)
      const trigger = ScriptApp.newTrigger('updateDocWithAnswers')
        .forForm(form)
        .onFormSubmit()
        .create();
    }
  } catch (e) {
    // TODO (Developer) - Handle exception
    Logger.log('Failed with error: %s', e.error);
  }
}





/**
 * Create the document based on the template
 *
 * @return {string} The document URL.
 */
function createDoc(name) {
  //Copie du modèle avec un nouveau nom et déplacement dans le dossier de destination
  return DriveApp.getFileById(PropertiesService.getDocumentProperties().getProperty("docId")).makeCopy(name).getId();

}

/**
 * update the created Doc With the respondant Answers
 *
 * @return {string} The document URL.
 */
function updateDocWithAnswers() {
  const addonTitle = 'Docaform';
  const props = PropertiesService.getDocumentProperties();
  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger requires authorization that has not
  // been granted yet; if so, warn the user via email. This check is required
  // when using triggers with add-ons to maintain functional triggers.
  if (authInfo.getAuthorizationStatus() ===
    ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to re-authorize; the normal trigger action is not
    // conducted, since it requires authorization first. Send at most one
    // "Authorization Required" email per day to avoid spamming users.
    const lastAuthEmailDate = props.getProperty('lastAuthEmailDate');
    const today = new Date().toDateString();
    if (lastAuthEmailDate !== today) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        const html = HtmlService.createTemplateFromFile('authEmail');
        html.url = authInfo.getAuthorizationUrl();
        html.addonTitle = addonTitle;
        const message = html.evaluate();
        MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
          name: addonTitle,
          htmlBody: message.getContent()
        }
        );
      }
      props.setProperty('lastAuthEmailDate', today);
    }
  } else {
    // Authorization has been granted, so continue to respond to the trigger.
    // Main trigger logic here.
    //FormApp.getActiveForm().setAcceptingResponses(false)

    let dateEvent = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/YYYY')
    let docId = createDoc("[" + dateEvent + "]-" + DriveApp.getFileById(PropertiesService.getDocumentProperties().getProperty("docId")).getName())
    getCorrespondanceTable().forEach(function (t) {
      let answer = getAnswerFromQuestion(t[0])
      if (answer != "") {
        Logger.log(t + "-" + answer)
        switch (t[2]) {
          case "PARAGRAPH_TEXT":
          case "TEXT":
            var txtRecherche = "{" + t[4] + "}"
            DocumentApp.openById(docId).getBody().replaceText(txtRecherche, answer)
            break
          case "MULTIPLE_CHOICE":
            //recherche de la question et de son index de colonne
            multipleChoice(t[3], answer, t[4], docId)
            break
          case "LIST":
            var txtRecherche = "{" + t[4] + "}"
            DocumentApp.openById(docId).getBody().replaceText(txtRecherche, answer)
            break

        }
      }
    })

  }
}





/**
 * Find  the multiple options in the Multiple Choix item
 *
 * @return {array} The item's choices.
 */
function findOptionInFormItem(itemId) {
  return FormApp.getActiveForm().getItemById(itemId).asMultipleChoiceItem().getChoices()
}



/**
 * Find the multiple choice label in the Google Docs template and remove the text in between 
 *
 * @return {} 
 */
function multipleChoice(questionChoice, answer, columnIndex, docId) {
  //Example : 
  //Multiple choice item : 
  //      Favorite Letter ? 
  //      Options : A,B,C
  //In the Google Docs template :
  // {A} this is the first letter{A}
  // {B} this is the second letter{B}
  // {C} this is the thrid letter{C}


  //if the questionChoice is A and the user chose A then we remove just the labels {A} and we keep the text in the middle
  //If not , we remove the labels and the text in between with the function removeBetweenSameElement
  if (answer != questionChoice) {
    removeBetweenSameElement("{" + columnIndex + "}", docId)
  }
  else {
    DocumentApp.openById(docId).getBody().replaceText("{" + columnIndex + "}", "")
  }
}


/**
 * Remove the labels and the text in between in a Google Gocs
 *
 * @return {} 
 */
function removeBetweenSameElement(elementTag, docId) {
  let doc = DocumentApp.openById(docId)
  let body = doc.getBody()
  let startElem = body.findText(elementTag)
  if (startElem != null) {
    startElem = startElem.getElement()
    let currentElement = startElem

    currentElement = currentElement.getParent().getNextSibling()
    Logger.log(currentElement.getText())
    while (currentElement.getText() != elementTag) {
      Logger.log(elementTag)
      Logger.log(currentElement.getText())

      currentElement.removeFromParent()

      currentElement = startElem.getParent().getNextSibling()

    }

  }
  cleanLabels(elementTag, docId)

}

function cleanLabels(elementTag, docId) {

  let body = DocumentApp.openById(docId).getBody()

  let n = body.getNumChildren();
  for (let i = n; i >= 0; i--) {
    try {
      let child = body.getChild(i);
      if (child.asText().findText(elementTag)) {
        try {
          child.removeFromParent();
        }
        catch (e) {
          Logger.log(n)
          Logger.log(e)
        }
        //--n;
      }
    }
    catch (e) {
      Logger.log(e)
    }
  }
  //Clean the last element 
  body.replaceText(elementTag, "")
}







/**
 * The function to get answer data from question and respondant id
 *
 * @param {string} questionID - question iD
 * @return {string}  the answer
**/
function getAnswerFromQuestion(questionId) {
  let answer = FormApp.getActiveForm().getResponses()[FormApp.getActiveForm().getResponses().length - 1].getItemResponses().filter(x => x.getItem().getId() == questionId)
  if (answer.length > 0) {
    return answer[0].getResponse()
  }
  else {
    return ""
  }
}














/**
 * To get the letter correspondance from an index 
 *
 * @param {integer} index 
 * @return {string}  the letter matching the index ( 1=A,2=B,...)
**/
function indexToLetter(index) {
  var temp, letter = '';
  while (index > 0) {
    temp = (index - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index = (index - temp - 1) / 26;
  }
  return letter;
}


/**
 * Create a matching table between the form items and the letter to let the user modify the Google Docs template with the right labels 
 *As of today, only 3 questions are available in the addon and therefore in the matching table 
 * @param {}  
 * @return {array}  the matching table
**/
function getCorrespondanceTable() {
  let table = []
  FormApp.getActiveForm().getItems().forEach((t, index) => {
    switch (t.getType().toString()) {
      case "PARAGRAPH_TEXT":
      case "TEXT":
        table.push([t.getId(), t.getTitle(), t.getType().toString(), "", indexToLetter(index + 1)])
        break
      case "MULTIPLE_CHOICE":
        let col = indexToLetter(index + 1)
        let reponses = findOptionInFormItem(t.getId())
        reponses.forEach((r, ind) => {

          table.push([t.getId(), t.getTitle(), t.getType().toString(), r.getValue(), col + (ind + 1)])
        })
        break
      case "LIST":
        table.push([t.getId(), t.getTitle(), t.getType().toString(), "", indexToLetter(index + 1)])
        break
    }
  })
  return table
}


/**
 * Set the Google Docs template ID in the forms (as document property) 
 * 
 * @param {}  
 * @return {array}  the matching table
**/
function updateGoogleDocsTemplateId(id) {
  PropertiesService.getDocumentProperties().setProperty("docId", id)
}



/**
 * Set the Google Docs template ID in the forms (as document property) 
 * 
 * @param {}  
 * @return {array}  the matching table
**/
function afficherTableCorrespondance() {
  let textAfficher = `<div style="margin-left:auto;margin-right:auto; width:100%; ">
<table style =" text-align:center; border:1px solid black; border-radius: 15px;  border-collapse: collapse; ">
  <tr style="text-align: center; padding: 15px;background-color: #D6EEEE;" >
    <th style="padding: 15px;">Question</th>
    <th style="padding: 15px;">Type</th>
    <th style="padding: 15px;">Choix</th>
    <th style="padding: 15px;">Label</th>
  </tr>`
  getCorrespondanceTable().forEach(function (t) {
    //textAfficher += t.join(" | ") + "\n"
    textAfficher += `<tr style ="border:1px solid black; border-radius: 15px;padding: 15px; text-align: center">
    <td style="padding: 15px;">`+ t[1] + `</td>
    <td style="padding: 15px;">`+ t[2] + `</td>
    <td style="padding: 15px;">`+ t[3] + `</td>
    <td style="padding: 15px;">`+ t[4] + `</td>
    
  </tr>`
  })
  textAfficher += `<table></div>`
  FormApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(textAfficher).setWidth(500).setHeight(500), 'Matching table')
}


/**
 * display the drive picker to set the Google Docs template
 * 
 * @param {}  
 * @return {}  
**/
function displayDrivePicker() {

  FormApp.getUi().showModelessDialog(HtmlService.createHtmlOutputFromFile("drivePicker"), 'Google Docs Template')

}


/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthTokenInit() {
  DriveApp.getFiles()

  return {
    oauthToken: ScriptApp.getOAuthToken(),
    // Add developer key from Google Cloud Project > APIs and services > Credentials
    developerKey: "AIzaSyDeNTwowoIBpJfmFUYRJJNRFWMxpcdhGO4"
  }
}


