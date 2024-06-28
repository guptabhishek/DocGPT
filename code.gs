const OPENAI_API_KEY = 'YOUR_OPEN_AI_KEY';
const MODEL = 'gpt-3.5-turbo-0125'; // Adjust model as needed
//const MODEL = 'gpt-4o-2024-05-13';
let doc;

function onOpen() {
  DocumentApp.getUi()
      .createMenu('OpenAI')
      .addItem('Start Conversation', 'startConversation')
      .addToUi();
}

function startConversation() {
  doc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Start Conversation with OpenAI:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var userInput = response.getResponseText();
    console.log(userInput);
    askOpenAI(userInput);
  }
}

function startConversationWithInitialText() {
  doc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Enter Initial Text for Conversation:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var initialText = response.getResponseText();
    appendInitialTextToDocument(initialText);
  }
}

function askOpenAI(userInput, isInitialText = false) {
  var url = 'https://api.openai.com/v1/chat/completions';
  var headers = {
    'Authorization': 'Bearer ' + OPENAI_API_KEY,
    'Content-Type': 'application/json'
  };
  
  let conversationHistory = getConversationHistory();
  if (!isInitialText) {
    conversationHistory.push({ 'role': 'user', 'content': userInput });
  }
  
  var data = {
    'model': MODEL,
    'messages': conversationHistory,
    'temperature': 0.7,
  };
  
  var options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(data),
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.choices && result.choices.length > 0 && result.choices[0].message && result.choices[0].message.content !== undefined) {
    var aiResponse = result.choices[0].message.content.trim();
    appendToDocument(userInput, aiResponse, isInitialText);
  } else {
    appendToDocument(userInput, "Error: Unable to retrieve response from OpenAI", isInitialText);
  }
}

function appendToDocument(userInput, aiResponse, isInitialText = false) {
  if (!isInitialText) {
    doc.getBody().appendParagraph('User: ' + userInput);
  }
  doc.getBody().appendParagraph('OpenAI says: ' + aiResponse);
  doc.getBody().appendParagraph(''); // Add an empty paragraph for spacing
  
  // Ensure the text is visible
  var paragraphs = doc.getBody().getParagraphs();
  paragraphs[paragraphs.length - 3].setFontSize(12).setForegroundColor("#000000");
  paragraphs[paragraphs.length - 2].setFontSize(12).setForegroundColor("#000000");
}

function appendInitialTextToDocument(initialText) {
  doc.getBody().appendParagraph('Initial Text: ' + initialText);
  doc.getBody().appendParagraph(''); // Add an empty paragraph for spacing
  // Do not call askOpenAI immediately; let the user initiate the next query
}

function getConversationHistory() {
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  var conversationHistory = [];
  
  paragraphs.forEach(paragraph => {
    var text = paragraph.getText();
    if (text.startsWith('User:')) {
      conversationHistory.push({ 'role': 'user', 'content': text.substring(6).trim() });
    } else if (text.startsWith('OpenAI says:')) {
      conversationHistory.push({ 'role': 'assistant', 'content': text.substring(12).trim() });
    } else if (text.startsWith('Initial Text:')) {
      conversationHistory.push({ 'role': 'system', 'content': text.substring(13).trim() });
    }
  });
  
  return conversationHistory;
}

