function boilerPlate() {
  Logger.log("Hello, I am boilerplate!");
  return;
}

function logCurrentDayOfWeek() {
  var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var currentDate = new Date(); // Gets the current date and time
  var dayOfWeekIndex = currentDate.getDay(); // Gets the day of the week as a number (0-6)
  var dayOfWeekName = daysOfWeek[dayOfWeekIndex]; // Maps the number to the corresponding day of the week

  Logger.log(dayOfWeekName); // Logs the day of the week
}

function extractDayOfWeek(response) {
  var daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  for (var i = 0; i < daysOfWeek.length; i++) {
    if (response.includes(daysOfWeek[i])) {
      Logger.log("Extracted Day: " + daysOfWeek[i] + " from Response: " + response);
      return daysOfWeek[i];
    }
  }
  Logger.log("No Day of the Week Found in Response: " + response);
  return '';
}
function callOpenAI(inputText, model) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const data = {
    model: model,
    messages: [{ "role": "user", "content": inputText }]
  };

  const headers = {
    'Authorization': 'Bearer ' + apiKey
  };

  // Conditionally add organization header
  if (organizationId) {
    headers['OpenAI-Organization'] = organizationId;
  }

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(data),
    'headers': headers
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.choices && result.choices.length > 0) {
      return result.choices[0].message.content.replace(/^\n\n/, '');
    } else {
      Logger.log("No response from OpenAI API");
      return null;
    }
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return null;
  }
}
function testPrompt() {
  Logger.log(callOpenAI('How can I make delicious japanese style shabu shabu?', 'gpt-4'));
}