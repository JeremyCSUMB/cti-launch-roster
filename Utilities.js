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
