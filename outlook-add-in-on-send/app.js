var mailboxItem;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
};

function calculateMeetingCost() {
  // Read and parse the 'pay.json' file containing email, name, and hourly rate data
  fetch("pay.json")
    .then((response) => response.json())
    .then((data) => {
      var attendees = mailboxItem.requiredAttendees;
      var duration = mailboxItem.start.getDuration();

      var meetingCost = 0;

      attendees.forEach(function (attendee) {
        var email = attendee.emailAddress;
        if (data[email]) {
          var hourlyRate = data[email].hourlyRate;
          var cost = hourlyRate * duration;
          meetingCost += cost;
        }
      });

      // Update the user interface or display the meeting cost
      console.log("Meeting Cost: $" + meetingCost.toFixed(2));
    })
    .catch((error) => console.error(error));
}
