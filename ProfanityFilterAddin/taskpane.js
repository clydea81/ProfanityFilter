Office.onReady(() => {
  const bannedWords = ["damn", "hell", "idiot"]; // Add more as needed

  function scanEmail() {
    Office.context.mailbox.item.body.getAsync("text", function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const body = result.value.toLowerCase();
        const found = bannedWords.some(word => body.includes(word));
        if (found) {
          Office.context.mailbox.item.notificationMessages.addAsync("ProfanityWarning", {
            type: "informationalMessage",
            message: "Profanity detected. Please revise your message.",
            icon: "icon16",
            persistent: true
          });
        }
      }
    });
  }

  // Scan every 10 seconds while composing
  setInterval(scanEmail, 10000);
});