# Google-Sheets-Email-Signup
100% free and working script for collecting any data via simple API request to Google Sheets

# How to use

## Google Sheet setup

 * From your Google Sheet, from the "Extensions" menu select "Apps Script"
 * Paste code from `index.js` into the script code editor and hit Save icon.
 * From the "Deploy" menu, select Deploy as web app
 * Choose to execute the app as yourself, and allow "Anyone", even anonymous to execute the script.
 * Now click Deploy. You may be asked to review permissions now.
 * The URL that you get will be the webhook that you can use in your POST request anywhere.
 * You can test this webhook in your browser first by pasting it. It will say "Use POST method to send data to this URL.".
 * Last all you have to do is set up your request function in your app or website.

## Frontend request setup

```js
const signupRequest = async () => {
  const email = document.querySelector("input[name='email']")?.value;
  const WEBHOOK_URL = 'https://script.google.com/macros/s/.../exec';
  if (email) {
    try {
      await fetch(WEBHOOK_URL,
        {
          method: "POST",
          mode: "no-cors", // Important for Google Apps Script when not handling CORS on server
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: new URLSearchParams({
            email: email,
            form_name: "waitlist", // Add a form_name here
            date: new Date().toLocaleString(),
          }),
        }
      );
      alert("Thank you for signing up! We will notify you when we launch.");
    } catch (error) {
      alert("An error occurred. Please try again later.");
    }
  }
};
```
