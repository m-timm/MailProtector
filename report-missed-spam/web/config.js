/*
 * Report Missed Spam - settings. EDIT THIS FILE to change wording or behavior.
 * Re-host the web/ folder after editing; no manifest change or tenant redeploy is needed.
 * (Users may need to restart Outlook to pick up a cached copy.)
 *
 * Email tokens (subjectTemplate / bodyTemplate; values are HTML-escaped in the body):
 *   {reporter}  the signed-in user's email address
 *   {sender}    original sender, "Name <address>"
 *   {subject}   original subject
 *   {date}      original message date
 */
const CONFIG = {
  reportTo: "Spam@emailservice.io",

  subjectTemplate: "Missed spam report: {subject}",

  // Body is HTML, max 32 KB.
  bodyTemplate:
    "<p>Reporting the attached message as missed spam (false negative).</p>" +
    "<p>" +
    "<b>Reported by:</b> {reporter}<br>" +
    "<b>Original sender:</b> {sender}<br>" +
    "<b>Original subject:</b> {subject}<br>" +
    "<b>Date:</b> {date}" +
    "</p>" +
    "<p>Please use the attached message to build new detection rules.</p>",

  // Question shown before the compose window opens. Set enabled:false to skip it.
  // Plain text only.
  quarantinePrompt: {
    enabled: true,
    title: "Report Missed Spam",
    question: "Did you release this message from your quarantine?",
    yesLabel: "Yes",
    noLabel: "No",
    yesMessage:
      "Mailprotector already quarantined this message. This button is only to be used in the " +
      "event a malicious email is not quarantined. Please block the sender in your digest report.",
    closeLabel: "Close",
  },
};
