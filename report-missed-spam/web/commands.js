/*
 * Report Missed Spam - function command.
 *
 * 1. (optional) Shows a dialog asking whether the user released the message from quarantine.
 *      Yes -> explanation shown, nothing is reported.
 *      No  -> continues.
 * 2. Opens a new compose window addressed to the false-negative mailbox with the selected
 *    message attached as an Outlook item (not a forward). The user reviews and clicks Send.
 *
 * Wording, recipient and the on/off switch for the dialog live in config.js (CONFIG).
 */

const NOTIFY_ID = "reportMissedSpam";
const MAX_SUBJECT = 200;

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fillTemplate(template, values, escape) {
  return template.replace(/\{(\w+)\}/g, (match, key) => {
    if (!Object.prototype.hasOwnProperty.call(values, key)) return match;
    return escape ? escapeHtml(values[key]) : values[key];
  });
}

function truncate(text, max) {
  return text.length > max ? text.slice(0, max - 1) + "…" : text;
}

function notify(item, type, message) {
  try {
    item.notificationMessages.replaceAsync(NOTIFY_ID, {
      type: type, // "informationalMessage" | "errorMessage"
      message: message,
      icon: "Icon.16",
      persistent: false,
    });
  } catch (e) {
    /* notifications are best effort */
  }
}

/*
 * Shows dialog.html and calls done(true) to continue or done(false) to stop.
 * If the dialog can't be opened we continue (fail open) so the button never goes dead.
 * Closing the dialog with the X stops without reporting.
 */
function askQuarantine(done) {
  let finished = false;
  function finish(proceed) {
    if (finished) return;
    finished = true;
    done(proceed);
  }

  const url = new URL("dialog.html", window.location.href).toString();
  Office.context.ui.displayDialogAsync(
    url,
    { height: 32, width: 28, displayInIframe: false },
    function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        finish(true);
        return;
      }
      const dialog = result.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        let action = "close";
        try {
          action = JSON.parse(arg.message).action;
        } catch (e) {
          /* treat unreadable messages as close */
        }
        try {
          dialog.close();
        } catch (e) {
          /* already closed */
        }
        finish(action === "continue");
      });

      // User closed the dialog with the X (or it errored): do not report.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        finish(false);
      });
    }
  );
}

function openCompose(item, event) {
  try {
    const originalSubject = item.subject || "(no subject)";
    const from = item.from;
    const sender = from
      ? from.displayName && from.displayName !== from.emailAddress
        ? from.displayName + " <" + from.emailAddress + ">"
        : from.emailAddress
      : "(unknown)";
    const date = item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : "(unknown)";

    const values = {
      reporter: Office.context.mailbox.userProfile.emailAddress || "",
      sender: sender,
      subject: truncate(originalSubject, MAX_SUBJECT),
      date: date,
    };

    const attachmentName = truncate(originalSubject.replace(/[\\/:*?"<>|\r\n]+/g, "-"), MAX_SUBJECT);

    Office.context.mailbox.displayNewMessageFormAsync(
      {
        toRecipients: [CONFIG.reportTo],
        subject: truncate(fillTemplate(CONFIG.subjectTemplate, values, false), 255),
        htmlBody: fillTemplate(CONFIG.bodyTemplate, values, true),
        attachments: [
          {
            type: Office.MailboxEnums.AttachmentType.Item,
            name: attachmentName,
            itemId: item.itemId, // EWS item ID, as required for item attachments
          },
        ],
      },
      function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          notify(item, "errorMessage", "Couldn't open the report message. Please try again.");
        }
        event.completed();
      }
    );
  } catch (e) {
    notify(item, "errorMessage", "Couldn't open the report message. Please try again.");
    event.completed();
  }
}

function reportMissedSpam(event) {
  const item = Office.context.mailbox.item;

  if (!item || !item.itemId) {
    notify(
      item,
      "errorMessage",
      "This message can't be reported from here. Open it from your mailbox and try again."
    );
    event.completed();
    return;
  }

  const prompt = CONFIG.quarantinePrompt;
  if (!prompt || !prompt.enabled) {
    openCompose(item, event);
    return;
  }

  // event.completed() is held until the whole flow ends so the runtime stays alive for the dialog.
  askQuarantine(function (proceed) {
    if (proceed) {
      openCompose(item, event);
    } else {
      event.completed();
    }
  });
}

// Must be registered globally so the manifest's FunctionName can find it.
Office.actions.associate("reportMissedSpam", reportMissedSpam);
