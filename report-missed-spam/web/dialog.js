/*
 * Dialog shown before the compose window opens (one dialog, two screens).
 *   Screen 1: "Did you release this message from your quarantine?"  [Yes] [No]
 *   Yes -> screen 2 with the explanation and a Close button; nothing is reported.
 *   No  -> tells the add-in to continue and open the compose window.
 * Wording comes from CONFIG.quarantinePrompt in config.js. All text is set via textContent.
 */

function send(action) {
  Office.context.ui.messageParent(JSON.stringify({ action: action }));
}

function setScreen(text, buttons) {
  document.getElementById("text").textContent = text;
  const box = document.getElementById("buttons");
  box.textContent = "";
  buttons.forEach(function (b, i) {
    const el = document.createElement("button");
    el.textContent = b.label;
    if (b.primary) el.className = "primary";
    el.addEventListener("click", b.onClick);
    box.appendChild(el);
    if (i === buttons.length - 1 || b.primary) el.dataset.focus = "1";
  });
  const focus = box.querySelector("[data-focus]") || box.firstChild;
  if (focus) focus.focus();
}

function showQuestion() {
  const p = CONFIG.quarantinePrompt;
  setScreen(p.question, [
    { label: p.yesLabel, onClick: showQuarantined },
    { label: p.noLabel, primary: true, onClick: function () { send("continue"); } },
  ]);
}

function showQuarantined() {
  const p = CONFIG.quarantinePrompt;
  setScreen(p.yesMessage, [
    { label: p.closeLabel, primary: true, onClick: function () { send("close"); } },
  ]);
}

function showPhishTest() {
  const p = CONFIG.phishTest;
  document.getElementById("title").textContent = p.title || CONFIG.quarantinePrompt.title;
  setScreen(p.message, [
    { label: p.closeLabel || CONFIG.quarantinePrompt.closeLabel, primary: true, onClick: function () { send("close"); } },
  ]);
}

Office.onReady(function () {
  if (new URLSearchParams(window.location.search).get("mode") === "phishtest" && CONFIG.phishTest) {
    showPhishTest();
    return;
  }
  document.getElementById("title").textContent = CONFIG.quarantinePrompt.title;
  showQuestion();
});
