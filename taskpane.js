/* global Office, document, navigator */

let deepLink = "";
let emailSubject = "";

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  emailSubject = item.subject || "(no subject)";
  document.getElementById("subject").textContent = emailSubject;

  deepLink = buildDeepLink();
  document.getElementById("linkbox").textContent = deepLink;

  loadSettings();

  document.getElementById("copyBtn").addEventListener("click", onCopyClick);
  document.getElementById("asanaBtn").addEventListener("click", onAsanaClick);
  document.getElementById("settingsToggle").addEventListener("click", () => {
    document.getElementById("settings").classList.toggle("open");
  });
  document.getElementById("saveSettings").addEventListener("click", saveSettings);
});

// ---------- Deep link ----------

function buildDeepLink() {
  const item = Office.context.mailbox.item;

  // On iOS the itemId is already REST-formatted; everywhere else (Windows
  // classic Outlook, new Outlook, Outlook on the web) it's an EWS-style id.
  // The OWA "owa/?ItemID=" reader link expects the raw EWS-style id, so we
  // do NOT convert it — we just percent-encode it as-is.
  const rawId = item.itemId;

  // restUrl looks like "https://outlook.office.com" — same backend that
  // outlook.cloud.microsoft and outlook.office365.com front onto, so this
  // link works no matter which of those you normally browse to.
  const host = (Office.context.mailbox.restUrl || "https://outlook.office.com").replace(/\/$/, "");

  return `${host}/owa/?ItemID=${encodeURIComponent(rawId)}&exvsurl=1&viewmodel=ReadMessageItem`;
}

// ---------- Copy to clipboard ----------

function onCopyClick() {
  const statusEl = document.getElementById("copyStatus");
  copyText(deepLink)
    .then(() => setStatus(statusEl, "Copied to clipboard.", "ok"))
    .catch(() => setStatus(statusEl, "Couldn't copy automatically — select the link above and copy manually.", "err"));
}

function copyText(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    return navigator.clipboard.writeText(text);
  }
  // Fallback for hosts without the async Clipboard API.
  return new Promise((resolve, reject) => {
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      ok ? resolve() : reject(new Error("execCommand copy failed"));
    } catch (e) {
      reject(e);
    }
  });
}

// ---------- Settings (roaming — synced to this mailbox across desktop & web) ----------

function loadSettings() {
  const rs = Office.context.roamingSettings;
  document.getElementById("pat").value = rs.get("asanaPat") || "";
  document.getElementById("project").value = rs.get("asanaProject") || "";
  document.getElementById("workspace").value = rs.get("asanaWorkspace") || "";
  // Auto-open settings the first time, so the user notices they need to configure it.
  if (!rs.get("asanaPat")) {
    document.getElementById("settings").classList.add("open");
  }
}

function saveSettings() {
  const rs = Office.context.roamingSettings;
  rs.set("asanaPat", document.getElementById("pat").value.trim());
  rs.set("asanaProject", document.getElementById("project").value.trim());
  rs.set("asanaWorkspace", document.getElementById("workspace").value.trim());
  rs.saveAsync((result) => {
    const statusEl = document.getElementById("settingsStatus");
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      setStatus(statusEl, "Settings saved.", "ok");
    } else {
      setStatus(statusEl, "Couldn't save settings: " + result.error.message, "err");
    }
  });
}

// ---------- Create Asana task ----------

function onAsanaClick() {
  const statusEl = document.getElementById("asanaStatus");
  const rs = Office.context.roamingSettings;
  const pat = rs.get("asanaPat");
  const project = rs.get("asanaProject");
  const workspace = rs.get("asanaWorkspace");

  if (!pat) {
    setStatus(statusEl, "Add your Asana token in ⚙ Asana settings first.", "err");
    document.getElementById("settings").classList.add("open");
    return;
  }
  if (!project && !workspace) {
    setStatus(statusEl, "Add a Project GID (or Workspace GID) in ⚙ Asana settings first.", "err");
    document.getElementById("settings").classList.add("open");
    return;
  }

  const btn = document.getElementById("asanaBtn");
  btn.disabled = true;
  setStatus(statusEl, "Creating task…", "");

  const data = {
    name: emailSubject,
    notes: deepLink
  };
  if (project) {
    data.projects = [project];
  } else {
    data.workspace = workspace;
    data.assignee = "me";
  }

  fetch("https://app.asana.com/api/1.0/tasks", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + pat
    },
    body: JSON.stringify({ data })
  })
    .then(async (resp) => {
      const body = await resp.json();
      if (!resp.ok) {
        const msg = (body.errors && body.errors[0] && body.errors[0].message) || resp.statusText;
        throw new Error(msg);
      }
      return body;
    })
    .then((body) => {
      const url = body.data && body.data.permalink_url;
      setStatus(
        statusEl,
        url ? "Task created ✓ (click to open)" : "Task created ✓",
        "ok"
      );
      if (url) {
        statusEl.style.cursor = "pointer";
        statusEl.onclick = () => window.open(url, "_blank");
      }
    })
    .catch((err) => {
      setStatus(statusEl, "Asana error: " + err.message, "err");
    })
    .finally(() => {
      btn.disabled = false;
    });
}

// ---------- Helpers ----------

function setStatus(el, text, cls) {
  el.textContent = text;
  el.className = "status" + (cls ? " " + cls : "");
}
