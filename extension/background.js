console.log("[dash-ext] background alive");
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.type !== "DOWNLOAD_JSON") return;

  try {
    const filename = msg.filename || "parse.json";
    const text = typeof msg.text === "string" ? msg.text : "";

    const url = "data:application/json;charset=utf-8," + encodeURIComponent(text);

    chrome.downloads.download(
      {
        url,
        filename,
        saveAs: false
      },
      (downloadId) => {
        const err = chrome.runtime.lastError?.message;
        if (err || !downloadId) {
          sendResponse({ ok: false, error: err || "download failed" });
          return;
        }
        sendResponse({ ok: true, downloadId });
      }
    );

    return true; // async response
  } catch (e) {
    sendResponse({ ok: false, error: String(e) });
  }
});
