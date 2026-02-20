const APP_UPLOAD_URL = "http://127.0.0.1:8000/upload-parse-raw";

console.log("[dash-ext] background alive");
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.type !== "UPLOAD_PARSE") return;
  (async () => {
    try {
      const text = typeof msg.text === "string" ? msg.text : "";

      const fd = new FormData();
      fd.append("file", new Blob([text], { type: "application/json;charset=utf-8" }), "parse.json");

      const resp = await fetch(APP_UPLOAD_URL, { method: "POST", body: fd });
      const data = await resp.json().catch(() => ({}));

      if (!resp.ok) {
        sendResponse({ ok: false, status: resp.status, error: data.detail || JSON.stringify(data) });
        return;
      }

      const finalUrl = data.download_url ? new URL(data.download_url, APP_UPLOAD_URL).toString() : null;

      // опционально: сразу скачать финальный отчет
      if (finalUrl) {
        chrome.downloads.download({ url: finalUrl, filename: "финальный_отчет.xlsx", saveAs: false }, () => {});
      }

      sendResponse({ ok: true, upload_id: data.upload_id, download_url: finalUrl });
    } catch (e) {
      sendResponse({ ok: false, error: String(e) });
    }
  })();

  return true;


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
