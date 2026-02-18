const blobUrlByDownloadId = new Map();

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.type !== "DOWNLOAD_JSON") return;

  try {
    const filename = msg.filename || "parse.json";
    const text = typeof msg.text === "string" ? msg.text : "";

    // Надёжнее, чем data: URL, если json большой
    const blob = new Blob([text], { type: "application/json;charset=utf-8" });
    const blobUrl = URL.createObjectURL(blob);

    chrome.downloads.download(
      {
        url: blobUrl,
        filename,
        saveAs: false
      },
      (downloadId) => {
        const err = chrome.runtime.lastError?.message;

        if (err || !downloadId) {
          // если не получилось — освобождаем blobUrl
          try { URL.revokeObjectURL(blobUrl); } catch {}
          sendResponse({ ok: false, error: err || "download failed" });
          return;
        }

        blobUrlByDownloadId.set(downloadId, blobUrl);
        sendResponse({ ok: true, downloadId });
      }
    );

    return true; // async
  } catch (e) {
    sendResponse({ ok: false, error: String(e) });
  }
});

chrome.downloads.onChanged.addListener((delta) => {
  if (!delta || !delta.id) return;
  if (!delta.state) return;

  const state = delta.state.current;
  if (state === "complete" || state === "interrupted") {
    const url = blobUrlByDownloadId.get(delta.id);
    if (url) {
      blobUrlByDownloadId.delete(delta.id);
      try { URL.revokeObjectURL(url); } catch {}
    }
  }
});
