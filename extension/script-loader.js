console.log("[dash-ext] content script loaded");

window.addEventListener("message", (event) => {
  if (event.source !== window) return;

  const data = event.data;
  if (!data || data.__from !== "dashbord-ext") return;

  console.log("[dash-ext] got window message:", data);

  if (data.type === "DOWNLOAD_JSON") {
    chrome.runtime.sendMessage(
      {
        type: "DOWNLOAD_JSON",
        filename: data.filename,
        text: data.text
      },
      () => {
        const err = chrome.runtime.lastError?.message;
        if (err) console.error("[dash-ext] runtime message error:", err);
        else console.log("[dash-ext] message sent to background ok");
      }
    );
  }
});

chrome.runtime.sendMessage(
  { type: "DOWNLOAD_JSON", filename: "cs-test.json", text: '{"from":"content-script"}' },
  (resp) => console.log("[dash-ext] cs-test resp:", resp, chrome.runtime.lastError?.message)
);
