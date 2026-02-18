console.log("[dash-ext] content script loaded", location.href);

// 1) инжект sd.js в контекст страницы
const src = chrome.runtime.getURL("sd.js");
console.log("[dash-ext] injecting:", src);

const script = document.createElement("script");
script.src = src;
script.type = "text/javascript";

script.onload = () => console.log("[dash-ext] injected sd.js onload");
script.onerror = (e) => console.error("[dash-ext] injected sd.js error", e);

(document.head || document.documentElement).appendChild(script);

// 2) мост: sd.js (page) -> content script -> background
window.addEventListener("message", (event) => {
  if (event.source !== window) return;

  const data = event.data;
  if (!data || data.__from !== "dashbord-ext") return;

  console.log("[dash-ext] got window message:", data.type);

  if (data.type === "UPLOAD_PARSE") {
    chrome.runtime.sendMessage({ type: "UPLOAD_PARSE", text: data.text }, (resp) => {
      const err = chrome.runtime.lastError?.message;
      const payload = err ? { ok: false, error: err } : resp;
      window.postMessage({ __from: "dashbord-ext", type: "UPLOAD_RESULT", payload }, "*");
    });
  }


  if (data.type === "DOWNLOAD_JSON") {
    chrome.runtime.sendMessage(
      { type: "DOWNLOAD_JSON", filename: data.filename, text: data.text },
      (resp) => {
        const err = chrome.runtime.lastError?.message;
        if (err) console.error("[dash-ext] runtime message error:", err);
        else console.log("[dash-ext] bg resp:", resp);
      }
    );
  }
});
