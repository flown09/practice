// 1) инжект sd.js в контекст страницы
const script = document.createElement("script");
script.src = chrome.runtime.getURL("sd.js");
script.type = "text/javascript";
(document.head || document.documentElement).appendChild(script);

// 2) мост: sd.js (page) -> content script -> background
window.addEventListener("message", (event) => {
  // сообщения должны быть от текущего окна и с нашего origin
  if (event.source !== window) return;
  if (event.origin !== window.location.origin) return;

  const data = event.data;
  if (!data || data.__from !== "dashbord-ext") return;

  if (data.type === "DOWNLOAD_JSON") {
    chrome.runtime.sendMessage(
      {
        type: "DOWNLOAD_JSON",
        filename: data.filename,
        text: data.text
      },
      () => {
        // ответ можно игнорировать, но при желивании — логировать ошибки
        const err = chrome.runtime.lastError?.message;
        if (err) console.error("runtime message error:", err);
      }
    );
  }
});
