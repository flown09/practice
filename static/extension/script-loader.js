var script = document.createElement("script");
script.src = chrome.runtime.getURL('/sd.js');
script.setAttribute("type", "text/javascript");
document.body.appendChild(script);
