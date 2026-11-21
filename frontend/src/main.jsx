import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

// ── Auto refresh every 10 minutes ────────────────────────────────
// The app stays open for hours during a service, so it is reloaded
// periodically to pick up new assets and drop any leaked memory.
// Session state is preserved first so an in-progress song is not lost.
const AUTO_REFRESH_MS = 10 * 60 * 1000;

function snapshotDrafts() {
  try {
    const data = {};
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key && key.startsWith("worshiptool:")) {
        data[key] = localStorage.getItem(key);
      }
    }
    sessionStorage.setItem(
      "worshiptool:autorefresh",
      JSON.stringify({ at: Date.now(), data }),
    );
  } catch {
    // Storage can be unavailable (private mode); a plain reload still works.
  }
}

function startAutoRefresh() {
  const timer = setInterval(() => {
    snapshotDrafts();
    window.location.reload();
  }, AUTO_REFRESH_MS);

  // Reload straight away if the tab has been hidden/backgrounded for longer
  // than the interval — browsers throttle timers in background tabs, so the
  // 10-minute tick may not have fired while the presenter was away.
  let hiddenAt = null;
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      hiddenAt = Date.now();
    } else if (hiddenAt && Date.now() - hiddenAt >= AUTO_REFRESH_MS) {
      snapshotDrafts();
      window.location.reload();
    }
  });

  return timer;
}

export const restoreAfterAutoRefresh = () => {
  try {
    const raw = sessionStorage.getItem("worshiptool:autorefresh");
    if (!raw) return;
    sessionStorage.removeItem("worshiptool:autorefresh");
    const { at, data } = JSON.parse(raw);
    // Only honour a recent snapshot (the reload itself is instant).
    if (Date.now() - at > 60_000) return;
    Object.entries(data || {}).forEach(([k, v]) => {
      if (typeof v === "string") localStorage.setItem(k, v);
    });
  } catch {
    // ignore
  }
};

restoreAfterAutoRefresh();
startAutoRefresh();

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
