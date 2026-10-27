import { useEffect, useRef, useState, useCallback } from "react";
import { ExternalLink, RefreshCw, ArrowLeft, ArrowRight } from "lucide-react";

const PROXY = "/api/bible-proxy/";
const HOME = "bibles/khmer/";

/**
 * Same-origin iframe reader for wordproject.org (Khmer Bible).
 * Selections inside the proxied page are forwarded via postMessage and
 * surfaced through onSelection(text).
 */
export default function BibleReader({ onSelection, onDeselection }) {
  const [url, setUrl] = useState(PROXY + HOME);
  const [back, setBack] = useState([]);
  const [forward, setForward] = useState([]);
  const [loading, setLoading] = useState(false);
  const frameRef = useRef(null);

  // Load the proxied page and inject a selection bridge into it.
  // The bridge posts {source:'wordproject-proxy', type:'selection', text}
  // and {source:'wordproject-proxy', type:'deselection'} messages to the parent window.
  useEffect(() => {
    const handler = (e) => {
      const d = e.data;
      if (d && d.source === "wordproject-proxy") {
        if (d.type === "selection") {
          onSelection?.(d.text);
        } else if (d.type === "deselection") {
          onDeselection?.();
        }
      }
    };
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler);
  }, [onSelection, onDeselection]);

  const navigate = useCallback(
    (path) => {
      const next = path.startsWith("http")
        ? path.replace("https://www.wordproject.org/", PROXY)
        : path.startsWith(PROXY)
          ? path
          : PROXY + path;
      setBack((b) => [...b, url]);
      setForward([]);
      setUrl(next);
    },
    [url],
  );

  const handleLoad = () => {
    setLoading(false);
    try {
      const f = frameRef.current;
      if (!f) return;
      // Track in-frame navigation for the address bar / history.
      const loc = f.contentWindow.location;
      const href = loc.pathname.replace(PROXY, "") + loc.search;
      if (href !== url.replace(PROXY, "")) {
        setBack((b) => [...b, url]);
        setForward([]);
        setUrl(PROXY + href);
      }
    } catch {
      /* same-origin thanks to the proxy — ignore transient errors */
    }
  };

  const goBack = () => {
    if (!back.length) return;
    const prev = back[back.length - 1];
    setBack((b) => b.slice(0, -1));
    setForward((f) => [url, ...f]);
    setUrl(prev);
  };

  const goForward = () => {
    if (!forward.length) return;
    const next = forward[0];
    setForward((f) => f.slice(1));
    setBack((b) => [...b, url]);
    setUrl(next);
  };

  return (
    <div className="flex flex-col h-full overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-card">
      {/* Reader toolbar */}
      <div className="flex items-center gap-2 px-3 py-2 border-b border-slate-100 bg-slate-50/80">
        <button
          onClick={goBack}
          disabled={!back.length}
          title="Back"
          className="p-1.5 rounded-lg hover:bg-slate-200/70 disabled:opacity-30 text-slate-500 transition"
        >
          <ArrowLeft size={15} />
        </button>
        <button
          onClick={goForward}
          disabled={!forward.length}
          title="Forward"
          className="p-1.5 rounded-lg hover:bg-slate-200/70 disabled:opacity-30 text-slate-500 transition"
        >
          <ArrowRight size={15} />
        </button>
        <button
          onClick={() => navigate(HOME)}
          title="Home — Khmer Bible"
          className="p-1.5 rounded-lg hover:bg-slate-200/70 text-slate-500 transition"
        >
          <RefreshCw size={14} />
        </button>

        <div className="flex-1 flex items-center gap-2 bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-xs text-slate-500 font-mono truncate">
          <span className="text-brand-500">🔒</span>
          <span className="truncate">
            wordproject.org/{url.replace(PROXY, "")}
          </span>
        </div>

        <a
          href={`https://www.wordproject.org/${url.replace(PROXY, "")}`}
          target="_blank"
          rel="noreferrer"
          title="Open in new tab"
          className="p-1.5 rounded-lg hover:bg-slate-200/70 text-slate-500 transition"
        >
          <ExternalLink size={14} />
        </a>
      </div>

      {/* iframe */}
      <div className="relative flex-1 min-h-0 bg-white">
        {loading && (
          <div className="absolute inset-0 bg-white/70 flex items-center justify-center z-10">
            <RefreshCw size={22} className="animate-spin text-brand-500" />
          </div>
        )}
        <iframe
          ref={frameRef}
          src={url}
          title="Khmer Bible Reader"
          className="w-full h-full border-0"
          onLoad={handleLoad}
        />
      </div>

      {/* Hint bar */}
      <div className="px-3 py-1.5 border-t border-slate-100 bg-amber-50 text-[11px] text-amber-700">
        💡 Select any verse text above — it will be added to the Bible Text
        editor automatically.
      </div>
    </div>
  );
}
