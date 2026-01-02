/* app.js
   Robust, dependency-free tab + button wiring.
   Works with common patterns:
   - Tabs: [role="tab"] inside [role="tablist"] OR elements with [data-tab-target] OR links with href="#panelId"
   - Panels: [role="tabpanel"] OR elements with id="panelId" OR [data-tab-panel="name"]
   - Buttons: elements with [data-action="..."] (see ACTIONS map below)
*/

(() => {
  "use strict";

  // -----------------------------
  // Utilities
  // -----------------------------
  const qs = (sel, root = document) => root.querySelector(sel);
  const qsa = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const isEl = (x) => x && x.nodeType === 1;
  const has = (obj, k) => Object.prototype.hasOwnProperty.call(obj, k);

  const safeJSONParse = (s, fallback = null) => {
    try {
      return JSON.parse(s);
    } catch {
      return fallback;
    }
  };

  const downloadText = (filename, text, mime = "text/plain;charset=utf-8") => {
    const blob = new Blob([text], { type: mime });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
  };

  const copyToClipboard = async (text) => {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
    // Fallback
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    ta.remove();
    return ok;
  };

  const closestWithAttr = (el, attr) => (isEl(el) ? el.closest(`[${attr}]`) : null);

  const getTargetSelectorFromTrigger = (trigger) => {
    if (!isEl(trigger)) return null;

    // Preferred: data-tab-target="#id" or "id"
    const dt = trigger.getAttribute("data-tab-target") || trigger.getAttribute("data-target");
    if (dt) return dt.startsWith("#") ? dt : `#${dt}`;

    // ARIA: aria-controls="panelId"
    const ac = trigger.getAttribute("aria-controls");
    if (ac) return `#${ac}`;

    // Links: href="#panelId"
    const href = trigger.getAttribute("href");
    if (href && href.startsWith("#") && href.length > 1) return href;

    // data-tab="name" with panel [data-tab-panel="name"]
    const name = trigger.getAttribute("data-tab");
    if (name) return `[data-tab-panel="${CSS.escape(name)}"]`;

    return null;
  };

  const showToast = (() => {
    let container = null;

    const ensure = () => {
      if (container) return container;
      container = document.createElement("div");
      container.id = "app-toast-container";
      container.style.position = "fixed";
      container.style.right = "16px";
      container.style.bottom = "16px";
      container.style.zIndex = "9999";
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.gap = "10px";
      document.body.appendChild(container);
      return container;
    };

    return (message, type = "info", timeoutMs = 2200) => {
      const c = ensure();
      const t = document.createElement("div");
      t.className = `app-toast app-toast--${type}`;
      t.textContent = message;

      // Minimal styling (won't override your CSS if you already have these classes)
      t.style.padding = "10px 12px";
      t.style.borderRadius = "12px";
      t.style.boxShadow = "0 10px 30px rgba(0,0,0,0.15)";
      t.style.background = type === "error" ? "rgba(220, 38, 38, 0.95)"
        : type === "success" ? "rgba(22, 163, 74, 0.95)"
        : "rgba(15, 23, 42, 0.95)";
      t.style.color = "white";
      t.style.fontSize = "13px";
      t.style.maxWidth = "360px";
      t.style.lineHeight = "1.35";

      c.appendChild(t);
      setTimeout(() => {
        t.style.opacity = "0";
        t.style.transform = "translateY(6px)";
        t.style.transition = "opacity 200ms ease, transform 200ms ease";
        setTimeout(() => t.remove(), 220);
      }, timeoutMs);
    };
  })();

  // -----------------------------
  // Tabs
  // -----------------------------
  const Tabs = (() => {
    const ACTIVE_CLASSES = ["active", "is-active", "show"];
    const PANEL_ACTIVE_CLASSES = ["active", "is-active", "show"];

    const setActive = (el, active, classes) => {
      if (!isEl(el)) return;
      classes.forEach((c) => el.classList.toggle(c, !!active));
    };

    const markTabState = (tab, isActive) => {
      if (!isEl(tab)) return;

      // Remove accidental disabling that makes tabs "inactive"
      if (tab.hasAttribute("disabled")) tab.removeAttribute("disabled");
      if (tab.getAttribute("aria-disabled") === "true") tab.setAttribute("aria-disabled", "false");
      tab.classList.remove("disabled");

      tab.setAttribute("aria-selected", isActive ? "true" : "false");
      tab.setAttribute("tabindex", isActive ? "0" : "-1");
      setActive(tab, isActive, ACTIVE_CLASSES);
    };

    const markPanelState = (panel, isActive) => {
      if (!isEl(panel)) return;
      panel.hidden = !isActive;
      setActive(panel, isActive, PANEL_ACTIVE_CLASSES);
      if (isActive) panel.removeAttribute("aria-hidden");
      else panel.setAttribute("aria-hidden", "true");
    };

    const getTabGroupRoot = (tab) => {
      // Prefer explicit group container
      const group = tab.closest("[data-tab-group]") || tab.closest("[role='tablist']") || tab.closest(".tabs") || tab.parentElement;
      return group || document;
    };

    const getTabsInGroup = (groupRoot) => {
      const byRole = qsa("[role='tab']", groupRoot);
      if (byRole.length) return byRole;
      // fallback common patterns
      const byData = qsa("[data-tab-target],[data-target],[data-tab]", groupRoot);
      if (byData.length) return byData;
      const byLinks = qsa("a[href^='#']", groupRoot).filter((a) => a.getAttribute("href").length > 1);
      return byLinks;
    };

    const resolvePanel = (tab) => {
      const sel = getTargetSelectorFromTrigger(tab);
      if (!sel) return null;

      // If selector is [data-tab-panel="x"], use it as-is
      if (sel.startsWith("[data-tab-panel=")) return qs(sel);

      // If selector is #id, try direct id first
      const p = qs(sel);
      if (p) return p;

      // Some frameworks prefix panel ids or use data attributes
      const id = sel.startsWith("#") ? sel.slice(1) : sel;
      return qs(`[data-tab-panel="${CSS.escape(id)}"]`) || qs(`[role="tabpanel"][id="${CSS.escape(id)}"]`);
    };

    const getAllPanelsForGroup = (tabs) => {
      const panels = [];
      tabs.forEach((t) => {
        const p = resolvePanel(t);
        if (p && !panels.includes(p)) panels.push(p);
      });
      return panels;
    };

    const activate = (tab, opts = {}) => {
      if (!isEl(tab)) return;

      const groupRoot = getTabGroupRoot(tab);
      const tabs = getTabsInGroup(groupRoot);
      const panels = getAllPanelsForGroup(tabs);

      const targetPanel = resolvePanel(tab);

      // If we can't resolve a panel, still mark tab active for UI consistency
      tabs.forEach((t) => markTabState(t, t === tab));
      panels.forEach((p) => markPanelState(p, p === targetPanel));

      if (opts.updateHash) {
        const sel = getTargetSelectorFromTrigger(tab);
        if (sel && sel.startsWith("#")) {
          history.replaceState(null, "", sel);
        }
      }

      if (opts.focusPanel && targetPanel) {
        targetPanel.setAttribute("tabindex", "-1");
        targetPanel.focus({ preventScroll: true });
      }
    };

    const keyboardNav = (e, tab) => {
      const key = e.key;
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(key)) return;

      const groupRoot = getTabGroupRoot(tab);
      const tabs = getTabsInGroup(groupRoot).filter((t) => !t.hasAttribute("disabled") && t.getAttribute("aria-disabled") !== "true");
      if (!tabs.length) return;

      const idx = tabs.indexOf(tab);
      if (idx < 0) return;

      e.preventDefault();

      let nextIdx = idx;
      if (key === "ArrowLeft") nextIdx = (idx - 1 + tabs.length) % tabs.length;
      if (key === "ArrowRight") nextIdx = (idx + 1) % tabs.length;
      if (key === "Home") nextIdx = 0;
      if (key === "End") nextIdx = tabs.length - 1;

      const next = tabs[nextIdx];
      next.focus();
      activate(next, { updateHash: true });
    };

    const initGroup = (groupRoot) => {
      const tabs = getTabsInGroup(groupRoot);
      if (!tabs.length) return;

      // Ensure roles
      const tablist = groupRoot.matches("[role='tablist']") ? groupRoot : groupRoot.querySelector("[role='tablist']");
      if (tablist && !tablist.hasAttribute("role")) tablist.setAttribute("role", "tablist");

      tabs.forEach((t) => {
        if (!t.hasAttribute("role")) t.setAttribute("role", "tab");
        t.setAttribute("aria-selected", "false");
        t.setAttribute("tabindex", "-1");

        const p = resolvePanel(t);
        const sel = getTargetSelectorFromTrigger(t);
        if (p) {
          if (!p.hasAttribute("role")) p.setAttribute("role", "tabpanel");
          if (!p.id && sel && sel.startsWith("#")) p.id = sel.slice(1);
          if (p.id && !t.getAttribute("aria-controls")) t.setAttribute("aria-controls", p.id);
          if (!p.getAttribute("aria-labelledby") && t.id) p.setAttribute("aria-labelledby", t.id);
        }
      });

      // Determine initial tab: hash -> existing active -> first
      const hash = window.location.hash;
      let initial = null;

      if (hash) {
        initial = tabs.find((t) => getTargetSelectorFromTrigger(t) === hash) || null;
        if (!initial) {
          // If panels exist with this hash, find the tab controlling it
          const id = hash.slice(1);
          initial = tabs.find((t) => (t.getAttribute("aria-controls") || "") === id) || null;
        }
      }

      if (!initial) {
        initial = tabs.find((t) => t.classList.contains("active") || t.getAttribute("aria-selected") === "true") || null;
      }

      if (!initial) initial = tabs[0];

      activate(initial, { updateHash: false });

      // Bind events
      tabs.forEach((t) => {
        t.addEventListener("click", (e) => {
          // Prevent default for anchor-based tabs (so we control panel behaviour + hash updates)
          if (t.tagName === "A") e.preventDefault();
          activate(t, { updateHash: true });
        });
        t.addEventListener("keydown", (e) => keyboardNav(e, t));
      });
    };

    const init = () => {
      // Prefer explicit groups
      const explicitGroups = qsa("[data-tab-group]");
      if (explicitGroups.length) {
        explicitGroups.forEach(initGroup);
      } else {
        // Fallback: each tablist is a group
        const tablists = qsa("[role='tablist']");
        if (tablists.length) {
          tablists.forEach(initGroup);
        } else {
          // Last fallback: any container with likely tab triggers
          const likely = qsa(".tabs, .tablist, .nav-tabs, .tab-buttons");
          if (likely.length) likely.forEach(initGroup);
          else initGroup(document);
        }
      }

      // If hash changes externally, activate matching tab
      window.addEventListener("hashchange", () => {
        const hash = window.location.hash;
        if (!hash) return;
        const allTabs = qsa("[role='tab'],[data-tab-target],[data-target],a[href^='#']");
        const match = allTabs.find((t) => getTargetSelectorFromTrigger(t) === hash) || null;
        if (match) activate(match, { updateHash: false });
      });
    };

    return { init, activate };
  })();

  // -----------------------------
  // Actions (Buttons)
  // -----------------------------
  const Actions = (() => {
    const getScopeRoot = (el) => {
      const sel = el.getAttribute("data-scope");
      if (sel) return qs(sel) || document;
      const closestScope = el.closest("[data-scope-root]");
      return closestScope || document;
    };

    const getTextFromTarget = (el) => {
      const targetSel = el.getAttribute("data-copy-target") || el.getAttribute("data-target") || el.getAttribute("data-from");
      const scope = getScopeRoot(el);

      if (targetSel) {
        const t = qs(targetSel, scope) || qs(targetSel);
        if (t) {
          if ("value" in t) return String(t.value ?? "");
          return String(t.textContent ?? "").trim();
        }
      }

      // fallback: nearest textarea / pre
      const ta = el.closest(".panel, [role='tabpanel'], .tab-panel")?.querySelector("textarea, pre, code");
      if (ta) {
        if ("value" in ta) return String(ta.value ?? "");
        return String(ta.textContent ?? "").trim();
      }

      return "";
    };

    const toCSV = (rows) => {
      if (!Array.isArray(rows) || !rows.length) return "";
      // If it's an array of objects, use keys union
      if (rows.every((r) => r && typeof r === "object" && !Array.isArray(r))) {
        const keys = Array.from(rows.reduce((set, r) => {
          Object.keys(r).forEach((k) => set.add(k));
          return set;
        }, new Set()));
        const esc = (v) => {
          const s = v == null ? "" : String(v);
          if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
          return s;
        };
        const header = keys.map(esc).join(",");
        const body = rows.map((r) => keys.map((k) => esc(r[k])).join(",")).join("\n");
        return `${header}\n${body}`;
      }
      // Otherwise, stringify line by line
      return rows.map((r) => (Array.isArray(r) ? r.join(",") : String(r))).join("\n");
    };

    const ACTIONS = {
      // Copies text from target into clipboard
      copy: async (btn) => {
        const text = getTextFromTarget(btn);
        if (!text) {
          showToast("Nothing to copy.", "error");
          return;
        }
        try {
          await copyToClipboard(text);
          showToast("Copied.", "success");
        } catch {
          showToast("Copy failed.", "error");
        }
      },

      // Downloads JSON from target (or from data-json literal)
      "download-json": (btn) => {
        const literal = btn.getAttribute("data-json");
        let obj = literal ? safeJSONParse(literal, null) : null;

        if (!obj) {
          const txt = getTextFromTarget(btn);
          obj = safeJSONParse(txt, null);
        }

        if (!obj) {
          showToast("No valid JSON found.", "error");
          return;
        }

        const name = btn.getAttribute("data-filename") || "export.json";
        downloadText(name, JSON.stringify(obj, null, 2), "application/json;charset=utf-8");
        showToast("JSON downloaded.", "success");
      },

      // Downloads CSV from target (expects JSON array or table-like text)
      "download-csv": (btn) => {
        const txt = getTextFromTarget(btn);
        let csv = "";

        // If target is already CSV-ish, keep it
        if (txt && txt.includes(",") && txt.includes("\n") && !txt.trim().startsWith("{") && !txt.trim().startsWith("[")) {
          csv = txt;
        } else {
          const parsed = safeJSONParse(txt, null);
          if (parsed) csv = toCSV(parsed);
        }

        if (!csv) {
          showToast("No CSV-compatible content found.", "error");
          return;
        }

        const name = btn.getAttribute("data-filename") || "export.csv";
        downloadText(name, csv, "text/csv;charset=utf-8");
        showToast("CSV downloaded.", "success");
      },

      // Opens Copilot (or any URL) in a new tab
      "open-copilot": (btn) => {
        const url = btn.getAttribute("data-url") || "https://copilot.microsoft.com/";
        window.open(url, "_blank", "noopener,noreferrer");
      },

      // Scroll to a selector or id
      "scroll-to": (btn) => {
        const sel = btn.getAttribute("data-scroll-target") || btn.getAttribute("data-target");
        if (!sel) return;
        const t = qs(sel) || qs(`#${CSS.escape(sel)}`);
        if (t) t.scrollIntoView({ behavior: "smooth", block: "start" });
      },

      // Toggle a class on a target selector (default: 'is-open')
      toggle: (btn) => {
        const sel = btn.getAttribute("data-toggle-target") || btn.getAttribute("data-target");
        const cls = btn.getAttribute("data-toggle-class") || "is-open";
        if (!sel) return;
        const t = qs(sel) || qs(`#${CSS.escape(sel)}`);
        if (t) t.classList.toggle(cls);
      },

      // Reset forms within a scope
      reset: (btn) => {
        const scopeSel = btn.getAttribute("data-reset-scope");
        const root = scopeSel ? (qs(scopeSel) || document) : getScopeRoot(btn);
        qsa("form", root).forEach((f) => f.reset());
        showToast("Reset complete.", "success");
      },

      // Activate a specific tab programmatically: data-activate-tab="#panelId" or "panelId"
      "activate-tab": (btn) => {
        const target = btn.getAttribute("data-activate-tab") || btn.getAttribute("data-target");
        if (!target) return;
        const sel = target.startsWith("#") ? target : `#${target}`;
        const allTabs = qsa("[role='tab'],[data-tab-target],[data-target],a[href^='#']");
        const match = allTabs.find((t) => getTargetSelectorFromTrigger(t) === sel) || null;
        if (match) Tabs.activate(match, { updateHash: true, focusPanel: true });
      }
    };

    const init = () => {
      // Delegated click handling for all buttons/links with data-action
      document.addEventListener("click", async (e) => {
        const trigger = closestWithAttr(e.target, "data-action");
        if (!trigger) return;

        // If it's a link used as a button, prevent navigation unless explicitly allowed
        const allowNav = trigger.getAttribute("data-allow-nav") === "true";
        if (trigger.tagName === "A" && !allowNav) e.preventDefault();

        const action = (trigger.getAttribute("data-action") || "").trim().toLowerCase();
        if (!action) return;

        const fn = ACTIONS[action];
        if (!fn) {
          showToast(`Unknown action: ${action}`, "error");
          return;
        }

        try {
          const out = fn(trigger);
          if (out && typeof out.then === "function") await out;
        } catch (err) {
          console.error(err);
          showToast("Action failed.", "error");
        }
      });
    };

    return { init };
  })();

  // -----------------------------
  // Generic “make UI clickable” hardening
  // -----------------------------
  const hardenUI = () => {
    // Some CSS frameworks render tab links with pointer-events: none via disabled classes.
    // This re-enables obvious candidates unless explicitly locked.
    qsa(".disabled,[aria-disabled='true']").forEach((el) => {
      if (el.getAttribute("data-locked") === "true") return;
      if (el.matches("[role='tab'],[data-tab-target],[data-tab],a[href^='#'],button")) {
        el.classList.remove("disabled");
        if (el.getAttribute("aria-disabled") === "true") el.setAttribute("aria-disabled", "false");
        if (el.hasAttribute("disabled")) el.removeAttribute("disabled");
      }
    });

    // Ensure any element styled as a button but implemented as div is clickable via keyboard
    qsa("[data-action]").forEach((el) => {
      if (el.tagName === "DIV" || el.tagName === "SPAN") {
        if (!el.hasAttribute("role")) el.setAttribute("role", "button");
        if (!el.hasAttribute("tabindex")) el.setAttribute("tabindex", "0");
        el.addEventListener("keydown", (e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            el.click();
          }
        });
      }
    });
  };

  // -----------------------------
  // Boot
  // -----------------------------
  const boot = () => {
    Tabs.init();
    Actions.init();
    hardenUI();

    // Re-harden after dynamic DOM changes (if your app injects content)
    const obs = new MutationObserver(() => hardenUI());
    obs.observe(document.body, { childList: true, subtree: true });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
