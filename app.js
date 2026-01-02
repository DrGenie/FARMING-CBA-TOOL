/* app.js — Commercial-grade, dependency-free UI controller
   Drop-in script that makes tabs, panels, buttons, modals, drawers, tooltips, accordions, toasts,
   exports (JSON/CSV), clipboard copy, hash routing, persistence, and accessibility work reliably.

   Markup conventions supported (any subset works):
   - Tabs:
     - [role="tab"] with aria-controls="panelId" inside [role="tablist"]
     - or [data-tab-target="#panelId"|"panelId"]
     - or <a href="#panelId"> as a tab trigger
     - Optional grouping via [data-tab-group="name"] on a container
   - Panels:
     - [role="tabpanel"] with id="panelId"
     - or [data-tab-panel="name"] paired with trigger [data-tab="name"]
   - Actions/buttons:
     - Any clickable element with [data-action="..."] (see ACTIONS registry)
   - Modals:
     - [data-modal="id"] on modal root; open with [data-action="modal-open"][data-modal-id="id"]
     - close with [data-action="modal-close"] inside modal, or click overlay, or Escape
   - Drawers:
     - [data-drawer="id"] on drawer root; open with [data-action="drawer-open"][data-drawer-id="id"]
   - Tooltips:
     - [data-tooltip="Text"] on any element, or [data-tooltip-target="#id"] to show rich tooltip content
   - Accordions:
     - [data-accordion] container
     - [data-acc-trigger] button + [data-acc-panel] panel (siblings or within same item)
   - Persistence:
     - Add [data-persist="key"] to inputs/selects/textareas to auto-save to localStorage
   - App config:
     - Optional <script type="application/json" id="app-config">{...}</script>

   Notes:
   - This file cannot “preserve” unknown bespoke behaviours from a different app.js without seeing it.
     It is built to be tolerant and broadly compatible so your existing IDs/classes can remain unchanged.
*/

(() => {
  "use strict";

  // ------------------------------------------------------------
  // Core utilities
  // ------------------------------------------------------------
  const App = {};
  const w = window;
  const d = document;

  const qs = (sel, root = d) => root.querySelector(sel);
  const qsa = (sel, root = d) => Array.from(root.querySelectorAll(sel));
  const isEl = (x) => x && x.nodeType === 1;
  const now = () => Date.now();

  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const debounce = (fn, ms = 180) => {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  };

  const safeJSON = {
    parse: (s, fallback = null) => {
      try {
        return JSON.parse(s);
      } catch {
        return fallback;
      }
    },
    stringify: (v, fallback = "null") => {
      try {
        return JSON.stringify(v);
      } catch {
        return fallback;
      }
    }
  };

  const escapeCSV = (v) => {
    const s = v == null ? "" : String(v);
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };

  const toCSV = (rows) => {
    if (!Array.isArray(rows) || rows.length === 0) return "";
    if (rows.every((r) => r && typeof r === "object" && !Array.isArray(r))) {
      const keys = Array.from(rows.reduce((set, r) => {
        Object.keys(r).forEach((k) => set.add(k));
        return set;
      }, new Set()));
      const header = keys.map(escapeCSV).join(",");
      const body = rows.map((r) => keys.map((k) => escapeCSV(r[k])).join(",")).join("\n");
      return `${header}\n${body}`;
    }
    return rows.map((r) => (Array.isArray(r) ? r.map(escapeCSV).join(",") : escapeCSV(r))).join("\n");
  };

  const downloadText = (filename, text, mime = "text/plain;charset=utf-8") => {
    const blob = new Blob([text], { type: mime });
    const a = d.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    d.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
  };

  const copyToClipboard = async (text) => {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
    const ta = d.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    d.body.appendChild(ta);
    ta.select();
    const ok = d.execCommand("copy");
    ta.remove();
    return ok;
  };

  const closestAttr = (el, attr) => (isEl(el) ? el.closest(`[${attr}]`) : null);

  // ------------------------------------------------------------
  // Lightweight event bus
  // ------------------------------------------------------------
  App.bus = (() => {
    const map = new Map();
    return {
      on: (evt, fn) => {
        if (!map.has(evt)) map.set(evt, new Set());
        map.get(evt).add(fn);
        return () => map.get(evt)?.delete(fn);
      },
      emit: (evt, payload) => {
        map.get(evt)?.forEach((fn) => {
          try {
            fn(payload);
          } catch (e) {
            console.error(`[bus] handler error for ${evt}`, e);
          }
        });
      }
    };
  })();

  // ------------------------------------------------------------
  // Config + logging
  // ------------------------------------------------------------
  App.config = (() => {
    const defaultConfig = {
      storagePrefix: "tool:",
      toastDuration: 2400,
      enableHashRouting: true,
      autoInitTooltips: true,
      autoInitAccordions: true,
      autoInitTabs: true,
      autoInitPersistence: true,
      strictAriaTabs: true
    };
    const el = qs("#app-config");
    const fromDOM = el ? safeJSON.parse(el.textContent || "", {}) : {};
    return Object.freeze({ ...defaultConfig, ...fromDOM });
  })();

  App.log = (() => {
    const tag = "[tool]";
    return {
      info: (...a) => console.log(tag, ...a),
      warn: (...a) => console.warn(tag, ...a),
      error: (...a) => console.error(tag, ...a)
    };
  })();

  // ------------------------------------------------------------
  // Store (simple reactive state + persistence)
  // ------------------------------------------------------------
  App.store = (() => {
    const prefix = App.config.storagePrefix;
    const listeners = new Set();

    const load = (key, fallback = null) => {
      const s = localStorage.getItem(prefix + key);
      if (s == null) return fallback;
      return safeJSON.parse(s, fallback);
    };

    const save = (key, value) => {
      localStorage.setItem(prefix + key, safeJSON.stringify(value, "null"));
    };

    const state = new Proxy(
      {
        ui: {
          theme: load("theme", "light")
        }
      },
      {
        set(target, prop, value) {
          target[prop] = value;
          listeners.forEach((fn) => {
            try {
              fn(target, prop, value);
            } catch (e) {
              App.log.error("store listener error", e);
            }
          });
          return true;
        }
      }
    );

    return {
      state,
      subscribe: (fn) => {
        listeners.add(fn);
        return () => listeners.delete(fn);
      },
      load,
      save
    };
  })();

  // ------------------------------------------------------------
  // Toasts
  // ------------------------------------------------------------
  App.toast = (() => {
    let container = null;

    const ensure = () => {
      if (container) return container;
      container = d.createElement("div");
      container.id = "app-toast-container";
      container.style.position = "fixed";
      container.style.right = "16px";
      container.style.bottom = "16px";
      container.style.zIndex = "10000";
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.gap = "10px";
      d.body.appendChild(container);
      return container;
    };

    const show = (message, type = "info", duration = App.config.toastDuration) => {
      const c = ensure();
      const t = d.createElement("div");
      t.className = `app-toast app-toast--${type}`;
      t.textContent = message;

      // Minimal styling (won’t override if your CSS defines these classes)
      t.style.padding = "10px 12px";
      t.style.borderRadius = "14px";
      t.style.boxShadow = "0 12px 40px rgba(0,0,0,0.18)";
      t.style.background =
        type === "error"
          ? "rgba(220, 38, 38, 0.96)"
          : type === "success"
          ? "rgba(22, 163, 74, 0.96)"
          : type === "warning"
          ? "rgba(245, 158, 11, 0.96)"
          : "rgba(15, 23, 42, 0.96)";
      t.style.color = "white";
      t.style.fontSize = "13px";
      t.style.lineHeight = "1.35";
      t.style.maxWidth = "380px";
      t.style.transform = "translateY(0)";
      t.style.opacity = "1";

      c.appendChild(t);

      const close = () => {
        t.style.transition = "opacity 180ms ease, transform 180ms ease";
        t.style.opacity = "0";
        t.style.transform = "translateY(6px)";
        setTimeout(() => t.remove(), 200);
      };

      const timer = setTimeout(close, duration);
      t.addEventListener("click", () => {
        clearTimeout(timer);
        close();
      });
    };

    return { show };
  })();

  // ------------------------------------------------------------
  // Focus trap (modals/drawers)
  // ------------------------------------------------------------
  App.focusTrap = (() => {
    const focusableSel = [
      "a[href]:not([tabindex='-1'])",
      "button:not([disabled]):not([tabindex='-1'])",
      "textarea:not([disabled]):not([tabindex='-1'])",
      "input:not([disabled]):not([type='hidden']):not([tabindex='-1'])",
      "select:not([disabled]):not([tabindex='-1'])",
      "[tabindex]:not([tabindex='-1'])"
    ].join(",");

    const trap = (root) => {
      if (!isEl(root)) return () => {};
      const onKeyDown = (e) => {
        if (e.key !== "Tab") return;
        const f = qsa(focusableSel, root).filter((el) => el.offsetParent !== null);
        if (!f.length) return;
        const first = f[0];
        const last = f[f.length - 1];
        if (e.shiftKey && d.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && d.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      };
      root.addEventListener("keydown", onKeyDown);
      return () => root.removeEventListener("keydown", onKeyDown);
    };

    return { trap };
  })();

  // ------------------------------------------------------------
  // Tabs (robust: groups, ARIA, keyboard, hash routing)
  // ------------------------------------------------------------
  App.tabs = (() => {
    const ACTIVE_TAB_CLASSES = ["active", "is-active", "show"];
    const ACTIVE_PANEL_CLASSES = ["active", "is-active", "show"];

    const setClasses = (el, on, classes) => {
      if (!isEl(el)) return;
      classes.forEach((c) => el.classList.toggle(c, !!on));
    };

    const getTargetSelector = (trigger) => {
      if (!isEl(trigger)) return null;

      const dt = trigger.getAttribute("data-tab-target") || trigger.getAttribute("data-target");
      if (dt) return dt.startsWith("#") ? dt : `#${dt}`;

      const ac = trigger.getAttribute("aria-controls");
      if (ac) return `#${ac}`;

      const href = trigger.getAttribute("href");
      if (href && href.startsWith("#") && href.length > 1) return href;

      const name = trigger.getAttribute("data-tab");
      if (name) return `[data-tab-panel="${CSS.escape(name)}"]`;

      return null;
    };

    const resolvePanel = (trigger) => {
      const sel = getTargetSelector(trigger);
      if (!sel) return null;
      const p = qs(sel);
      if (p) return p;

      // fallback: if sel is #id but not found, try data-tab-panel="id"
      if (sel.startsWith("#")) {
        const id = sel.slice(1);
        return qs(`[data-tab-panel="${CSS.escape(id)}"]`);
      }
      return null;
    };

    const getGroupRoot = (trigger) => {
      return (
        trigger.closest("[data-tab-group]") ||
        trigger.closest("[role='tablist']") ||
        trigger.closest(".tabs") ||
        trigger.closest(".tablist") ||
        trigger.parentElement ||
        d
      );
    };

    const getTabsInGroup = (groupRoot) => {
      const byRole = qsa("[role='tab']", groupRoot);
      if (byRole.length) return byRole;
      const byData = qsa("[data-tab-target],[data-target],[data-tab]", groupRoot);
      if (byData.length) return byData;
      return qsa("a[href^='#']", groupRoot).filter((a) => a.getAttribute("href").length > 1);
    };

    const getPanelsForGroup = (tabs) => {
      const panels = [];
      tabs.forEach((t) => {
        const p = resolvePanel(t);
        if (p && !panels.includes(p)) panels.push(p);
      });
      return panels;
    };

    const hardenTrigger = (t) => {
      if (!isEl(t)) return;
      // remove disabling that breaks UX, unless explicitly locked
      if (t.getAttribute("data-locked") === "true") return;
      if (t.hasAttribute("disabled")) t.removeAttribute("disabled");
      if (t.getAttribute("aria-disabled") === "true") t.setAttribute("aria-disabled", "false");
      t.classList.remove("disabled");
      t.style.pointerEvents = "";
    };

    const markActive = (tab, isActive) => {
      if (!isEl(tab)) return;
      hardenTrigger(tab);

      if (App.config.strictAriaTabs) {
        tab.setAttribute("aria-selected", isActive ? "true" : "false");
        tab.setAttribute("tabindex", isActive ? "0" : "-1");
        if (!tab.hasAttribute("role")) tab.setAttribute("role", "tab");
      }
      setClasses(tab, isActive, ACTIVE_TAB_CLASSES);
    };

    const markPanel = (panel, isActive) => {
      if (!isEl(panel)) return;
      panel.hidden = !isActive;
      panel.setAttribute("aria-hidden", isActive ? "false" : "true");
      setClasses(panel, isActive, ACTIVE_PANEL_CLASSES);
    };

    const activate = (tab, opts = {}) => {
      if (!isEl(tab)) return;
      const groupRoot = getGroupRoot(tab);
      const tabs = getTabsInGroup(groupRoot);
      const panels = getPanelsForGroup(tabs);
      const targetPanel = resolvePanel(tab);

      tabs.forEach((t) => markActive(t, t === tab));
      panels.forEach((p) => markPanel(p, p === targetPanel));

      if (opts.updateHash && App.config.enableHashRouting) {
        const sel = getTargetSelector(tab);
        if (sel && sel.startsWith("#")) history.replaceState(null, "", sel);
      }

      if (opts.focusPanel && targetPanel) {
        targetPanel.setAttribute("tabindex", "-1");
        targetPanel.focus({ preventScroll: true });
      }

      App.bus.emit("tab:activated", { tab, panel: targetPanel });
    };

    const keyboardNav = (e, tab) => {
      const key = e.key;
      if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(key)) return;

      const groupRoot = getGroupRoot(tab);
      const tabs = getTabsInGroup(groupRoot).filter((t) => {
        const dis = t.hasAttribute("disabled") || t.getAttribute("aria-disabled") === "true" || t.classList.contains("disabled");
        return !dis;
      });

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

      // establish ARIA relationships where possible
      tabs.forEach((t, i) => {
        hardenTrigger(t);
        if (App.config.strictAriaTabs) {
          if (!t.hasAttribute("role")) t.setAttribute("role", "tab");
          if (!t.id) t.id = `tab-${Math.random().toString(36).slice(2)}-${i}`;
        }
        const p = resolvePanel(t);
        if (p) {
          if (!p.hasAttribute("role")) p.setAttribute("role", "tabpanel");
          const sel = getTargetSelector(t);
          if (!p.id && sel && sel.startsWith("#")) p.id = sel.slice(1);
          if (p.id && !t.getAttribute("aria-controls")) t.setAttribute("aria-controls", p.id);
          if (!p.getAttribute("aria-labelledby")) p.setAttribute("aria-labelledby", t.id);
        }
      });

      // initial tab: hash > active class > first
      let initial = null;
      const hash = w.location.hash;
      if (App.config.enableHashRouting && hash) {
        initial = tabs.find((t) => getTargetSelector(t) === hash) || null;
        if (!initial) {
          const id = hash.slice(1);
          initial = tabs.find((t) => (t.getAttribute("aria-controls") || "") === id) || null;
        }
      }
      if (!initial) {
        initial =
          tabs.find((t) => t.classList.contains("active") || t.classList.contains("is-active") || t.getAttribute("aria-selected") === "true") ||
          null;
      }
      if (!initial) initial = tabs[0];

      activate(initial, { updateHash: false });

      // bind events
      tabs.forEach((t) => {
        t.addEventListener("click", (e) => {
          if (t.tagName === "A") e.preventDefault();
          activate(t, { updateHash: true });
        });
        t.addEventListener("keydown", (e) => keyboardNav(e, t));
      });
    };

    const init = () => {
      if (!App.config.autoInitTabs) return;
      const groups = qsa("[data-tab-group]");
      if (groups.length) groups.forEach(initGroup);
      else {
        const tablists = qsa("[role='tablist']");
        if (tablists.length) tablists.forEach(initGroup);
        else initGroup(d);
      }

      if (App.config.enableHashRouting) {
        w.addEventListener("hashchange", () => {
          const hash = w.location.hash;
          if (!hash) return;
          const allTabs = qsa("[role='tab'],[data-tab-target],[data-target],[data-tab],a[href^='#']");
          const match = allTabs.find((t) => getTargetSelector(t) === hash) || null;
          if (match) activate(match, { updateHash: false });
        });
      }
    };

    return { init, activate };
  })();

  // ------------------------------------------------------------
  // Modals (accessible, focus trap, Escape, overlay click)
  // ------------------------------------------------------------
  App.modals = (() => {
    const openSet = new Set();
    const traps = new Map();
    const lastFocus = new Map();

    const getModal = (idOrEl) => {
      if (isEl(idOrEl)) return idOrEl;
      if (!idOrEl) return null;
      return qs(`[data-modal="${CSS.escape(idOrEl)}"]`) || qs(`#${CSS.escape(idOrEl)}`) || null;
    };

    const setOpen = (modal, open) => {
      if (!isEl(modal)) return;

      modal.classList.toggle("is-open", !!open);
      modal.hidden = !open;
      modal.setAttribute("aria-hidden", open ? "false" : "true");

      if (open) {
        openSet.add(modal);
        lastFocus.set(modal, d.activeElement);
        const release = App.focusTrap.trap(modal);
        traps.set(modal, release);

        // focus first focusable or modal itself
        const focusable = qsa(
          "button:not([disabled]),a[href],input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex='-1'])",
          modal
        ).filter((el) => el.offsetParent !== null);
        (focusable[0] || modal).focus({ preventScroll: true });
        App.bus.emit("modal:open", { modal });
      } else {
        openSet.delete(modal);
        const release = traps.get(modal);
        if (release) release();
        traps.delete(modal);

        const prev = lastFocus.get(modal);
        if (prev && prev.focus) prev.focus({ preventScroll: true });
        lastFocus.delete(modal);

        App.bus.emit("modal:close", { modal });
      }
    };

    const open = (idOrEl) => {
      const modal = getModal(idOrEl);
      if (!modal) return;
      setOpen(modal, true);
    };

    const close = (idOrEl) => {
      const modal = getModal(idOrEl);
      if (!modal) return;
      setOpen(modal, false);
    };

    const closeTop = () => {
      const arr = Array.from(openSet);
      if (!arr.length) return;
      close(arr[arr.length - 1]);
    };

    const init = () => {
      // ensure modal semantics
      qsa("[data-modal]").forEach((m) => {
        if (!m.hasAttribute("role")) m.setAttribute("role", "dialog");
        m.setAttribute("aria-modal", "true");
        if (!m.hasAttribute("tabindex")) m.setAttribute("tabindex", "-1");
        if (!m.hasAttribute("aria-hidden")) m.setAttribute("aria-hidden", "true");
        if (m.hidden == null) m.hidden = true;
      });

      // overlay click (if modal has data-modal-overlay or the modal itself is overlay)
      d.addEventListener("click", (e) => {
        const modal = e.target.closest("[data-modal]");
        if (!modal) return;

        // close if click on overlay region marked or if target is modal root and it is intended as overlay
        const overlaySel = modal.getAttribute("data-modal-overlay");
        if (overlaySel) {
          const overlay = qs(overlaySel, modal);
          if (overlay && overlay.contains(e.target) && e.target === overlay) close(modal);
          return;
        }
        if (e.target === modal && modal.getAttribute("data-overlay-close") !== "false") close(modal);
      });

      // Escape closes topmost
      d.addEventListener("keydown", (e) => {
        if (e.key !== "Escape") return;
        if (openSet.size) closeTop();
      });
    };

    return { init, open, close, closeTop };
  })();

  // ------------------------------------------------------------
  // Drawers (side panels)
  // ------------------------------------------------------------
  App.drawers = (() => {
    const openSet = new Set();
    const traps = new Map();
    const lastFocus = new Map();

    const getDrawer = (idOrEl) => {
      if (isEl(idOrEl)) return idOrEl;
      if (!idOrEl) return null;
      return qs(`[data-drawer="${CSS.escape(idOrEl)}"]`) || qs(`#${CSS.escape(idOrEl)}`) || null;
    };

    const setOpen = (drawer, open) => {
      if (!isEl(drawer)) return;

      drawer.classList.toggle("is-open", !!open);
      drawer.hidden = !open;
      drawer.setAttribute("aria-hidden", open ? "false" : "true");

      if (open) {
        openSet.add(drawer);
        lastFocus.set(drawer, d.activeElement);
        const release = App.focusTrap.trap(drawer);
        traps.set(drawer, release);

        const focusable = qsa(
          "button:not([disabled]),a[href],input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex='-1'])",
          drawer
        ).filter((el) => el.offsetParent !== null);
        (focusable[0] || drawer).focus({ preventScroll: true });
        App.bus.emit("drawer:open", { drawer });
      } else {
        openSet.delete(drawer);
        const release = traps.get(drawer);
        if (release) release();
        traps.delete(drawer);

        const prev = lastFocus.get(drawer);
        if (prev && prev.focus) prev.focus({ preventScroll: true });
        lastFocus.delete(drawer);

        App.bus.emit("drawer:close", { drawer });
      }
    };

    const open = (idOrEl) => {
      const drawer = getDrawer(idOrEl);
      if (!drawer) return;
      setOpen(drawer, true);
    };

    const close = (idOrEl) => {
      const drawer = getDrawer(idOrEl);
      if (!drawer) return;
      setOpen(drawer, false);
    };

    const init = () => {
      qsa("[data-drawer]").forEach((dr) => {
        if (!dr.hasAttribute("role")) dr.setAttribute("role", "complementary");
        if (!dr.hasAttribute("tabindex")) dr.setAttribute("tabindex", "-1");
        if (!dr.hasAttribute("aria-hidden")) dr.setAttribute("aria-hidden", "true");
        if (dr.hidden == null) dr.hidden = true;
      });

      d.addEventListener("keydown", (e) => {
        if (e.key !== "Escape") return;
        // close latest drawer if open
        const arr = Array.from(openSet);
        if (!arr.length) return;
        close(arr[arr.length - 1]);
      });
    };

    return { init, open, close };
  })();

  // ------------------------------------------------------------
  // Tooltips (lightweight, hover/focus, supports rich tooltip content)
  // ------------------------------------------------------------
  App.tooltips = (() => {
    let tip = null;
    let tipInner = null;
    let activeEl = null;

    const ensure = () => {
      if (tip) return tip;
      tip = d.createElement("div");
      tip.className = "app-tooltip";
      tip.style.position = "fixed";
      tip.style.zIndex = "10000";
      tip.style.pointerEvents = "none";
      tip.style.padding = "8px 10px";
      tip.style.borderRadius = "12px";
      tip.style.boxShadow = "0 12px 40px rgba(0,0,0,0.18)";
      tip.style.background = "rgba(15, 23, 42, 0.96)";
      tip.style.color = "white";
      tip.style.fontSize = "12px";
      tip.style.maxWidth = "320px";
      tip.style.lineHeight = "1.35";
      tip.style.opacity = "0";
      tip.style.transform = "translateY(2px)";
      tip.style.transition = "opacity 120ms ease, transform 120ms ease";
      tipInner = d.createElement("div");
      tip.appendChild(tipInner);
      d.body.appendChild(tip);
      return tip;
    };

    const place = (el) => {
      if (!tip || !isEl(el)) return;
      const r = el.getBoundingClientRect();
      const pad = 10;
      const wTip = tip.offsetWidth || 240;
      const hTip = tip.offsetHeight || 40;
      const x = clamp(r.left + r.width / 2 - wTip / 2, pad, w.innerWidth - wTip - pad);
      const yAbove = r.top - hTip - 10;
      const yBelow = r.bottom + 10;
      const y = yAbove > pad ? yAbove : yBelow;
      tip.style.left = `${Math.round(x)}px`;
      tip.style.top = `${Math.round(y)}px`;
    };

    const show = (el) => {
      if (!isEl(el)) return;
      const text = el.getAttribute("data-tooltip");
      const targetSel = el.getAttribute("data-tooltip-target");
      if (!text && !targetSel) return;

      ensure();
      activeEl = el;

      if (targetSel) {
        const target = qs(targetSel) || qs(`#${CSS.escape(targetSel)}`);
        tipInner.innerHTML = target ? target.innerHTML : "";
      } else {
        tipInner.textContent = text || "";
      }

      // display first to measure size
      tip.style.opacity = "0";
      tip.style.transform = "translateY(2px)";
      tip.style.display = "block";
      place(el);

      // animate in
      requestAnimationFrame(() => {
        tip.style.opacity = "1";
        tip.style.transform = "translateY(0)";
      });
    };

    const hide = () => {
      if (!tip) return;
      tip.style.opacity = "0";
      tip.style.transform = "translateY(2px)";
      activeEl = null;
    };

    const init = () => {
      if (!App.config.autoInitTooltips) return;

      const onOver = (e) => {
        const el = e.target.closest("[data-tooltip],[data-tooltip-target]");
        if (!el) return;
        show(el);
      };

      const onOut = (e) => {
        const related = e.relatedTarget;
        if (activeEl && (related === activeEl || (isEl(related) && activeEl.contains(related)))) return;
        hide();
      };

      const onFocus = (e) => {
        const el = e.target.closest("[data-tooltip],[data-tooltip-target]");
        if (!el) return;
        show(el);
      };

      const onBlur = () => hide();

      d.addEventListener("mouseover", onOver);
      d.addEventListener("mouseout", onOut);
      d.addEventListener("focusin", onFocus);
      d.addEventListener("focusout", onBlur);

      w.addEventListener(
        "scroll",
        debounce(() => {
          if (activeEl) place(activeEl);
        }, 16),
        { passive: true }
      );
      w.addEventListener(
        "resize",
        debounce(() => {
          if (activeEl) place(activeEl);
        }, 60)
      );
    };

    return { init, show, hide };
  })();

  // ------------------------------------------------------------
  // Accordions
  // ------------------------------------------------------------
  App.accordions = (() => {
    const init = () => {
      if (!App.config.autoInitAccordions) return;

      qsa("[data-accordion]").forEach((acc) => {
        const items = qsa("[data-acc-item]", acc);
        // fallback: treat any direct children with trigger/panel as items
        const roots = items.length ? items : Array.from(acc.children);

        roots.forEach((item) => {
          const trigger = qs("[data-acc-trigger]", item) || item.querySelector("button, [role='button']");
          const panel = qs("[data-acc-panel]", item) || item.querySelector(".panel, .content, [role='region']");
          if (!trigger || !panel) return;

          if (!trigger.hasAttribute("aria-expanded")) trigger.setAttribute("aria-expanded", "false");
          if (!panel.hasAttribute("aria-hidden")) panel.setAttribute("aria-hidden", "true");
          panel.hidden = trigger.getAttribute("aria-expanded") !== "true";

          trigger.addEventListener("click", (e) => {
            e.preventDefault();
            const expanded = trigger.getAttribute("aria-expanded") === "true";
            trigger.setAttribute("aria-expanded", expanded ? "false" : "true");
            panel.setAttribute("aria-hidden", expanded ? "true" : "false");
            panel.hidden = expanded;
            item.classList.toggle("is-open", !expanded);
          });
        });
      });
    };

    return { init };
  })();

  // ------------------------------------------------------------
  // Persistence for form controls via data-persist
  // ------------------------------------------------------------
  App.persistence = (() => {
    const prefix = App.config.storagePrefix + "persist:";

    const keyFor = (el) => el.getAttribute("data-persist");

    const readValue = (el) => {
      if (!isEl(el)) return null;
      if (el.type === "checkbox") return !!el.checked;
      if (el.type === "radio") return el.checked ? el.value : null;
      return "value" in el ? el.value : el.textContent;
    };

    const writeValue = (el, v) => {
      if (!isEl(el)) return;
      if (el.type === "checkbox") el.checked = !!v;
      else if (el.type === "radio") el.checked = String(el.value) === String(v);
      else if ("value" in el) el.value = v == null ? "" : String(v);
      else el.textContent = v == null ? "" : String(v);
    };

    const init = () => {
      if (!App.config.autoInitPersistence) return;

      const els = qsa("[data-persist]");
      els.forEach((el) => {
        const k = keyFor(el);
        if (!k) return;
        const stored = localStorage.getItem(prefix + k);
        if (stored != null) {
          const v = safeJSON.parse(stored, stored);
          writeValue(el, v);
        }

        const save = debounce(() => {
          const v = readValue(el);
          localStorage.setItem(prefix + k, safeJSON.stringify(v, ""));
          App.bus.emit("persist:save", { el, key: k, value: v });
        }, 150);

        el.addEventListener("input", save);
        el.addEventListener("change", save);
      });
    };

    return { init };
  })();

  // ------------------------------------------------------------
  // Action dispatcher (buttons/links)
  // ------------------------------------------------------------
  App.actions = (() => {
    const getScopeRoot = (el) => {
      const sel = el.getAttribute("data-scope");
      if (sel) return qs(sel) || d;
      const closestScope = el.closest("[data-scope-root]");
      return closestScope || d;
    };

    const getTextFromTarget = (el) => {
      const scope = getScopeRoot(el);
      const targetSel =
        el.getAttribute("data-copy-target") ||
        el.getAttribute("data-target") ||
        el.getAttribute("data-from") ||
        el.getAttribute("data-export-target");

      if (targetSel) {
        const t = qs(targetSel, scope) || qs(targetSel);
        if (t) {
          if ("value" in t) return String(t.value ?? "");
          return String(t.textContent ?? "").trim();
        }
      }

      // fallback: nearest code/textarea/pre within panel context
      const panel = el.closest("[role='tabpanel'], .tab-panel, .panel, .card, section, article");
      if (panel) {
        const cand = panel.querySelector("textarea, pre, code");
        if (cand) return "value" in cand ? String(cand.value ?? "") : String(cand.textContent ?? "").trim();
      }

      return "";
    };

    const hardenClickable = (el) => {
      if (!isEl(el)) return;
      if (el.getAttribute("data-locked") === "true") return;

      // remove common disabling artefacts
      if (el.hasAttribute("disabled")) el.removeAttribute("disabled");
      if (el.getAttribute("aria-disabled") === "true") el.setAttribute("aria-disabled", "false");
      el.classList.remove("disabled");
      el.style.pointerEvents = "";
    };

    const ACTIONS = {
      // Tabs
      "activate-tab": (btn) => {
        const target = btn.getAttribute("data-activate-tab") || btn.getAttribute("data-target");
        if (!target) return;
        const sel = target.startsWith("#") ? target : `#${target}`;
        const allTabs = qsa("[role='tab'],[data-tab-target],[data-target],[data-tab],a[href^='#']");
        const match = allTabs.find((t) => {
          const dt = t.getAttribute("data-tab-target") || t.getAttribute("data-target");
          const ac = t.getAttribute("aria-controls");
          const href = t.getAttribute("href");
          return (dt && (dt === sel || `#${dt}` === sel)) || (ac && `#${ac}` === sel) || (href && href === sel);
        });
        if (match) App.tabs.activate(match, { updateHash: true, focusPanel: true });
      },

      // Clipboard
      copy: async (btn) => {
        const text = getTextFromTarget(btn);
        if (!text) return App.toast.show("Nothing to copy.", "error");
        try {
          await copyToClipboard(text);
          App.toast.show("Copied.", "success");
        } catch {
          App.toast.show("Copy failed.", "error");
        }
      },

      // Exports
      "download-json": (btn) => {
        const literal = btn.getAttribute("data-json");
        let obj = literal ? safeJSON.parse(literal, null) : null;
        if (!obj) {
          const txt = getTextFromTarget(btn);
          obj = safeJSON.parse(txt, null);
        }
        if (!obj) return App.toast.show("No valid JSON found.", "error");
        const name = btn.getAttribute("data-filename") || "export.json";
        downloadText(name, safeJSON.stringify(obj, "null"), "application/json;charset=utf-8");
        App.toast.show("JSON downloaded.", "success");
      },

      "download-csv": (btn) => {
        const txt = getTextFromTarget(btn);
        let csv = "";
        if (txt && txt.includes(",") && txt.includes("\n") && !txt.trim().startsWith("{") && !txt.trim().startsWith("[")) {
          csv = txt;
        } else {
          const parsed = safeJSON.parse(txt, null);
          if (parsed) csv = toCSV(parsed);
        }
        if (!csv) return App.toast.show("No CSV-compatible content found.", "error");
        const name = btn.getAttribute("data-filename") || "export.csv";
        downloadText(name, csv, "text/csv;charset=utf-8");
        App.toast.show("CSV downloaded.", "success");
      },

      "import-json": async (btn) => {
        // Opens file picker, reads JSON, injects into target element if provided.
        const targetSel = btn.getAttribute("data-import-target") || btn.getAttribute("data-target");
        const scope = getScopeRoot(btn);
        const target = targetSel ? qs(targetSel, scope) || qs(targetSel) : null;

        const input = d.createElement("input");
        input.type = "file";
        input.accept = "application/json,.json";
        input.style.display = "none";
        d.body.appendChild(input);

        const readFile = () =>
          new Promise((resolve) => {
            input.addEventListener(
              "change",
              () => {
                const file = input.files && input.files[0];
                if (!file) return resolve(null);
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = () => resolve(null);
                reader.readAsText(file);
              },
              { once: true }
            );
            input.click();
          });

        const text = await readFile();
        input.remove();

        if (!text) return App.toast.show("No file selected.", "warning");
        const obj = safeJSON.parse(String(text), null);
        if (!obj) return App.toast.show("Invalid JSON file.", "error");

        if (target) {
          if ("value" in target) target.value = JSON.stringify(obj, null, 2);
          else target.textContent = JSON.stringify(obj, null, 2);
          App.toast.show("Imported.", "success");
          App.bus.emit("import:json", { obj, target });
        } else {
          App.toast.show("Imported (no target bound).", "success");
          App.bus.emit("import:json", { obj, target: null });
        }
      },

      // Navigation / external open
      "open-url": (btn) => {
        const url = btn.getAttribute("data-url");
        if (!url) return;
        const sameTab = btn.getAttribute("data-same-tab") === "true";
        if (sameTab) w.location.href = url;
        else w.open(url, "_blank", "noopener,noreferrer");
      },

      "open-copilot": (btn) => {
        const url = btn.getAttribute("data-url") || "https://copilot.microsoft.com/";
        w.open(url, "_blank", "noopener,noreferrer");
      },

      // UI helpers
      "scroll-to": (btn) => {
        const sel = btn.getAttribute("data-scroll-target") || btn.getAttribute("data-target");
        if (!sel) return;
        const t = qs(sel) || qs(`#${CSS.escape(sel)}`);
        if (t) t.scrollIntoView({ behavior: "smooth", block: "start" });
      },

      toggle: (btn) => {
        const sel = btn.getAttribute("data-toggle-target") || btn.getAttribute("data-target");
        const cls = btn.getAttribute("data-toggle-class") || "is-open";
        if (!sel) return;
        const t = qs(sel) || qs(`#${CSS.escape(sel)}`);
        if (t) t.classList.toggle(cls);
      },

      reset: (btn) => {
        const scopeSel = btn.getAttribute("data-reset-scope");
        const root = scopeSel ? qs(scopeSel) || d : getScopeRoot(btn);
        qsa("form", root).forEach((f) => f.reset());
        App.toast.show("Reset complete.", "success");
      },

      print: () => w.print(),

      // Modals & drawers
      "modal-open": (btn) => {
        const id = btn.getAttribute("data-modal-id") || btn.getAttribute("data-target");
        if (!id) return;
        App.modals.open(id);
      },
      "modal-close": (btn) => {
        const id = btn.getAttribute("data-modal-id") || btn.closest("[data-modal]")?.getAttribute("data-modal");
        if (!id) return;
        App.modals.close(id);
      },
      "drawer-open": (btn) => {
        const id = btn.getAttribute("data-drawer-id") || btn.getAttribute("data-target");
        if (!id) return;
        App.drawers.open(id);
      },
      "drawer-close": (btn) => {
        const id = btn.getAttribute("data-drawer-id") || btn.closest("[data-drawer]")?.getAttribute("data-drawer");
        if (!id) return;
        App.drawers.close(id);
      },

      // Theme
      "theme-toggle": () => {
        const cur = App.store.state.ui.theme;
        const next = cur === "dark" ? "light" : "dark";
        App.store.state.ui = { ...App.store.state.ui, theme: next };
        App.store.save("theme", next);
        d.documentElement.setAttribute("data-theme", next);
        App.toast.show(`Theme: ${next}`, "success");
      },

      // Notify
      toast: (btn) => {
        const msg = btn.getAttribute("data-toast") || "Done.";
        const type = btn.getAttribute("data-toast-type") || "info";
        App.toast.show(msg, type);
      }
    };

    const init = () => {
      // Delegated click handling
      d.addEventListener("click", async (e) => {
        const trigger = closestAttr(e.target, "data-action");
        if (!trigger) return;

        hardenClickable(trigger);

        // prevent navigation for anchor-buttons unless allowed
        const allowNav = trigger.getAttribute("data-allow-nav") === "true";
        if (trigger.tagName === "A" && !allowNav) e.preventDefault();

        const action = (trigger.getAttribute("data-action") || "").trim().toLowerCase();
        if (!action) return;

        const fn = ACTIONS[action];
        if (!fn) {
          App.toast.show(`Unknown action: ${action}`, "error");
          return;
        }

        try {
          const out = fn(trigger, e);
          if (out && typeof out.then === "function") await out;
        } catch (err) {
          console.error(err);
          App.toast.show("Action failed.", "error");
        }
      });

      // Keyboard activation for non-button elements with data-action
      d.addEventListener("keydown", (e) => {
        if (e.key !== "Enter" && e.key !== " ") return;
        const el = closestAttr(e.target, "data-action");
        if (!el) return;
        if (el.tagName === "BUTTON" || el.tagName === "A" || el.getAttribute("role") === "button") return;
        e.preventDefault();
        el.click();
      });

      // Harden all data-action elements for accessibility
      const hardenAll = () => {
        qsa("[data-action]").forEach((el) => {
          hardenClickable(el);
          if (el.tagName === "DIV" || el.tagName === "SPAN") {
            if (!el.hasAttribute("role")) el.setAttribute("role", "button");
            if (!el.hasAttribute("tabindex")) el.setAttribute("tabindex", "0");
          }
        });
      };
      hardenAll();

      const obs = new MutationObserver(debounce(hardenAll, 50));
      obs.observe(d.body, { childList: true, subtree: true });
    };

    return { init };
  })();

  // ------------------------------------------------------------
  // Global “hardening” to fix inactive tabs/buttons caused by disabled classes/pointer events
  // ------------------------------------------------------------
  App.hardenUI = (() => {
    const run = () => {
      qsa(".disabled,[aria-disabled='true']").forEach((el) => {
        if (el.getAttribute("data-locked") === "true") return;
        if (el.matches("[role='tab'],[data-tab-target],[data-tab],a[href^='#'],button,[data-action]")) {
          el.classList.remove("disabled");
          if (el.getAttribute("aria-disabled") === "true") el.setAttribute("aria-disabled", "false");
          if (el.hasAttribute("disabled")) el.removeAttribute("disabled");
          el.style.pointerEvents = "";
        }
      });

      // If something is visually a tab but missing role/attributes, try to fix gently
      qsa("[data-tab-target],[data-target],[data-tab]").forEach((t) => {
        if (t.getAttribute("data-locked") === "true") return;
        if (!t.hasAttribute("role")) t.setAttribute("role", "tab");
      });
    };

    const init = () => {
      run();
      const obs = new MutationObserver(debounce(run, 60));
      obs.observe(d.body, { childList: true, subtree: true });
    };

    return { init, run };
  })();

  // ------------------------------------------------------------
  // Error boundary (global)
  // ------------------------------------------------------------
  App.errors = (() => {
    const init = () => {
      w.addEventListener("error", (e) => {
        App.log.error("Unhandled error:", e.error || e.message);
        App.toast.show("Something went wrong. Check console for details.", "error", 4200);
      });
      w.addEventListener("unhandledrejection", (e) => {
        App.log.error("Unhandled rejection:", e.reason);
        App.toast.show("An action failed. Check console for details.", "error", 4200);
      });
    };
    return { init };
  })();

  // ------------------------------------------------------------
  // Boot
  // ------------------------------------------------------------
  App.boot = () => {
    // Theme apply
    const theme = App.store.load("theme", App.store.state.ui.theme);
    App.store.state.ui = { ...App.store.state.ui, theme };
    d.documentElement.setAttribute("data-theme", theme);

    App.errors.init();
    App.modals.init();
    App.drawers.init();
    App.tooltips.init();
    App.accordions.init();
    App.persistence.init();
    App.tabs.init();
    App.actions.init();
    App.hardenUI.init();

    // If hash exists, tabs init already attempts to respect it; re-run harden for safety
    App.hardenUI.run();

    App.bus.emit("app:ready", { ts: now() });
  };

  if (d.readyState === "loading") d.addEventListener("DOMContentLoaded", App.boot);
  else App.boot();
})();
