/* Pipeline Flow – Tableau dashboard extension.
 * Reads Segment/Category/Sales Region/Sales/Profit from the "Flow Data" worksheet, draws an animated
 * 3-column Sankey coloured by margin, derives three insight chips, and turns clicks into mark
 * selections on native sheets so the dashboard's filter actions cross-filter everything.
 * Outside Tableau it falls back to demo data so the layout can be previewed in a browser.
 */
(function () {
  "use strict";

  const DATA_SHEET = "Flow Data";
  const SEGMENT_SHEET = "Segment Mix";
  const GRID_SHEET = "Region x Category";
  const COLS = ["Segment", "Category", "Sales Region"];
  const HEADS = ["Segment", "Category", "Region"];
  const NODE_COLORS = {
    Consumer: "#4C8DFF", Corporate: "#2EC4B6", "Home Office": "#FFB547",
    Furniture: "#9AA8C7", "Office Supplies": "#9AA8C7", Technology: "#9AA8C7",
  };
  const REGION_COLOR = "#C9D3E6";
  const SVGNS = "http://www.w3.org/2000/svg";

  let dashboard = null;
  let rows = [];
  let selected = false;

  // ------------------------------------------------------------------ formatting
  const money = (v) => {
    const a = Math.abs(v), s = v < 0 ? "-" : "";
    if (a >= 1e6) return `${s}$${(a / 1e6).toFixed(2)}M`;
    if (a >= 1e3) return `${s}$${Math.round(a / 1e3)}K`;
    return `${s}$${Math.round(a)}`;
  };
  const pct = (v, d = 1) => `${(v * 100).toFixed(d)}%`;

  // red (<=0%) -> amber (~8%) -> teal (>=20%)
  function marginColor(m) {
    const stops = [[0, [255, 92, 108]], [0.08, [255, 181, 71]], [0.2, [46, 196, 182]]];
    const t = Math.max(stops[0][0], Math.min(stops[2][0], m));
    const [a, b] = t <= stops[1][0] ? [stops[0], stops[1]] : [stops[1], stops[2]];
    const f = (t - a[0]) / (b[0] - a[0]);
    const c = a[1].map((x, i) => Math.round(x + (b[1][i] - x) * f));
    return `rgb(${c.join(",")})`;
  }

  // ------------------------------------------------------------------ data
  function demoRows() {
    const seg = { Consumer: 0.5, Corporate: 0.31, "Home Office": 0.19 };
    const grid = { // region -> category -> [sales, margin]
      Central: { Furniture: [164538, -0.017], "Office Supplies": [168217, 0.053], Technology: [170416, 0.198] },
      East: { Furniture: [212232, 0.016], "Office Supplies": [211658, 0.203], Technology: [267938, 0.181] },
      South: { Furniture: [117299, 0.058], "Office Supplies": [125651, 0.159], Technology: [148772, 0.134] },
      West: { Furniture: [260680, 0.047], "Office Supplies": [226367, 0.239], Technology: [252767, 0.176] },
    };
    const out = [];
    for (const [s, share] of Object.entries(seg))
      for (const [r, cats] of Object.entries(grid))
        for (const [c, [sales, m]] of Object.entries(cats)) {
          const jitter = 1 + (s.length % 3 - 1) * 0.04;
          out.push({ Segment: s, Category: c, "Sales Region": r, sales: sales * share, profit: sales * share * m * jitter });
        }
    return out;
  }

  function pick(cols, name) {
    return cols.findIndex((c) => c.fieldName === name || c.fieldName.replace(/^(SUM|AGG)\((.*)\)$/, "$2") === name);
  }

  async function loadRows() {
    const ws = dashboard.worksheets.find((w) => w.name === DATA_SHEET);
    if (!ws) throw new Error(`Worksheet "${DATA_SHEET}" is not on this dashboard`);
    const table = await ws.getSummaryDataAsync({ ignoreSelection: true });
    const cols = table.columns;
    const idx = { Segment: pick(cols, "Segment"), Category: pick(cols, "Category"),
      "Sales Region": pick(cols, "Sales Region"), sales: pick(cols, "Sales"), profit: pick(cols, "Profit") };
    return table.data.map((r) => ({
      Segment: r[idx.Segment].formattedValue, Category: r[idx.Category].formattedValue,
      "Sales Region": r[idx["Sales Region"]].formattedValue,
      sales: Number(r[idx.sales].value) || 0, profit: Number(r[idx.profit].value) || 0,
    })).filter((r) => r.sales > 0);
  }

  // ------------------------------------------------------------------ sankey model
  function buildModel(data) {
    const nodes = new Map(), links = new Map();
    const node = (col, name) => {
      const k = `${col}|${name}`;
      if (!nodes.has(k)) nodes.set(k, { key: k, col, name, sales: 0, profit: 0, in: [], out: [] });
      return nodes.get(k);
    };
    const link = (a, b, r) => {
      const k = `${a.key}>${b.key}`;
      if (!links.has(k)) { const l = { key: k, source: a, target: b, sales: 0, profit: 0 }; links.set(k, l); a.out.push(l); b.in.push(l); }
      const l = links.get(k); l.sales += r.sales; l.profit += r.profit;
    };
    for (const r of data) {
      const ns = COLS.map((c, i) => node(i, r[c]));
      ns.forEach((n) => { n.sales += r.sales; n.profit += r.profit; });
      link(ns[0], ns[1], r); link(ns[1], ns[2], r);
    }
    // node.sales was counted once per row, which equals the throughput for every column
    return { nodes: [...nodes.values()], links: [...links.values()] };
  }

  function layout(model, W, H) {
    const top = 18, bottom = 4, labelW = 118, nodeW = 12, pad = 10;
    const colsX = [labelW, (W - nodeW) / 2, W - labelW - nodeW];
    const byCol = [0, 1, 2].map((c) => model.nodes.filter((n) => n.col === c).sort((a, b) => b.sales - a.sales));
    const total = byCol[0].reduce((s, n) => s + n.sales, 0) || 1;
    const k = Math.min(...byCol.map((ns) => (H - top - bottom - pad * (ns.length - 1)) / total));
    byCol.forEach((ns, c) => {
      const used = total * k + pad * (ns.length - 1);
      let y = top + (H - top - bottom - used) / 2;
      ns.forEach((n) => { n.x = colsX[c]; n.y = y; n.h = Math.max(1, n.sales * k); n.w = nodeW; y += n.h + pad; });
    });
    // two passes: link widths are set while stacking outgoing links, then incoming links reuse them
    for (const n of model.nodes) {
      n.out.sort((a, b) => a.target.y - b.target.y);
      let oy = n.y;
      n.out.forEach((l) => { l.w = Math.max(1, l.sales * k); l.y0 = oy + l.w / 2; oy += l.w; });
    }
    for (const n of model.nodes) {
      n.in.sort((a, b) => a.source.y - b.source.y);
      let iy = n.y;
      n.in.forEach((l) => { l.y1 = iy + l.w / 2; iy += l.w; });
    }
    return { colsX, nodeW, byCol };
  }

  // ------------------------------------------------------------------ rendering
  const el = (tag, attrs = {}, parent) => {
    const e = document.createElementNS(SVGNS, tag);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    if (parent) parent.appendChild(e);
    return e;
  };

  function render() {
    const svg = document.getElementById("svg");
    const box = document.getElementById("stage").getBoundingClientRect();
    const W = Math.max(320, box.width), H = Math.max(160, box.height);
    svg.setAttribute("viewBox", `0 0 ${W} ${H}`);
    svg.innerHTML = "";
    const model = buildModel(rows);
    if (!model.nodes.length) { el("text", { x: W / 2, y: H / 2, "text-anchor": "middle", class: "lbl" }, svg).textContent = "No data for the current filters"; return; }
    const L = layout(model, W, H);

    HEADS.forEach((h, c) => {
      const x = c === 0 ? L.colsX[0] + L.nodeW : c === 2 ? L.colsX[2] : L.colsX[1] + L.nodeW / 2;
      el("text", { x, y: 8, class: "colhead", "text-anchor": c === 0 ? "end" : c === 2 ? "start" : "middle" }, svg).textContent = h;
    });

    const gLinks = el("g", {}, svg), gNodes = el("g", {}, svg);
    const tip = document.getElementById("tip");
    const linkEls = [];

    model.links.forEach((l, i) => {
      const x0 = l.source.x + L.nodeW, x1 = l.target.x, xm = (x0 + x1) / 2;
      const d = `M${x0},${l.y0} C${xm},${l.y0} ${xm},${l.y1} ${x1},${l.y1}`;
      const m = l.sales ? l.profit / l.sales : 0;
      const p = el("path", { d, class: "link", stroke: marginColor(m), "stroke-width": l.w, "stroke-opacity": 0.42 }, gLinks);
      const len = p.getTotalLength();
      p.style.strokeDasharray = `${len} ${len}`;
      p.style.strokeDashoffset = len;
      p.style.transition = `stroke-dashoffset 900ms cubic-bezier(.2,.7,.2,1) ${80 + i * 18}ms, stroke-opacity .25s`;
      requestAnimationFrame(() => requestAnimationFrame(() => { p.style.strokeDashoffset = 0; }));
      p._l = l; linkEls.push(p);
      p.addEventListener("mousemove", (ev) => showTip(ev, `<b>${l.source.name} → ${l.target.name}</b><br>${money(l.sales)} sales · ${pct(l.sales / total(model))} of pipeline<br>Profit ${money(l.profit)} · <span style="color:${marginColor(m)}">${pct(m)} margin</span>`));
      p.addEventListener("mouseenter", () => focus((x) => x === l));
      p.addEventListener("mouseleave", unfocus);
      p.addEventListener("click", () => selectFlow(l));
    });

    model.nodes.forEach((n, i) => {
      const g = el("g", { class: "node" }, gNodes);
      const color = n.col === 2 ? REGION_COLOR : (NODE_COLORS[n.name] || "#9AA8C7");
      const r = el("rect", { x: n.x, y: n.y, width: n.w, height: n.h, rx: 3, fill: color }, g);
      r.style.transformBox = "fill-box"; r.style.transformOrigin = "center";
      r.style.transform = "scaleY(0)";
      r.style.transition = `transform 650ms cubic-bezier(.2,.7,.2,1) ${i * 35}ms, opacity .25s`;
      requestAnimationFrame(() => requestAnimationFrame(() => { r.style.transform = "scaleY(1)"; }));
      const m = n.sales ? n.profit / n.sales : 0;
      const left = n.col === 0;
      const tx = left ? n.x - 8 : n.col === 2 ? n.x + n.w + 8 : n.x + n.w + 6;
      const t = el("text", { x: tx, y: n.y + n.h / 2 + 3.5, class: "lbl", "text-anchor": left ? "end" : "start" }, g);
      t.textContent = n.name + " ";
      const v = el("tspan", { class: "val" }, t);
      v.textContent = money(n.sales);
      if (n.col === 1) {
        // centre label on the middle node with a dark pill so it stays readable over the links
        t.setAttribute("text-anchor", "middle"); t.setAttribute("x", n.x + n.w / 2); t.setAttribute("y", n.y + n.h / 2 + 3.5);
        requestAnimationFrame(() => {
          const b = t.getBBox();
          const pill = el("rect", { x: b.x - 7, y: b.y - 3, width: b.width + 14, height: b.height + 6, rx: (b.height + 6) / 2,
            fill: "rgba(11,20,38,.88)", stroke: "#22314F" });
          g.insertBefore(pill, t);
        });
      }
      r.addEventListener("mousemove", (ev) => showTip(ev, `<b>${n.name}</b><br>${money(n.sales)} sales · ${pct(n.sales / total(model))}<br>Profit ${money(n.profit)} · <span style="color:${marginColor(m)}">${pct(m)} margin</span>`));
      r.addEventListener("mouseenter", () => focus((x) => x.source === n || x.target === n));
      r.addEventListener("mouseleave", unfocus);
      r.addEventListener("click", () => selectNode(n));
    });

    function focus(pred) { linkEls.forEach((p) => { p.style.strokeOpacity = pred(p._l) ? 0.85 : 0.08; }); }
    function unfocus() { linkEls.forEach((p) => { p.style.strokeOpacity = 0.42; }); tip.style.opacity = 0; }
    function showTip(ev, html) {
      const s = document.getElementById("stage").getBoundingClientRect();
      tip.innerHTML = html; tip.style.opacity = 1;
      const x = ev.clientX - s.left + 14, y = ev.clientY - s.top + 10;
      tip.style.left = `${Math.min(x, s.width - tip.offsetWidth - 4)}px`;
      tip.style.top = `${Math.min(y, s.height - tip.offsetHeight - 4)}px`;
    }
    insights(model);
  }

  const total = (model) => model.nodes.filter((n) => n.col === 0).reduce((s, n) => s + n.sales, 0) || 1;

  function insights(model) {
    const T = total(model);
    // full 3-hop paths from the raw rows
    const paths = new Map();
    for (const r of rows) {
      const k = COLS.map((c) => r[c]).join(" → ");
      const p = paths.get(k) || { k, sales: 0, profit: 0 };
      p.sales += r.sales; p.profit += r.profit; paths.set(k, p);
    }
    const list = [...paths.values()].sort((a, b) => b.sales - a.sales);
    const set = (id, v, d) => { document.getElementById(id + "v").innerHTML = v; document.getElementById(id + "d").textContent = d; };
    if (!list.length) return;
    const big = list[0];
    set("c1", `${money(big.sales)} <span style="color:var(--muted);font-size:11px">${pct(big.sales / T)}</span>`, big.k);
    const material = list.filter((p) => p.sales >= T * 0.015);
    const worst = (material.length ? material : list).reduce((a, b) => (a.profit / a.sales < b.profit / b.sales ? a : b));
    const wm = worst.profit / worst.sales;
    set("c2", `<span style="color:${marginColor(wm)}">${pct(wm)}</span> margin`, `${worst.k} · ${money(worst.profit)}`);
    const n = Math.max(1, Math.round(list.length * 0.2));
    const share = list.slice(0, n).reduce((s, p) => s + p.sales, 0) / T;
    set("c3", `${pct(share, 0)} of sales`, `from the top ${n} of ${list.length} flows`);
    document.querySelectorAll(".chip").forEach((c, i) => { c.classList.remove("show"); setTimeout(() => c.classList.add("show"), 150 + i * 120); });
  }

  // ------------------------------------------------------------------ interaction -> native selections
  function sheet(name) { return dashboard && dashboard.worksheets.find((w) => w.name === name); }

  async function select(name, criteria) {
    const ws = sheet(name);
    if (!ws) return;
    await ws.selectMarksByValueAsync(criteria, tableau.SelectionUpdateType.Replace);
  }

  async function selectFlow(l) {
    if (!dashboard) return;
    const [a, b] = [l.source, l.target];
    if (a.col === 0) await select(SEGMENT_SHEET, [{ fieldName: "Segment", value: [a.name] }]);
    const cat = a.col === 1 ? a.name : b.name;
    const crit = [{ fieldName: "Category", value: [cat] }];
    if (b.col === 2) crit.push({ fieldName: "Sales Region", value: [b.name] });
    await select(GRID_SHEET, crit);
    markSelected(true);
  }

  async function selectNode(n) {
    if (!dashboard) return;
    if (n.col === 0) await select(SEGMENT_SHEET, [{ fieldName: "Segment", value: [n.name] }]);
    else await select(GRID_SHEET, [{ fieldName: COLS[n.col], value: [n.name] }]);
    markSelected(true);
  }

  function markSelected(on) {
    selected = on;
    document.getElementById("reset").classList.toggle("on", on);
  }

  document.getElementById("reset").addEventListener("click", async () => {
    for (const name of [SEGMENT_SHEET, GRID_SHEET]) { const ws = sheet(name); if (ws) await ws.clearSelectedMarksAsync(); }
    markSelected(false);
  });

  // ------------------------------------------------------------------ boot
  let pending = null;
  function refreshSoon() {
    clearTimeout(pending);
    pending = setTimeout(async () => {
      try { rows = await loadRows(); render(); status(""); } catch (e) { status(e.message); }
    }, 250);
  }
  const status = (t) => { document.getElementById("status").textContent = t; };

  async function boot() {
    // Desktop hosts extensions in their own webview (not an iframe), so probe the API instead of window.parent
    const timeout = new Promise((_, reject) => setTimeout(() => reject(new Error("no Tableau host")), 4000));
    try {
      if (!(window.tableau && tableau.extensions)) throw new Error("no Extensions API");
      await Promise.race([tableau.extensions.initializeAsync(), timeout]);
    } catch (e) {
      rows = demoRows(); render(); status("demo data"); return;
    }
    dashboard = tableau.extensions.dashboardContent.dashboard;
    const T = tableau.TableauEventType;
    for (const ws of dashboard.worksheets) {
      ws.addEventListener(T.FilterChanged, refreshSoon);
      if (ws.name !== DATA_SHEET) ws.addEventListener(T.MarkSelectionChanged, (e) => {
        if (ws.name === SEGMENT_SHEET || ws.name === GRID_SHEET) e.getMarksAsync().then((m) => markSelected(m.data.some((t) => t.data.length)));
        refreshSoon();
      });
      if (T.SummaryDataChanged && ws.name === DATA_SHEET) ws.addEventListener(T.SummaryDataChanged, refreshSoon);
    }
    refreshSoon();
  }

  let rz = null;
  window.addEventListener("resize", () => { clearTimeout(rz); rz = setTimeout(() => rows.length && render(), 120); });
  boot().catch((e) => { rows = demoRows(); render(); status("Extension error: " + e.message); });
})();
