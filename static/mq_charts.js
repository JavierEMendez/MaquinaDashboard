/* Maquina chart kit — Chart.js builders in the house style, i.e. the Ember
   company page's charts: Inter (Chart.defaults, set in base.html), muted
   ticks, no vertical grid, 2.5px smoothed lines without points, translucent
   fills from the --mq-c* tokens, index-mode tooltips, projections faded.
   For pages that draw their charts from data in JS (the Ranman tableros).

   One chart per host element: the first draw builds the canvas and registers
   it with MQ (so it rebuilds on a theme flip and resizes when its tab opens);
   later draws with new data update it in place. Call from DOMContentLoaded or
   later — MQ is defined at the end of base.html. */
(function () {
  "use strict";
  const tok = (n) => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
  const color = (c) => (c && c.startsWith('--')) ? tok(c) : c;
  const alpha = (c, a) => MQ.hexA(color(c), a);
  // Fixed categorical order for many-series charts, validated for colour-blind
  // and normal-vision separation of neighbours in both themes: blue, rose,
  // indigo, green, purple, amber, teal. Gray is the last slot, for "Otros".
  const CAT = ['--mq-c1', '--mq-c7', '--mq-c2', '--mq-c3', '--mq-c6', '--mq-c4', '--mq-c8', '--mq-c5'];

  function mount(host, height, make, overlay) {
    let e = host._mqc;
    if (!e) {
      host.innerHTML = `<div class="chart-wrap" style="height:${height}px"><canvas></canvas>${overlay || ''}</div>`;
      const cv = host.querySelector('canvas');
      e = host._mqc = {make, chart: null, fn: () => new Chart(cv, e.make())};
      MQ.charts.push(e);
      e.chart = e.fn();
      return e.chart;
    }
    e.make = make;
    if (overlay != null) {
      const ov = host.querySelector('.mqc-overlay');
      if (ov) ov.outerHTML = overlay;
    }
    const cfg = make();
    e.chart.data = cfg.data;
    e.chart.options = cfg.options;
    e.chart.update();
    return e.chart;
  }

  // Replace a chart with a message (and drop it from the theme registry).
  function clear(host, html) {
    const e = host._mqc;
    if (e) {
      try { e.chart.destroy(); } catch (x) {}
      const k = MQ.charts.indexOf(e);
      if (k >= 0) MQ.charts.splice(k, 1);
      delete host._mqc;
    }
    host.innerHTML = html || '';
  }

  // Vertical dashed markers with a label at the top (a minimum, a bank leaving).
  // list: [{x, label, hot}] — x is a category index or, on a time axis, a timestamp.
  const EVENTS = {
    id: 'mqEvents',
    afterDatasetsDraw(chart, _a, opts) {
      const list = (opts && opts.list) || [];
      if (!list.length) return;
      const {ctx, chartArea: area, scales: {x}} = chart;
      ctx.save();
      ctx.font = "600 11px 'Inter', system-ui, sans-serif";
      ctx.textBaseline = 'top';
      const placed = [];                     // [x0, x1, row] — labels that would collide drop a row
      for (const ev of list.slice().sort((a, b) => a.x - b.x)) {
        const px = x.getPixelForValue(ev.x);
        if (!isFinite(px) || px < area.left - 1 || px > area.right + 1) continue;
        const c = ev.hot ? MQ.bad() : MQ.muted();
        ctx.strokeStyle = c; ctx.fillStyle = c; ctx.lineWidth = 1; ctx.setLineDash([3, 3]);
        ctx.beginPath(); ctx.moveTo(px, area.top); ctx.lineTo(px, area.bottom); ctx.stroke();
        const w = ctx.measureText(ev.label).width;
        const lx = px + 6 + w > area.right ? px - 6 - w : px + 6;
        let row = 0;
        while (placed.some(([a, b, r]) => r === row && lx < b + 6 && lx + w > a - 6)) row++;
        placed.push([lx, lx + w, row]);
        ctx.fillText(ev.label, lx, area.top + 2 + row * 14);
      }
      ctx.restore();
    },
  };

  // The year's plan as ticks over a bar: solid = target, dashed = plan to date.
  // A dataset carries marks: [{base, ytd} | null per index].
  const MARKS = {
    id: 'mqMarks',
    afterDatasetsDraw(chart) {
      const {ctx} = chart, done = new Set();
      chart.data.datasets.forEach((ds, di) => {
        if (!ds.marks) return;
        const meta = chart.getDatasetMeta(di), y = chart.scales[ds.yAxisID || 'y'];
        meta.data.forEach((bar, i) => {
          const m = ds.marks[i], key = ds.stack + ':' + i;
          if (!m || ds.data[i] == null || done.has(key) || !isFinite(bar.x)) return;
          done.add(key);
          const half = bar.width / 2 + 3;
          ctx.save();
          ctx.strokeStyle = MQ.ink(); ctx.lineWidth = 2; ctx.lineCap = 'round';
          for (const [v, dash] of [[m.base, []], [m.ytd, [4, 3]]]) {
            if (v == null) continue;
            const py = y.getPixelForValue(v);
            ctx.setLineDash(dash);
            ctx.beginPath(); ctx.moveTo(bar.x - half, py); ctx.lineTo(bar.x + half, py); ctx.stroke();
          }
          ctx.restore();
        });
      });
    },
  };

  // Year ticks on a time axis (timestamps), one at each year's first point.
  function timeAxis(ts) {
    const firsts = [];
    ts.forEach((t) => {
      const y = new Date(t).getFullYear();
      if (!firsts.length || new Date(firsts[firsts.length - 1]).getFullYear() !== y) firsts.push(t);
    });
    return {
      type: 'linear', min: ts[0], max: ts[ts.length - 1], grid: {display: false},
      ticks: {color: MQ.muted(), maxRotation: 0, autoSkipPadding: 24,
              callback: (v) => String(new Date(v).getFullYear())},
      afterBuildTicks: (ax) => { ax.ticks = firsts.map((value) => ({value})); },
    };
  }

  // Options shared by every cartesian chart, as the Ember charts set them.
  function options(o) {
    const x = Object.assign({grid: {display: false}, ticks: {color: MQ.muted(), maxTicksLimit: o.xTicks || 10, maxRotation: 0}},
                            o.time ? timeAxis(o.time) : {});
    const y = {grid: {color: MQ.grid()}, ticks: {color: MQ.muted()}};
    if (o.yFmt) y.ticks.callback = o.yFmt;
    if (o.zero) y.beginAtZero = true;
    if (o.yTitle) y.title = {display: true, text: o.yTitle, color: MQ.muted(), font: {size: 11}};
    // Stack values only: a stacked x would add up the timestamps of a time
    // axis (bars group their stacks on x themselves).
    if (o.stacked) y.stacked = true;
    return {
      responsive: true, maintainAspectRatio: false,
      interaction: {mode: 'index', intersect: false},
      scales: {x, y},
      plugins: {
        legend: {display: false},
        tooltip: Object.assign({itemSort: (a, b) => b.datasetIndex - a.datasetIndex}, o.tooltip || {}),
        mqEvents: {list: o.events || []},
      },
    };
  }
  // A time axis plots {x, y} points; a category axis plots values by index.
  const pts = (o, vals) => o.time ? vals.map((v, i) => ({x: o.time[i], y: v})) : vals;

  // Lines with the area tinted to zero (Ember's cumulative-distributions chart).
  // series: [{name, color, vals}]; o.zero keeps 0 on the axis.
  function line(host, h, labels, series, o) {
    o = o || {};
    return mount(host, h, () => ({
      type: 'line',
      data: {labels, datasets: series.map((s) => ({
        label: s.name, data: pts(o, s.vals), borderColor: color(s.color),
        backgroundColor: alpha(s.color, 0.08), borderWidth: 2.5, tension: 0.3,
        pointRadius: 0, pointHoverRadius: 4, fill: 'origin', spanGaps: true,
      }))},
      options: options(o), plugins: [EVENTS],
    }));
  }

  // Stacked areas; o.total adds an unstacked line over the stack.
  function area(host, h, labels, series, o) {
    o = o || {};
    return mount(host, h, () => {
      const ds = series.map((s, i) => ({
        label: s.name, data: pts(o, s.vals), stack: 'a', borderColor: color(s.color),
        backgroundColor: alpha(s.color, 0.7), borderWidth: 1.5, cubicInterpolationMode: 'monotone',
        pointRadius: 0, pointHoverRadius: 3, fill: i ? '-1' : 'origin',
      }));
      if (o.total) {
        const n = series.length ? series[0].vals.length : 0;
        const tot = Array.from({length: n}, (_, i) => series.reduce((a, s) => a + (s.vals[i] || 0), 0));
        ds.push({label: o.total, data: pts(o, tot), stack: 'total', borderColor: MQ.ink(),
                 backgroundColor: 'transparent', borderWidth: 1.5, cubicInterpolationMode: 'monotone',
                 pointRadius: 0, pointHoverRadius: 3, fill: false, _total: true});
      }
      // A stack reads from zero: a cut baseline would shrink the bottom band.
      const opts = options(Object.assign({}, o, {stacked: true, zero: true}));
      // Hide series with nothing at that point; top band first.
      opts.plugins.tooltip.filter = (it) => it.dataset._total || Math.abs(it.parsed.y) > 0.05;
      return {type: 'line', data: {labels, datasets: ds}, options: opts, plugins: [EVENTS]};
    });
  }

  // Bars with the site's projection convention: what's real is solid, what's
  // projected faded. Each series is two datasets sharing a stack — real part
  // and projected part — so a year in progress splits, real below.
  // series: [{name, color, neg?, real: [], proj: [], marks?: [], tip?: i => [lines]}]
  function bars(host, h, labels, series, o) {
    o = o || {};
    return mount(host, h, () => {
      const fill = (s, vals, a) => vals.map((v) => alpha(v != null && v < 0 && s.neg ? s.neg : s.color, a));
      const ds = [];
      series.forEach((s, k) => {
        const common = {stack: 's' + k, borderWidth: 0, maxBarThickness: o.maxBar || 44, marks: s.marks, _series: k};
        ds.push(Object.assign({label: s.name, data: s.real, backgroundColor: fill(s, s.real, 0.75), _part: 'real'}, common));
        ds.push(Object.assign({label: s.name, data: s.proj, backgroundColor: fill(s, s.proj, 0.32), _part: 'proj'}, common));
      });
      const opts = options(Object.assign({}, o, {stacked: true}));
      opts.scales.x.stacked = true;          // real + projected share one bar
      // Keep the plan ticks inside the plot when a target tops the bars.
      const mk = series.flatMap((s) => (s.marks || []).flatMap((m) => m ? [m.base, m.ytd] : []))
                       .filter((v) => v != null);
      if (mk.length) opts.scales.y.suggestedMax = Math.max(...mk) * 1.04;
      // One tooltip entry per series (its real part, else its projected part).
      Object.assign(opts.plugins.tooltip, {
        itemSort: (a, b) => a.datasetIndex - b.datasetIndex,
        filter: (it) => it.raw != null && (it.dataset._part === 'real'
          || it.chart.data.datasets[it.datasetIndex - 1].data[it.dataIndex] == null),
        callbacks: {label: (c) => {
          const s = series[c.dataset._series];
          return s.tip ? s.tip(c.dataIndex) : `${s.name}: ${c.raw}`;
        }},
      });
      return {type: 'bar', data: {labels, datasets: ds}, options: opts, plugins: [MARKS]};
    });
  }

  // Doughnut as Ember's equity-by-class donut; `center` is HTML laid over the hole.
  function donut(host, h, labels, vals, colors, o) {
    o = o || {};
    const overlay = o.center != null ? `<div class="mqc-overlay">${o.center}</div>` : null;
    return mount(host, h, () => ({
      type: 'doughnut',
      data: {labels, datasets: [{data: vals, backgroundColor: colors.map(color), borderWidth: 0, cutout: '62%'}]},
      options: {responsive: true, maintainAspectRatio: false,
                plugins: {legend: {display: false}, tooltip: o.tooltip || {}}},
    }), overlay);
  }

  window.MQC = {CAT, color, alpha, mount, clear, line, area, bars, donut};
})();
