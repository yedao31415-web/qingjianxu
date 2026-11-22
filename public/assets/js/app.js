/*!
 * 青简叙 · QingJianXu
 * 考试纪事与倒数日。纯前端、零运行时依赖，所有数据仅保存在本机 IndexedDB。
 * 与 2.0「考试雷达」使用不同的数据库与缓存命名空间，互不影响。
 */
(function () {
  "use strict";

  var Store = window.QingJianStore;
  var Exporter = window.QingJianExport;

  var S = {
    ready: false,
    datasets: [],
    activeId: null,
    records: [],
    cache: new Map(),
    profiles: [],
    profile: null,
    view: "countdown",
    scope: "class",
    filters: { grade: "", major: "", classKey: "", q: "" },
    hashClass: "",
    pastOpen: false,
    starred: new Set(),
    entries: [],
    events: [],
    groups: {},
    scopeRecords: [],
    wallpaper: null,
    listLimit: 120
  };

  var WEEK = ["日", "一", "二", "三", "四", "五", "六"];
  var libs = {};

  /* ================================ 工具 ================================ */
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function pad(n) { return String(n).padStart(2, "0"); }

  function todayStart() {
    var d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }

  function daysUntil(dateStr) {
    var d = new Date(dateStr + "T00:00:00");
    if (isNaN(d)) return null;
    return Math.round((d - todayStart()) / 86400000);
  }

  function mdLabel(dateStr) {
    var p = String(dateStr || "").split("-");
    return p.length === 3 ? (+p[1]) + "/" + p[2] : dateStr;
  }

  function fullDateLabel(dateStr) {
    var p = String(dateStr || "").split("-");
    if (p.length !== 3) return dateStr;
    var y = +p[0];
    var prefix = y === new Date().getFullYear() ? "" : y + "年";
    return prefix + (+p[1]) + "月" + (+p[2]) + "日";
  }

  function weekdayOf(dateStr) {
    var d = new Date(dateStr + "T00:00:00");
    return isNaN(d) ? "" : "周" + WEEK[d.getDay()];
  }

  function timeRange(r) {
    if (!r.start) return "时间待定";
    return r.start + (r.end ? "-" + r.end : "");
  }

  function roomsText(r) {
    return (r.rooms && r.rooms.length) ? r.rooms.join("、") : "地点待定";
  }

  function uniqueBy(arr, fn) {
    var seen = new Set();
    return arr.filter(function (x) {
      var k = fn(x);
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }

  var toastTimer = null;
  function toast(msg, kind) {
    var el = $("#toast");
    el.textContent = msg;
    el.className = "toast is-show" + (kind === "error" ? " is-error" : "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.className = "toast"; }, kind === "error" ? 3600 : 2400);
  }

  function loading(on, text, sub) {
    var el = $("#loading");
    $("#loadingText").textContent = text || "处理中…";
    $("#loadingSub").textContent = sub || "";
    el.hidden = !on;
  }

  function loadScript(src) {
    if (libs[src]) return libs[src];
    libs[src] = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src;
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () { reject(new Error("资源加载失败：" + src)); };
      document.head.appendChild(s);
    });
    return libs[src];
  }

  async function ensureXLSX() {
    if (!window.XLSX) await loadScript("vendor/xlsx.full.min.js");
    if (!window.QingJianParser) await loadScript("assets/js/parser.js");
    return window.XLSX;
  }

  async function ensurePDF() {
    if (!window.jspdf) await loadScript("vendor/jspdf.umd.min.js");
    if (!window.html2canvas) await loadScript("vendor/html2canvas.min.js");
  }

  /* ============================== 底部面板 ============================== */
  var sheetState = null;
  var sheetCloseTimer = null;

  function openSheet(opts) {
    if (sheetCloseTimer) {
      clearTimeout(sheetCloseTimer);
      sheetCloseTimer = null;
    }
    return new Promise(function (resolve) {
      var layer = $("#sheetLayer");
      $("#sheetTitle").textContent = opts.title || "";
      $("#sheetBody").innerHTML = opts.body || "";

      var foot = $("#sheetFoot");
      if (opts.foot) {
        foot.hidden = false;
        foot.innerHTML = opts.foot;
      } else {
        foot.hidden = true;
        foot.innerHTML = "";
      }

      layer.hidden = false;
      layer.classList.remove("is-closing");
      document.body.style.overflow = "hidden";
      sheetState = { resolve: resolve, opts: opts };

      if (opts.onMount) opts.onMount($("#sheetBody"), closeSheet);
      if (opts.foot && opts.onFoot) opts.onFoot(foot, closeSheet);
    });
  }

  function closeSheet(value) {
    var layer = $("#sheetLayer");
    if (!sheetState) return;
    var st = sheetState;
    sheetState = null;
    layer.classList.add("is-closing");
    document.body.style.overflow = "";
    sheetCloseTimer = setTimeout(function () {
      sheetCloseTimer = null;
      layer.hidden = true;
      layer.classList.remove("is-closing");
      $("#sheetBody").innerHTML = "";
    }, 190);
    st.resolve(value);
  }

  function confirmSheet(opts) {
    return openSheet({
      title: opts.title || "确认",
      body: '<p style="font-size:12.5px;line-height:1.8;color:var(--ink-2);margin:2px 0 6px">' + (opts.message || "") + "</p>",
      foot: '<div style="display:flex;gap:8px">' +
        '<button class="btn btn--ghost" style="flex:1" data-sheet-cancel>取消</button>' +
        '<button class="btn ' + (opts.danger ? "btn--danger" : "btn--primary") + '" style="flex:1" data-sheet-ok>' +
        esc(opts.okText || "确定") + "</button></div>",
      onFoot: function (foot, close) {
        var cancel = $("[data-sheet-cancel]", foot);
        var ok = $("[data-sheet-ok]", foot);
        if (cancel) cancel.onclick = function () { close(false); };
        if (ok) ok.onclick = function () { close(true); };
      }
    });
  }

  $("#sheetBackdrop").addEventListener("click", function () { closeSheet(null); });
  $("#sheetClose").addEventListener("click", function () { closeSheet(null); });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && sheetState) closeSheet(null);
  });

  /* ============================== 数据装载 ============================== */
  async function boot() {
    await Store.open();

    // 1) 使用者档案
    S.profiles = await Store.listProfiles();
    if (!S.profiles.length) {
      var first = {
        id: "p-" + Date.now().toString(36),
        name: "我",
        grade: "",
        major: "",
        classKey: "",
        cls: "",
        starred: [],
        scope: "class",
        createdAt: new Date().toISOString()
      };
      await Store.saveProfile(first);
      S.profiles = [first];
    }
    var pid = await Store.getSetting("activeProfile", "");
    S.profile = S.profiles.find(function (p) { return p.id === pid; }) || S.profiles[0];
    if (!S.profile.starred) S.profile.starred = [];
    S.starred = new Set(S.profile.starred);
    S.scope = S.profile.scope || "class";

    S.entries = await Store.listEntries(S.profile.id);
    await loadWallpaper();
    initCredit();

    // 2) 总表列表（共享表 + 当前档案私有表）
    await refreshDatasets();

    // 3) 首次访问导入内置种子
    if (!S.datasets.length) {
      var seeded = await Store.getSetting("seeded", false);
      if (!seeded) {
        try { await importSeed(); } catch (e) { console.warn("[青简叙] 种子数据加载失败", e); }
        await refreshDatasets();
        await Store.setSetting("seeded", true);
      }
    }

    var aid = await Store.getSetting("activeDataset", "");
    S.activeId = findVisible(aid) ? aid : (S.datasets[0] ? S.datasets[0].id : null);

    // 4) URL 携带班级（同学分享链接）
    applyHashProfile();
    await loadActiveRecords();
    await restoreFilters();

    S.ready = true;
    renderChrome();
    renderView();
    registerSW();
  }

  function findVisible(id) {
    return S.datasets.some(function (d) { return d.id === id; });
  }

  async function refreshDatasets() {
    var all = await Store.listDatasets();
    S.datasets = all
      .filter(function (d) { return d.ownerId === "shared" || d.ownerId === S.profile.id; })
      .sort(function (a, b) {
        if (!!a.builtin !== !!b.builtin) return a.builtin ? -1 : 1;
        return String(b.importedAt).localeCompare(String(a.importedAt));
      });
  }

  async function importSeed() {
    var res = await fetch("assets/data/seed.json", { cache: "no-store" });
    if (!res.ok) throw new Error("种子数据缺失");
    var seed = await res.json();
    for (var i = 0; i < (seed.datasets || []).length; i++) {
      var item = seed.datasets[i];
      var meta = Object.assign({}, item.meta, {
        ownerId: "shared",
        builtin: true,
        importedAt: item.meta.importedAt || new Date().toISOString()
      });
      meta.id = meta.id || "ds-seed-" + i;
      var records = item.records.map(function (r, idx) {
        r.id = meta.id + "#" + (idx + 1);
        r.datasetId = meta.id;
        return r;
      });
      await Store.saveDataset(meta, records);
      S.cache.set(meta.id, records);
    }
  }

  async function loadActiveRecords() {
    if (!S.activeId) { S.records = []; return; }
    if (S.cache.has(S.activeId)) { S.records = S.cache.get(S.activeId); return; }
    var recs = await Store.getRecords(S.activeId);
    recs.sort(function (a, b) { return (a.date + a.start).localeCompare(b.date + b.start); });
    S.cache.set(S.activeId, recs);
    S.records = recs;
  }

  function activeDataset() {
    return S.datasets.find(function (d) { return d.id === S.activeId; }) || null;
  }

  function applyHashProfile() {
    var m = location.hash.match(/c=([^&]+)/);
    if (!m) return;
    var key = decodeURIComponent(m[1]);
    var parts = key.split("|");
    if (parts.length !== 3) return;
    S.profile.grade = parts[0];
    S.profile.major = parts[1];
    S.profile.classNo = parts[2];
    S.profile.classKey = key;
    S.profile.cls = parts[0] + parts[1] + parts[2] + "班";
    S.hashClass = key; // 分享链接优先于本机上次筛选
    Store.saveProfile(S.profile);
    history.replaceState(null, "", location.pathname + location.search);
  }

  function filterStorageKey() {
    return "filters:" + S.profile.id + ":" + (S.activeId || "none");
  }

  async function restoreFilters() {
    // 来自分享链接时，强制使用链接里的班级，避免被本机旧筛选覆盖
    if (S.hashClass) {
      var h = S.hashClass.split("|");
      S.filters = { grade: h[0], major: h[1], classKey: S.hashClass, q: "" };
      S.hashClass = "";
      persistFilters();
      var hs = $("#searchInput");
      if (hs) hs.value = "";
      return;
    }

    var saved = await Store.getSetting(filterStorageKey(), null);
    if (saved && typeof saved === "object") {
      S.filters = {
        grade: saved.grade || "",
        major: saved.major || "",
        classKey: saved.classKey || "",
        q: saved.q || ""
      };
      // 换总表后旧筛选可能已失效
      if (S.filters.classKey && S.records.length &&
          !S.records.some(function (r) { return r.classKey === S.filters.classKey; })) {
        S.filters = { grade: "", major: "", classKey: "", q: "" };
      }
    } else if (S.profile.classKey && S.records.some(function (r) { return r.classKey === S.profile.classKey; })) {
      var p = S.profile.classKey.split("|");
      S.filters = { grade: p[0], major: p[1], classKey: S.profile.classKey, q: "" };
    }
    var si = $("#searchInput");
    if (si) si.value = S.filters.q || "";
  }

  var saveFilterTimer = null;
  function persistFilters() {
    clearTimeout(saveFilterTimer);
    saveFilterTimer = setTimeout(function () {
      Store.setSetting(filterStorageKey(), S.filters);
      Store.saveProfile(S.profile);
    }, 260);
  }

  /* ============================== 作用范围 ============================== */
  function scopedRecords(scope, p) {
    p = p || S.profile;
    var list = S.records;
    if (scope === "class") {
      if (!p.classKey) return null;
      return list.filter(function (r) { return r.classKey === p.classKey; });
    }
    if (scope === "major") {
      if (!p.major) return null;
      return list.filter(function (r) { return r.major === p.major && (!p.grade || r.grade === p.grade); });
    }
    if (scope === "grade") {
      if (!p.grade) return null;
      return list.filter(function (r) { return r.grade === p.grade; });
    }
    if (scope === "starred") {
      var base = scopedRecords(p.classKey ? "class" : (p.major ? "major" : "grade"), p) || list;
      return base.filter(function (r) { return S.starred.has(r.course); });
    }
    return list.slice();
  }

  function scopeLabel(scope) {
    scope = scope || S.scope;
    var p = S.profile;
    if (scope === "class") return p.cls || "未设置班级";
    if (scope === "major") return (p.grade || "") + (p.major || "未设置专业");
    if (scope === "grade") return (p.grade || "未设置年级") + "全部专业";
    if (scope === "starred") return "我的关注";
    return "全部考程";
  }

  function availableScopes() {
    var p = S.profile;
    var out = [];
    if (p.classKey) out.push({ key: "class", label: "本班" });
    if (p.major) out.push({ key: "major", label: "本专业" });
    if (p.grade) out.push({ key: "grade", label: "本年级" });
    if (S.starred.size) out.push({ key: "starred", label: "关注" });
    out.push({ key: "all", label: "全部" });
    return out;
  }

  /* ===================== 手动录入的实验 / 考试（3.0） ===================== */
  function plusMinutes(hhmm, mins) {
    var p = String(hhmm || "").split(":");
    if (p.length < 2) return "";
    var total = (+p[0]) * 60 + (+p[1]) + mins;
    total = ((total % 1440) + 1440) % 1440;
    return pad(Math.floor(total / 60)) + ":" + pad(total % 60);
  }

  function entryToEvent(e) {
    var end = e.end || (e.start ? plusMinutes(e.start, 110) : "23:59");
    return {
      kind: "custom",
      id: e.id,
      entryId: e.id,
      course: e.title,
      date: e.date,
      start: e.start || "",
      end: end,
      rooms: e.place ? [e.place] : [],
      classes: [],
      headcount: 0,
      note: e.note || "",
      entryType: e.type || "exam",
      iso: e.date + "T" + end + ":00",
      key: "c:" + e.id
    };
  }

  /** 自建条目不受年级 / 专业 / 班级筛选影响，仅受搜索词约束 */
  function customEvents(q) {
    var kw = (q || "").trim().toLowerCase();
    return S.entries
      .filter(function (e) {
        if (!kw) return true;
        return [e.title, e.place, e.note].join(" ").toLowerCase().indexOf(kw) > -1;
      })
      .map(entryToEvent);
  }

  function sortEvents(list) {
    return list.sort(function (a, b) {
      var d = (a.date + (a.start || "99:99")).localeCompare(b.date + (b.start || "99:99"));
      return d || a.course.localeCompare(b.course, "zh");
    });
  }

  async function reloadEntries() {
    S.entries = await Store.listEntries(S.profile.id);
  }

  function openEntryForm(id) {
    var editing = id ? S.entries.find(function (e) { return e.id === id; }) : null;
    var cur = editing || {
      type: "lab", title: "", date: todayISO(), start: "", end: "", place: "", note: ""
    };
    var isExam = cur.type === "exam";

    var body =
      '<div class="field"><label>类别</label><div class="choice-row" id="entryType">' +
      '<button type="button" class="choice' + (!isExam ? " is-active" : "") + '" data-type="lab">实验</button>' +
      '<button type="button" class="choice' + (isExam ? " is-active" : "") + '" data-type="exam">考试</button>' +
      "</div></div>" +
      '<div class="field"><label for="etTitle">名称 <span style="color:var(--cinnabar)">*</span></label>' +
      '<input id="etTitle" maxlength="40" placeholder="如：生理学实验（三）" value="' + esc(cur.title) + '" /></div>' +
      '<div class="field"><label for="etDate">日期 <span style="color:var(--cinnabar)">*</span></label>' +
      '<input id="etDate" type="date" value="' + esc(cur.date) + '" /></div>' +
      '<div class="field" style="display:flex;gap:9px">' +
      '<div style="flex:1;min-width:0"><label for="etStart">开始时间</label><input id="etStart" type="time" value="' + esc(cur.start) + '" /></div>' +
      '<div style="flex:1;min-width:0"><label for="etEnd">结束时间</label><input id="etEnd" type="time" value="' + esc(cur.end) + '" /></div>' +
      "</div>" +
      '<div class="field"><label for="etPlace">地点</label>' +
      '<input id="etPlace" maxlength="40" placeholder="如：教学楼 A305 / 实验楼 2 栋" value="' + esc(cur.place) + '" /></div>' +
      '<div class="field"><label for="etNote">备注</label>' +
      '<textarea id="etNote" maxlength="200" rows="3" placeholder="如：需穿白大褂，提前 15 分钟到场">' + esc(cur.note) + "</textarea></div>" +
      '<p class="hint" style="margin-top:-4px">自建条目只保存在本机，会与总表里的考试合并成一条考程时间轴。</p>';

    var foot = '<div style="display:flex;gap:8px">' +
      (editing ? '<button class="btn btn--danger" style="flex:0 0 auto" data-entry-del>删除</button>' : "") +
      '<button class="btn btn--primary" style="flex:1" data-entry-save>' + (editing ? "保存修改" : "记入青简") + "</button></div>";

    var chosenType = cur.type;

    openSheet({
      title: editing ? "编辑条目" : "手动录入",
      body: body,
      foot: foot,
      onMount: function (root) {
        root.addEventListener("click", function (e) {
          var c = e.target.closest("[data-type]");
          if (!c) return;
          chosenType = c.dataset.type;
          $$(".choice", root).forEach(function (x) { x.classList.remove("is-active"); });
          c.classList.add("is-active");
        });
      },
      onFoot: function (footEl, close) {
        var saveBtn = $("[data-entry-save]", footEl);
        var delBtn = $("[data-entry-del]", footEl);

        if (delBtn) {
          delBtn.onclick = async function () {
            var ok = await confirmSheet({
              title: "删除条目",
              message: "将删除「" + esc(cur.title) + "」，不可撤销。",
              okText: "删除", danger: true
            });
            if (!ok) return;
            await Store.deleteEntry(cur.id);
            await reloadEntries();
            close(null);
            renderView();
            toast("已删除");
          };
        }

        saveBtn.onclick = async function () {
          var title = $("#etTitle").value.trim();
          var date = $("#etDate").value;
          var start = $("#etStart").value;
          var end = $("#etEnd").value;
          var place = $("#etPlace").value.trim();
          var note = $("#etNote").value.trim();

          if (!title) { toast("请填写名称", "error"); return; }
          if (!date) { toast("请选择日期", "error"); return; }
          if (start && end && end <= start) { toast("结束时间需晚于开始时间", "error"); return; }

          var entry = {
            id: cur.id || "e-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
            ownerId: S.profile.id,
            type: chosenType,
            title: title,
            date: date,
            start: start,
            end: end,
            place: place,
            note: note,
            createdAt: cur.createdAt || new Date().toISOString(),
            updatedAt: new Date().toISOString()
          };
          await Store.saveEntry(entry);
          await reloadEntries();
          close(true);
          renderView();
          toast(editing ? "已保存修改" : "已记入「" + title + "」");
        };
      }
    });
  }

  /* ============================ 倒数日壁纸（3.0） ============================ */
  function wallpaperKey() { return "wallpaper:" + S.profile.id; }

  async function loadWallpaper() {
    var wp = await Store.getSetting(wallpaperKey(), null);
    S.wallpaper = wp && wp.url ? wp : null;
    applyWallpaper();
  }

  function applyWallpaper() {
    var layer = $("#wallpaperLayer");
    var img = $("#wallpaperImg");
    var scrim = $("#wallpaperScrim");
    if (!layer || !img || !scrim) return;

    if (!S.wallpaper || !S.wallpaper.url) {
      layer.hidden = true;
      document.body.classList.remove("has-wallpaper", "show-wallpaper", "wallpaper-light");
      return;
    }
    layer.hidden = false;
    img.style.backgroundImage = 'url("' + S.wallpaper.url + '")';
    var blur = Number(S.wallpaper.blur) || 0;
    img.style.filter = blur ? "blur(" + blur + "px)" : "none";
    img.style.transform = blur ? "scale(" + (1 + Math.min(0.08, blur / 180)).toFixed(3) + ")" : "none";
    scrim.style.setProperty("--wp-scrim", String(S.wallpaper.scrim == null ? 0.3 : S.wallpaper.scrim));
    document.body.classList.add("has-wallpaper");
    document.body.classList.toggle("wallpaper-light", !!S.wallpaper.light);
    applyWallpaperVisibility();
  }

  function applyWallpaperVisibility() {
    var layer = $("#wallpaperLayer");
    var on = !!(S.wallpaper && S.wallpaper.url) && S.view === "countdown";
    document.body.classList.toggle("show-wallpaper", on);
    if (layer && !layer.hidden) layer.classList.toggle("is-off", !on);
  }

  async function saveWallpaper(patch) {
    S.wallpaper = Object.assign({}, S.wallpaper || {}, patch);
    if (!S.wallpaper.url) {
      await Store.setSetting(wallpaperKey(), null);
    } else {
      await Store.setSetting(wallpaperKey(), {
        url: S.wallpaper.url,
        light: !!S.wallpaper.light,
        blur: Number(S.wallpaper.blur) || 0,
        scrim: S.wallpaper.scrim == null ? 0.3 : Number(S.wallpaper.scrim)
      });
    }
    applyWallpaper();
  }

  /** 本地压缩：最长边 1440px，JPEG 0.82；同时采样亮度，决定壁纸上的文字用深墨还是米色 */
  function compressImage(file) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function () {
        try {
          var max = 1440;
          var w = img.naturalWidth || img.width;
          var h = img.naturalHeight || img.height;
          var scale = Math.min(1, max / Math.max(w, h));
          var cw = Math.max(1, Math.round(w * scale));
          var ch = Math.max(1, Math.round(h * scale));
          var canvas = document.createElement("canvas");
          canvas.width = cw; canvas.height = ch;
          var ctx = canvas.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, cw, ch);
          ctx.drawImage(img, 0, 0, cw, ch);
          var out = canvas.toDataURL("image/jpeg", 0.82);

          // 24×24 缩略图算平均亮度，足够判断整体明暗
          var sample = document.createElement("canvas");
          sample.width = 24; sample.height = 24;
          var sctx = sample.getContext("2d");
          sctx.drawImage(canvas, 0, 0, 24, 24);
          var pixels = sctx.getImageData(0, 0, 24, 24).data;
          var sum = 0;
          for (var i = 0; i < pixels.length; i += 4) {
            sum += 0.2126 * pixels[i] + 0.7152 * pixels[i + 1] + 0.0722 * pixels[i + 2];
          }
          var light = sum / (24 * 24) > 150;

          URL.revokeObjectURL(url);
          resolve({ url: out, light: light });
        } catch (err) {
          URL.revokeObjectURL(url);
          reject(err);
        }
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("图片无法读取，请换一张试试"));
      };
      img.src = url;
    });
  }

  function openWallpaperSheet() {
    var has = !!(S.wallpaper && S.wallpaper.url);
    var scrimPct = Math.round((S.wallpaper && S.wallpaper.scrim != null ? Number(S.wallpaper.scrim) : 0.3) * 100);
    var body =
      '<div class="wp-preview" id="wpPreview">' + (has ? "" : "<span>还没有设置壁纸</span>") + "</div>" +
      '<div style="display:flex;gap:8px;margin:10px 0 4px">' +
      '<button class="btn btn--primary" style="flex:1" data-wp-pick>' + (has ? "更换图片" : "选择图片") + "</button>" +
      (has ? '<button class="btn btn--ghost" data-wp-clear>移除</button>' : "") +
      "</div>" +
      '<div class="field" style="margin-top:14px"><label>模糊程度 <span id="wpBlurVal" style="color:var(--faint);font-weight:400">' +
      (Number(S.wallpaper && S.wallpaper.blur) || 0) + " px</span></label>" +
      '<input id="wpBlur" type="range" min="0" max="20" step="1" value="' + (Number(S.wallpaper && S.wallpaper.blur) || 0) + '" /></div>' +
      '<div class="field"><label>遮罩强度 <span id="wpScrimVal" style="color:var(--faint);font-weight:400">' + scrimPct + "%</span></label>" +
      '<input id="wpScrim" type="range" min="0" max="90" step="5" value="' + scrimPct + '" /></div>' +
      '<p class="hint">图片只保存在本机，不会上传。遮罩用于保证任何壁纸上的文字都清晰可读；仅作用于倒数日页。</p>';

    openSheet({
      title: "倒数日壁纸",
      body: body,
      onMount: function (root, close) {
        var preview = $("#wpPreview", root);
        if (has && preview) preview.style.backgroundImage = 'url("' + S.wallpaper.url + '")';

        var input = $("#wallpaperInput");
        $("[data-wp-pick]", root).onclick = function () { input.click(); };

        var clearBtn = $("[data-wp-clear]", root);
        if (clearBtn) {
          clearBtn.onclick = async function () {
            await saveWallpaper({ url: "", blur: 0, scrim: 0.3 });
            S.wallpaper = null;
            close(null);
            renderView();
            toast("已移除壁纸");
          };
        }

        var blur = $("#wpBlur", root);
        var scrim = $("#wpScrim", root);

        blur.addEventListener("input", function () {
          $("#wpBlurVal", root).textContent = blur.value + " px";
          if (S.wallpaper && S.wallpaper.url) {
            S.wallpaper.blur = Number(blur.value);
            applyWallpaper();
          }
        });
        blur.addEventListener("change", function () {
          if (S.wallpaper && S.wallpaper.url) saveWallpaper({ blur: Number(blur.value) });
        });

        scrim.addEventListener("input", function () {
          $("#wpScrimVal", root).textContent = scrim.value + "%";
          if (S.wallpaper && S.wallpaper.url) {
            S.wallpaper.scrim = Number(scrim.value) / 100;
            applyWallpaper();
          }
        });
        scrim.addEventListener("change", function () {
          if (S.wallpaper && S.wallpaper.url) saveWallpaper({ scrim: Number(scrim.value) / 100 });
        });
      }
    });
  }

  // 选图后立即压缩、保存、应用
  $("#wallpaperInput").addEventListener("change", async function (e) {
    var file = (e.target.files || [])[0];
    e.target.value = "";
    if (!file) return;
    if (!/^image\//.test(file.type)) { toast("请选择图片文件", "error"); return; }
    if (file.size > 12 * 1024 * 1024) { toast("图片超过 12MB，请先压缩", "error"); return; }

    loading(true, "正在铺开壁纸…");
    try {
      var shot = await compressImage(file);
      var prev = S.wallpaper || {};
      await saveWallpaper({
        url: shot.url,
        light: shot.light,
        blur: prev.blur == null ? 0 : prev.blur,
        scrim: prev.scrim == null ? 0.3 : prev.scrim
      });
      loading(false);
      closeSheet(true);
      renderView();
      toast("壁纸已应用");
    } catch (err) {
      loading(false);
      toast(err.message || "壁纸处理失败", "error");
    }
  });

  /** 制作者签名：图片加载失败时退回文字，保证标识始终可见 */
  function initCredit() {
    var img = $("#creditSign");
    var text = $("#creditText");
    if (!img || !text) return;
    function showImage() { img.hidden = false; text.hidden = true; }
    function showText() { img.style.display = "none"; text.hidden = false; }
    img.addEventListener("load", showImage);
    img.addEventListener("error", showText);
    if (img.complete) {
      if (img.naturalWidth > 0) showImage(); else showText();
    }
  }

  /* ============================== 筛选逻辑 ============================== */
  function filteredRecords() {
    var f = S.filters;
    var q = f.q.trim().toLowerCase();
    return S.records.filter(function (r) {
      if (f.grade && r.grade !== f.grade) return false;
      if (f.major && r.major !== f.major) return false;
      if (f.classKey && r.classKey !== f.classKey) return false;
      if (q) {
        var hay = [r.course, r.code, r.cls, r.college, (r.rooms || []).join(" "), r.teaching, r.major]
          .join(" ").toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }

  function filterScopeLabel() {
    var f = S.filters;
    var parts = [];
    if (f.grade) parts.push(f.grade);
    if (f.major) parts.push(f.major);
    if (f.classKey) parts.push(f.classKey.split("|")[2] + "班");
    if (f.q.trim()) parts.push("搜索「" + f.q.trim() + "」");
    return parts.join(" · ") || "全部考程";
  }

  /* ================================ 渲染 ================================ */
  function renderChrome() {
    var ds = activeDataset();
    if (ds) {
      $("#datasetName").textContent = ds.title;
      var s = ds.summary || {};
      $("#datasetMeta").textContent = (s.recordCount || 0) + " 条 · " + (s.courseCount || 0) + " 门";
    } else {
      $("#datasetName").textContent = "尚未载入总表";
      $("#datasetMeta").textContent = "";
    }
    var p = S.profile;
    $("#avatarText").textContent = (p.name || "我").slice(0, 1);
    $$(".tab").forEach(function (t) {
      var on = t.dataset.view === S.view;
      t.classList.toggle("is-active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
    });
  }

  function renderView() {
    $("#view-countdown").hidden = S.view !== "countdown";
    $("#view-exams").hidden = S.view !== "exams";
    $("#view-me").hidden = S.view !== "me";
    if (S.view === "countdown") renderCountdown();
    if (S.view === "exams") renderExams();
    if (S.view === "me") renderMe();
    applyWallpaperVisibility();
    window.scrollTo(0, 0);
  }

  function emptyState(title, text, btnLabel, act) {
    return '<div class="empty"><div class="empty__icon">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' +
      '<rect x="3.5" y="4.5" width="17" height="16" rx="3"/><path d="M8 3v4M16 3v4M3.5 9.5h17"/>' +
      '<path d="M8 14h3M8 17h6"/></svg></div>' +
      "<h3>" + esc(title) + "</h3><p>" + text + "</p>" +
      (btnLabel ? '<button class="btn btn--primary" data-act="' + act + '">' + esc(btnLabel) + "</button>" : "") +
      "</div>";
  }

  function noDataHTML() {
    return emptyState(
      "还没有载入考试总表",
      "导入教务下发的 Excel 安排表，即可按班级筛选、查看倒数日并导出结果；也可以先手动录入一条实验或考试。",
      "导入总表",
      "import"
    );
  }

  /* ------------------------------ 倒数日视图 ------------------------------ */
  /**
   * 把「班级 × 课程」明细聚合成「考试场次」：同一课程、同一天、同一起考时间视为一场。
   * 本专业 / 本年级视角下同一门课会因多个班级重复出现，聚合后才符合「一场考试」的直觉。
   */
  function groupExams(records) {
    var map = new Map();
    records.forEach(function (r) {
      var k = r.course + "|" + r.date + "|" + (r.start || "");
      var g = map.get(k);
      if (!g) {
        g = {
          key: k, course: r.course, code: r.code, date: r.date, start: r.start, end: r.end,
          rooms: [], classes: [], headcount: 0, iso: r.iso
        };
        map.set(k, g);
      }
      (r.rooms || []).forEach(function (x) { if (g.rooms.indexOf(x) === -1) g.rooms.push(x); });
      if (r.cls && g.classes.indexOf(r.cls) === -1) g.classes.push(r.cls);
      g.headcount += r.headcount || 0;
      if (Date.parse(r.iso) > Date.parse(g.iso)) g.iso = r.iso;
    });
    return Array.from(map.values()).sort(function (a, b) {
      var d = (a.date + a.start).localeCompare(b.date + b.start);
      return d || a.course.localeCompare(b.course, "zh");
    });
  }

  function renderCountdown() {
    var root = $("#countdownBody");

    if (!S.records.length && !S.entries.length) { root.innerHTML = noDataHTML(); return; }

    var p = S.profile;
    var needsSetup = !p.classKey && !p.major && !p.grade;

    if (needsSetup && !S.entries.length) { root.innerHTML = onboardingHTML(); return; }

    var scopes = availableScopes();
    if (!scopes.some(function (s) { return s.key === S.scope; })) S.scope = scopes[0].key;

    var scoped = scopedRecords(S.scope) || [];
    S.scopeRecords = scoped;
    var groups = groupExams(scoped);
    S.groups = groups.reduce(function (acc, g) { acc[g.key] = g; return acc; }, {});

    // 总表场次与手动录入条目合并成一条时间轴
    var events = sortEvents(groups.concat(customEvents()));
    S.events = events;

    var now = Date.now();
    var upcoming = events.filter(function (g) { return Date.parse(g.iso) >= now; });
    var past = events.filter(function (g) { return Date.parse(g.iso) < now; }).reverse();

    var html = toolbarHTML();

    if (upcoming.length) {
      html += heroHTML(upcoming[0], events.indexOf(upcoming[0]) + 1, events.length, events[events.length - 1].date);
    } else if (events.length) {
      html += '<div class="hero"><div class="hero__top"><span class="hero__scope">' + esc(scopeLabel()) +
        '</span><span class="hero__seq">共 ' + events.length + " 场</span></div>" +
        '<div class="hero__label">这一程已经走完</div>' +
        '<div class="hero__course" style="margin-top:6px">辛苦了，注意休息</div>' +
        '<div class="hero__foot"><span>本范围暂无待考科目</span><span>已结束 ' + past.length + " 场</span></div></div>";
    } else {
      html += '<div class="hero"><div class="hero__top"><span class="hero__scope">' + esc(scopeLabel()) + '</span></div>' +
        '<div class="hero__label">此范围暂无考程</div>' +
        '<div class="hero__course">换个范围，或手动记下一条实验 / 考试</div>' +
        '<div class="hero__foot"><span>青简只记与这一程有关的事</span></div></div>';
    }

    // 范围切换
    html += '<div class="seg" role="tablist">' + scopes.map(function (s) {
      return '<button type="button" data-act="scope" data-scope="' + s.key + '" class="' +
        (S.scope === s.key ? "is-active" : "") + '">' + esc(s.label) + "</button>";
    }).join("") + "</div>";

    html += '<div class="section-head section-head--action">' +
      "<h3>即将开考<span>" + upcoming.length + " 场待考</span></h3>" +
      '<button class="link-btn" type="button" data-act="add-entry">＋ 录入实验/考试</button></div>';

    if (!upcoming.length) {
      html += '<div class="cd-card"><div class="empty" style="padding:28px 20px">' +
        '<h3 style="font-size:13.5px">该范围暂无待考科目</h3>' +
        '<p style="font-size:11px">' + esc(scopeLabel()) + " 的考试均已结束或未安排。<br/>可以手动录入尚未公布时间的实验与考试。</p>" +
        '<button class="btn btn--soft btn--sm" style="margin-top:12px" data-act="add-entry">录入一条</button>' +
        "</div></div>";
    } else {
      html += '<div class="cd-card">' + upcoming.slice(0, 40).map(function (g) {
        return cdItemHTML(g, false);
      }).join("") + "</div>";
      if (upcoming.length > 40) {
        html += '<p class="me-note">仅展开最近 40 场，共 ' + upcoming.length + " 场</p>";
      }
    }

    if (past.length) {
      html += '<button class="collapse-head' + (S.pastOpen ? " is-open" : "") + '" data-act="toggle-past">' +
        "<span>已结束 · " + past.length + " 场</span>" +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 9l6 6 6-6"/></svg>' +
        "</button>";
      if (S.pastOpen) {
        html += '<div class="cd-card" style="opacity:.78">' + past.map(function (g) {
          return cdItemHTML(g, true);
        }).join("") + "</div>";
      }
    }

    html += '<p class="me-note">倒数日按结束时间计算 · 自建条目仅存本机 · 数据不出设备</p>';
    root.innerHTML = html;
  }


  function toolbarHTML() {
    var label = S.wallpaper && S.wallpaper.url ? "更换壁纸" : "壁纸";
    return '<div class="cd-toolbar">' +
      '<span class="cd-toolbar__hint">' + esc(fullDateLabel(todayISO())) + " · " + weekdayOf(todayISO()) + "</span>" +
      '<button class="ghost-chip' + (S.wallpaper && S.wallpaper.url ? " is-on" : "") + '" type="button" data-act="wallpaper">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
      '<rect x="3" y="4.5" width="18" height="15" rx="3"/><circle cx="8.6" cy="9.6" r="1.4"/>' +
      '<path d="M4 17l4.6-4.3 3.6 3.2 3-2.6L20 18"/></svg>' + label + "</button>" +
      "</div>";
  }

  function onboardingHTML() {
    var all = groupExams(S.records);
    var days = {};
    all.forEach(function (g) { days[g.date] = 1; });
    var first = all[0] ? all[0].date : "";
    var last = all.length ? all[all.length - 1].date : "";
    var ds = activeDataset();
    var userIcon = '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="#fffdf7" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">' +
      '<circle cx="12" cy="8.4" r="3.4"/><path d="M5.4 19.6c.9-3.4 3.4-5.2 6.6-5.2s5.7 1.8 6.6 5.2"/></svg>';

    return toolbarHTML() + '<div class="profile-hero">' +
      '<div class="profile-hero__avatar" style="background:linear-gradient(150deg,#c08a4a,#a8452f)">' + userIcon + "</div>" +
      '<div class="profile-hero__text"><h2>先写下你的班级</h2>' +
      "<p>此后青简只记与你有关的考程与倒数日</p></div>" +
      '<button class="btn btn--primary btn--sm" data-act="edit-profile">设置</button></div>' +

      '<div class="cd-card" style="padding:16px 16px 14px;margin-top:12px">' +
      '<h3 style="margin:0 0 12px;font-size:13px;font-family:var(--serif);letter-spacing:.5px">' + esc(ds ? ds.title : "本学期考程") + "</h3>" +
      '<div class="stat-grid">' +
      '<div><b>' + all.length + "</b><span>场考试</span></div>" +
      '<div><b>' + Object.keys(days).length + "</b><span>个考试日</span></div>" +
      '<div><b>' + (ds && ds.summary ? ds.summary.classCount : 0) + "</b><span>个班级</span></div>" +
      "</div>" +
      '<p style="margin:12px 0 0;font-size:10px;color:var(--faint);line-height:1.7">' +
      "考试区间 " + esc(first) + " ~ " + esc(last) + "</p>" +
      '<button class="btn btn--ghost btn--block" style="margin-top:14px" data-act="browse-all">先翻阅全部考程</button>' +
      '<button class="btn btn--soft btn--block" style="margin-top:8px" data-act="add-entry">手动录入实验 / 考试</button>' +
      "</div>" +

      '<div class="step-card">' +
      '<h3>三步，记下这一程</h3>' +
      '<ol>' +
      "<li><b>1</b><span>设置班级，只看与自己有关的考试</span></li>" +
      "<li><b>2</b><span>筛选、检索，或补录未进总表的实验与考试</span></li>" +
      "<li><b>3</b><span>导出 Excel / PDF / CSV / Word，或把链接分享给同学</span></li>" +
      "</ol></div>" +

      '<p class="me-note">设置班级后，这里会变成你的专属考程倒数日。</p>';
  }

  function heroHTML(next, seq, total, lastDate) {
    var days = daysUntil(next.date);
    var dayText = days === 0 ? "今天" : (days === 1 ? "明天" : (days < 0 ? "已结束" : days));
    var numHTML = days === 0 || days === 1
      ? '<span class="hero__num hero__num--text">' + dayText + "</span>"
      : '<span class="hero__num">' + days + "</span>";
    var unitHTML = (days === 0 || days === 1) ? "" : '<span class="hero__unit">天</span>';
    var pct = Math.max(4, Math.min(100, Math.round(((seq - 1) / Math.max(total, 1)) * 100)));
    var who = next.kind === "custom"
      ? (next.entryType === "lab" ? "自建实验" : "自建考试")
      : (next.classes.length > 1 ? next.classes.length + " 个班合考" : (next.classes[0] || "多班合考"));

    return '<div class="hero">' +
      '<div class="hero__top">' +
      '<span class="hero__scope">' + esc(scopeLabel()) + "</span>" +
      '<span class="hero__seq">第 ' + seq + " / " + total + " 场</span>" +
      "</div>" +
      '<div class="hero__label">距离下一场，还有</div>' +
      '<div class="hero__num-row">' + numHTML + unitHTML +
      '<div class="hero__when"><b>' + esc(fullDateLabel(next.date)) + " " + weekdayOf(next.date) + "</b>" +
      "<span>" + esc(timeRange(next)) + "</span></div></div>" +
      '<div class="hero__course">' +
      (next.kind === "custom" ? '<span class="tag tag--' + (next.entryType === "lab" ? "lab" : "exam") + '">' + (next.entryType === "lab" ? "实验" : "考试") + "</span>" : "") +
      esc(next.course) + "</div>" +
      '<div class="hero__bar"><i style="width:' + pct + '%"></i></div>' +
      '<div class="hero__foot"><span>' + esc(roomsText(next)) + "</span><span>" +
      (lastDate ? "考程至 " + esc(mdLabel(lastDate)) : "共 " + total + " 场") + "</span></div>" +
      "</div>";
  }

  function cdItemHTML(g, isPast) {
    var days = daysUntil(g.date);
    var cls = "cd-days";
    if (isPast) cls += " cd-days--past";
    else if (days <= 3) cls += " cd-days--hot";
    else if (days <= 7) cls += " cd-days--warm";
    else if (days <= 14) cls += " cd-days--cool";

    var isCustom = g.kind === "custom";
    var extra = isCustom
      ? (g.note ? " · " + g.note : "")
      : (g.classes.length > 1 ? " · " + g.classes.length + " 个班" : "");
    var action = isCustom
      ? 'data-act="edit-entry" data-id="' + esc(g.entryId) + '"'
      : 'data-act="gdetail" data-g="' + encodeURIComponent(g.key) + '"';
    var tag = isCustom
      ? '<span class="tag tag--' + (g.entryType === "lab" ? "lab" : "exam") + '">' + (g.entryType === "lab" ? "实验" : "考试") + "</span>"
      : "";

    return '<button class="cd-item" ' + action + ' style="width:100%;text-align:left">' +
      '<div class="cd-item__date"><b>' + mdLabel(g.date) + "</b><span>" + weekdayOf(g.date) + "</span></div>" +
      '<div class="cd-item__body"><h4>' + tag + esc(g.course) + "</h4><p>" +
      esc(timeRange(g)) + " · " + esc(roomsText(g)) + esc(extra) + "</p></div>" +
      '<div class="' + cls + '"><b>' + (isPast ? "✓" : days) + "</b><span>" + (isPast ? "已考" : "天") + "</span></div>" +
      "</button>";
  }

  /* ------------------------------- 考程表视图 ------------------------------- */
  function renderExams() {
    var list = filteredRecords();
    var custom = customEvents(S.filters.q);

    var byDate = {};
    list.forEach(function (r) { (byDate[r.date] = byDate[r.date] || []).push(r); });
    custom.forEach(function (e) { (byDate[e.date] = byDate[e.date] || []).push(e); });
    var order = Object.keys(byDate).sort();
    order.forEach(function (d) {
      byDate[d].sort(function (a, b) {
        return (a.start || "99:99").localeCompare(b.start || "99:99");
      });
    });

    var total = list.length + custom.length;
    $("#resultCount").textContent = total;
    var courses = uniqueBy(list.concat(custom), function (r) { return r.course; }).length;
    $("#resultSub").textContent = total
      ? " · " + courses + " 个科目" + (custom.length ? " · 含 " + custom.length + " 条自建" : "")
      : "";
    $("#exportBtn").disabled = !total;

    var f = S.filters;
    setChip("grade", f.grade || "全部", !!f.grade);
    setChip("major", f.major || "全部", !!f.major);
    setChip("class", f.classKey ? f.classKey.split("|")[2] + "班" : "全部", !!f.classKey);
    $("#resetFilter").hidden = !(f.grade || f.major || f.classKey || f.q);
    $("#searchClear").hidden = !f.q;

    if (!S.records.length && !S.entries.length) { $("#examList").innerHTML = noDataHTML(); return; }
    if (!total) {
      $("#examList").innerHTML = emptyState(
        "没有匹配的考程",
        "试试放宽筛选条件，或换一个班级。当前范围：" + esc(filterScopeLabel()) + "。",
        "清空筛选",
        "reset-filter"
      );
      return;
    }

    var now = Date.now();
    var limit = S.listLimit || 120;
    var shown = 0;
    var html = "";

    for (var i = 0; i < order.length; i++) {
      if (shown >= limit) break;
      var date = order[i];
      var items = byDate[date];
      var take = items.slice(0, limit - shown);
      shown += take.length;

      var allPast = take.every(function (x) { return Date.parse(x.iso) < now; });
      var todayMark = date === todayISO() ? " · 今天" : "";
      html += '<div class="day-head' + (allPast ? " is-past" : "") + '">' +
        "<b>" + esc(fullDateLabel(date)) + "</b><em>" + weekdayOf(date) + todayMark + "</em>" +
        (items.length > 1 ? "<i>" + items.length + " 场</i>" : "") + "</div>" +
        '<div class="exam-list">' + take.map(function (r) {
          return r.kind === "custom"
            ? customItemHTML(r, Date.parse(r.iso) < now)
            : examItemHTML(r, Date.parse(r.iso) < now);
        }).join("") + "</div>";
    }

    if (shown < total) {
      html += '<button class="btn btn--ghost btn--block" style="margin:16px 16px 2px;width:calc(100% - 32px)" data-act="more">' +
        "继续翻阅（剩余 " + (total - shown) + " 场）</button>";
    }

    $("#examList").innerHTML = html;
  }

  function setChip(key, value, active) {
    var el = $('.fchip[data-filter="' + key + '"]');
    if (!el) return;
    $(".fchip__value", el).textContent = value;
    el.classList.toggle("is-set", !!active);
  }

  function examItemHTML(r, isPast) {
    var days = daysUntil(r.date);
    var starred = S.starred.has(r.course);
    var dayTag = days === 0 ? "今天" : (days > 0 ? days + " 天" : "已考");
    return '<article class="exam-item' + (isPast ? " is-past" : "") + '">' +
      '<div class="exam-time"><b>' + esc(r.start || "--:--") + "</b>" +
      "<span>" + esc(r.end || "") + "</span>" +
      (isPast ? "" : '<em class="' + (days === 0 ? "is-today" : "") + '">' + dayTag + "</em>") +
      "</div>" +
      '<div class="exam-body"><h4>' + esc(r.course) + "</h4>" +
      '<div class="exam-meta"><span class="room">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s7-6.1 7-11a7 7 0 1 0-14 0c0 4.9 7 11 7 11z"/><circle cx="12" cy="10" r="2.4"/></svg>' +
      esc(roomsText(r)) + "</span>" +
      "<span>" + esc(r.cls || "多班合考") + "</span>" +
      "</div></div>" +
      '<button class="star-btn' + (starred ? " is-on" : "") + '" data-act="star" data-course="' + esc(r.course) +
      '" aria-label="' + (starred ? "取消关注" : "关注") + esc(r.course) + '" aria-pressed="' + (starred ? "true" : "false") + '">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round">' +
      '<path d="M12 3.6l2.6 5.3 5.9.9-4.3 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8-4.3-4.1 5.9-.9z"/></svg>' +
      "</button></article>";
  }

  /** 自建条目卡片：与总表场次同一视觉体系，用标签区分来源 */
  function customItemHTML(e, isPast) {
    var days = daysUntil(e.date);
    var dayTag = days === 0 ? "今天" : (days > 0 ? days + " 天" : "已考");
    var label = e.entryType === "lab" ? "实验" : "考试";
    return '<article class="exam-item exam-item--custom' + (isPast ? " is-past" : "") + '">' +
      '<div class="exam-time"><b>' + esc(e.start || "--:--") + "</b>" +
      "<span>" + esc(e.start ? e.end : "") + "</span>" +
      (isPast ? "" : '<em class="' + (days === 0 ? "is-today" : "") + '">' + dayTag + "</em>") +
      "</div>" +
      '<div class="exam-body"><h4><span class="tag tag--' + (e.entryType === "lab" ? "lab" : "exam") + '">' + label + "</span>" + esc(e.course) + "</h4>" +
      '<div class="exam-meta"><span class="room">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s7-6.1 7-11a7 7 0 1 0-14 0c0 4.9 7 11 7 11z"/><circle cx="12" cy="10" r="2.4"/></svg>' +
      esc(roomsText(e)) + "</span>" +
      (e.note ? "<span>" + esc(e.note) + "</span>" : "") +
      "</div></div>" +
      '<button class="star-btn star-btn--edit" data-act="edit-entry" data-id="' + esc(e.entryId) +
      '" aria-label="编辑' + esc(e.course) + '">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M4 20h4l10.5-10.5a2.1 2.1 0 0 0-3-3L5 17v3z"/><path d="M13.5 6.5l3 3"/></svg>' +
      "</button></article>";
  }

  /* -------------------------------- 我的视图 -------------------------------- */
  function renderMe() {
    var p = S.profile;
    var ds = activeDataset();
    var clsLine = p.classKey
      ? [p.grade, p.major, p.classNo + "班"].filter(Boolean).join(" · ")
      : "尚未设置年级 / 专业 / 班级";

    var html = '<div class="profile-hero">' +
      '<div class="profile-hero__avatar">' + esc((p.name || "我").slice(0, 1)) + "</div>" +
      '<div class="profile-hero__text"><h2>' + esc(p.name || "我") + "</h2><p>" + esc(clsLine) + "</p></div>" +
      '<button class="btn btn--soft btn--sm" data-act="edit-profile">编辑</button></div>';

    html += '<div class="menu-card">' +
      menuRow("users", "切换使用者", "同机多人使用，配置互不干扰", S.profiles.length + " 个档案", "switch-profile") +
      menuRow("database", "总表管理", ds ? (ds.summary.recordCount + " 条 · " + ds.summary.classCount + " 个班级") : "尚未载入", ds ? "当前" : "", "manage-dataset") +
      menuRow("upload", "导入考试总表", "支持 xlsx / xls / csv / json 数据包", "", "import") +
      "</div>";

    html += '<div class="menu-card">' +
      menuRow("plus", "手动录入实验 / 考试", "名称、时间、地点、备注，与总表合并展示",
        S.entries.length ? S.entries.length + " 条" : "未录入", "add-entry") +
      menuRow("image", "倒数日壁纸", S.wallpaper ? "可调模糊与遮罩，仅本机保存" : "导入自己的图片作为倒数日背景",
        S.wallpaper ? "已设置" : "未设置", "wallpaper") +
      "</div>";

    html += '<div class="menu-card">' +
      menuRow("share", "分享给同学", "链接自动带上你的班级筛选", "", "share") +
      menuRow("package", "导出数据包", "把整份总表发给同学离线导入", "", "export-bundle") +
      menuRow("star", "已关注课程", "倒数日里优先查看这些科目", S.starred.size ? S.starred.size + " 门" : "未关注", "manage-star") +
      "</div>";

    html += '<div class="menu-card">' +
      menuRow("info", "关于与数据说明", "数据来源、隐私、版本与制作者", "", "about") +
      menuRow("trash", "清空本机数据", "删除所有总表与本地配置", "", "wipe", true) +
      "</div>";

    html += '<p class="me-note">所有数据仅保存在本机浏览器（IndexedDB），不上传任何服务器。<br/>' +
      "换设备或清除浏览器数据后需重新导入。</p>";

    $("#meBody").innerHTML = html;
  }

  function menuRow(icon, title, sub, tail, act, danger) {
    var icons = {
      users: '<circle cx="9" cy="8" r="3.2"/><path d="M3.5 19c.7-3.2 3-5 5.5-5s4.8 1.8 5.5 5"/><path d="M16 5.4a3 3 0 0 1 0 5.6M17.5 14.4c2 .6 3.3 2.3 3.8 4.6"/>',
      database: '<ellipse cx="12" cy="6" rx="7.5" ry="3"/><path d="M4.5 6v6c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3V6"/><path d="M4.5 12v6c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3v-6"/>',
      upload: '<path d="M12 16V4"/><path d="M8 8l4-4 4 4"/><path d="M4 16v2.5A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5V16"/>',
      share: '<circle cx="17.5" cy="6" r="2.6"/><circle cx="6.5" cy="12" r="2.6"/><circle cx="17.5" cy="18" r="2.6"/><path d="M9 10.8l6-3.2M9 13.2l6 3.2"/>',
      package: '<path d="M12 3.5l7.5 4v9l-7.5 4-7.5-4v-9z"/><path d="M4.8 7.8L12 11.7l7.2-3.9M12 11.7V20"/>',
      star: '<path d="M12 3.6l2.6 5.3 5.9.9-4.3 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8-4.3-4.1 5.9-.9z"/>',
      info: '<circle cx="12" cy="12" r="8.4"/><path d="M12 11v5.2M12 7.9h.01"/>',
      plus: '<path d="M12 5v14M5 12h14"/>',
      image: '<rect x="3" y="4.5" width="18" height="15" rx="3"/><circle cx="8.6" cy="9.6" r="1.4"/><path d="M4 17l4.6-4.3 3.6 3.2 3-2.6L20 18"/>',
      trash: '<path d="M4.5 7h15M9.5 7V4.8h5V7M6.5 7l.8 12.2a1.5 1.5 0 0 0 1.5 1.4h6.4a1.5 1.5 0 0 0 1.5-1.4L17.5 7"/>'
    };
    return '<button class="menu-row' + (danger ? " menu-row--danger" : "") + '" data-act="' + act + '">' +
      '<span class="menu-row__icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' +
      (icons[icon] || "") + "</svg></span>" +
      '<span class="menu-row__main"><b>' + esc(title) + "</b><small>" + esc(sub) + "</small></span>" +
      '<span class="menu-row__tail">' + esc(tail || "") +
      '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M9 6l6 6-6 6"/></svg>' +
      "</span></button>";
  }

  /* ================================ 交互 ================================ */
  document.addEventListener("click", async function (e) {
    var tab = e.target.closest(".tab");
    if (tab) {
      S.view = tab.dataset.view;
      renderChrome();
      renderView();
      return;
    }

    var el = e.target.closest("[data-act]");
    if (!el) return;
    var act = el.dataset.act;

    if (act === "scope") {
      S.scope = el.dataset.scope;
      S.profile.scope = S.scope;
      Store.saveProfile(S.profile);
      renderCountdown();
      return;
    }
    if (act === "toggle-past") {
      S.pastOpen = !S.pastOpen;
      renderCountdown();
      return;
    }
    if (act === "star") {
      e.stopPropagation();
      var c = el.dataset.course;
      if (S.starred.has(c)) S.starred.delete(c); else S.starred.add(c);
      S.profile.starred = Array.from(S.starred);
      await Store.saveProfile(S.profile);
      if (S.view === "exams") renderExams(); else renderCountdown();
      toast(S.starred.has(c) ? "已关注「" + c + "」" : "已取消关注");
      return;
    }
    if (act === "gdetail") {
      openGroupDetail(decodeURIComponent(el.dataset.g || ""));
      return;
    }
    if (act === "more") {
      S.listLimit = (S.listLimit || 120) + 150;
      renderExams();
      return;
    }
    if (act === "browse-all") {
      S.view = "exams";
      renderChrome();
      renderView();
      return;
    }
    if (act === "edit-profile") { openProfileForm(); return; }
    if (act === "add-entry") { openEntryForm(null); return; }
    if (act === "edit-entry") { openEntryForm(el.dataset.id); return; }
    if (act === "wallpaper") { openWallpaperSheet(); return; }
    if (act === "switch-profile") { openProfileList(); return; }
    if (act === "manage-dataset") { openDatasetSheet(); return; }
    if (act === "import") {
      // 从总表面板内触发时，先收起面板，避免两个面板互相覆盖
      if (sheetState) {
        closeSheet(null);
        setTimeout(function () { $("#fileInput").click(); }, 210);
      } else {
        $("#fileInput").click();
      }
      return;
    }
    if (act === "share") { shareApp(); return; }
    if (act === "export-bundle") { exportBundle(); return; }
    if (act === "manage-star") { openStarSheet(); return; }
    if (act === "about") { openAbout(); return; }
    if (act === "wipe") { wipeAll(); return; }
    if (act === "reset-filter") { resetFilters(); return; }
    if (act === "export") { openExportSheet(); return; }
    if (act === "reload") { location.reload(); return; }
  });

  $("#datasetBtn").addEventListener("click", openDatasetSheet);
  $("#profileBtn").addEventListener("click", function () {
    S.view = "me";
    renderChrome(); renderView();
  });

  $("#searchInput").addEventListener("input", function (e) {
    S.filters.q = e.target.value;
    $("#searchClear").hidden = !S.filters.q;
    $("#resetFilter").hidden = !(S.filters.grade || S.filters.major || S.filters.classKey || S.filters.q);
    clearTimeout(this._t);
    var self = this;
    this._t = setTimeout(function () {
      renderExams();
      persistFilters();
    }, 120);
  });

  $("#searchClear").addEventListener("click", function () {
    S.filters.q = "";
    $("#searchInput").value = "";
    $("#searchClear").hidden = true;
    renderExams();
    persistFilters();
  });

  $("#resetFilter").addEventListener("click", resetFilters);

  function resetFilters() {
    S.filters = { grade: "", major: "", classKey: "", q: "" };
    $("#searchInput").value = "";
    renderExams();
    persistFilters();
  }

  $$(".fchip[data-filter]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var key = btn.dataset.filter;
      if (key === "grade") openPicker("选择年级", gradeOptions(), S.filters.grade, function (v) {
        S.filters.grade = v;
        if (v && S.filters.major && !majorOptions(v).some(function (o) { return o.value === S.filters.major; })) S.filters.major = "";
        if (v && S.filters.classKey && S.filters.grade !== v) S.filters.classKey = "";
        refreshAfterFilter();
      });
      if (key === "major") openPicker("选择专业", majorOptions(S.filters.grade), S.filters.major, function (v) {
        S.filters.major = v;
        if (v && S.filters.classKey && S.filters.classKey.split("|")[1] !== v) S.filters.classKey = "";
        refreshAfterFilter();
      });
      if (key === "class") openPicker("选择班级", classOptions(S.filters.grade, S.filters.major), S.filters.classKey, function (v) {
        S.filters.classKey = v;
        if (v) {
          var p = v.split("|");
          S.filters.grade = p[0];
          S.filters.major = p[1];
        }
        refreshAfterFilter();
      }, true);
    });
  });

  function refreshAfterFilter() {
    renderExams();
    persistFilters();
  }

  function openPicker(title, options, activeValue, onPick, searchable) {
    var body = (searchable ? '<div class="search sheet-search"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.6-3.6"/></svg>' +
      '<input type="search" id="pickerSearch" placeholder="输入关键词筛选" autocomplete="off" /></div>' : "") +
      '<div class="opt-list" id="pickerList"></div>';

    openSheet({
      title: title,
      body: body,
      onMount: function (root, close) {
        var listEl = $("#pickerList", root);
        var searchEl = $("#pickerSearch", root);

        function draw(keyword) {
          var kw = (keyword || "").trim().toLowerCase();
          var html = '<button class="opt' + (!activeValue ? " is-active" : "") + '" data-value="">' +
            '<span class="opt__text"><b>全部</b></span>' +
            '<svg class="opt__check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M5 12.5l4.5 4.5L19 7"/></svg></button>';

          options
            .filter(function (o) { return !kw || o.label.toLowerCase().indexOf(kw) > -1 || (o.sub || "").toLowerCase().indexOf(kw) > -1; })
            .forEach(function (o) {
              html += '<button class="opt' + (o.value === activeValue ? " is-active" : "") + '" data-value="' + esc(o.value) + '">' +
                '<span class="opt__text"><b>' + esc(o.label) + "</b>" + (o.sub ? "<small>" + esc(o.sub) + "</small>" : "") + "</span>" +
                (o.count != null ? '<span class="opt__count">' + o.count + "</span>" : "") +
                '<svg class="opt__check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M5 12.5l4.5 4.5L19 7"/></svg></button>';
            });

          if (!html) html = '<p class="me-note">没有匹配项</p>';
          listEl.innerHTML = html;
        }
        draw("");

        listEl.addEventListener("click", function (e) {
          var o = e.target.closest(".opt");
          if (!o) return;
          close(o.dataset.value || "");
        });

        if (searchEl) {
          searchEl.addEventListener("input", function () { draw(this.value); });
          setTimeout(function () { searchEl.focus(); }, 240);
        }
      }
    }).then(function (v) {
      if (v === null || v === undefined) return;
      onPick(v);
    });
  }

  function gradeOptions() {
    var counts = {};
    S.records.forEach(function (r) { if (r.grade) counts[r.grade] = (counts[r.grade] || 0) + 1; });
    return Object.keys(counts).sort().map(function (g) {
      return { value: g, label: g, count: counts[g] };
    });
  }

  function majorOptions(grade) {
    var counts = {};
    S.records.forEach(function (r) {
      if (!r.major) return;
      if (grade && r.grade !== grade) return;
      counts[r.major] = (counts[r.major] || 0) + 1;
    });
    return Object.keys(counts).sort(function (a, b) { return a.localeCompare(b, "zh"); }).map(function (m) {
      return { value: m, label: m, count: counts[m] };
    });
  }

  function classOptions(grade, major) {
    var map = {};
    S.records.forEach(function (r) {
      if (!r.classKey) return;
      if (grade && r.grade !== grade) return;
      if (major && r.major !== major) return;
      if (!map[r.classKey]) map[r.classKey] = { label: r.cls, sub: r.grade + " · " + r.major, count: 0 };
      map[r.classKey].count++;
    });
    return Object.keys(map).sort().map(function (k) {
      return { value: k, label: map[k].label, sub: map[k].sub, count: map[k].count };
    });
  }

  /* ============================== 考试详情 ============================== */
  function kv(k, v) {
    return '<div class="kv"><span>' + esc(k) + "</span><b>" + esc(v) + "</b></div>";
  }

  /** 场次详情：同一课程同一时段可能覆盖多个班级与教室，这里合并展示 */
  function openGroupDetail(key) {
    var g = (S.groups || {})[key];
    if (!g) return;

    var rows = (S.scopeRecords || []).filter(function (r) {
      return r.course === g.course && r.date === g.date && (r.start || "") === (g.start || "");
    });
    var totalHead = rows.reduce(function (n, r) { return n + (r.headcount || 0); }, 0);
    var starred = S.starred.has(g.course);

    var body = '<div style="padding:2px 0 10px">' +
      '<h3 style="margin:0;font-size:18px;line-height:1.4;font-family:var(--serif);letter-spacing:.5px">' + esc(g.course) + "</h3>" +
      '<p style="margin:6px 0 0;font-size:10.5px;color:var(--faint)">' +
      esc([rows[0] && rows[0].code, rows[0] && rows[0].college].filter(Boolean).join(" · ") || "—") + "</p></div>" +
      '<div style="margin-top:4px">' +
      kv("考试日期", fullDateLabel(g.date) + " " + weekdayOf(g.date)) +
      kv("考试时间", timeRange(g)) +
      kv("考试地点", roomsText(g)) +
      kv("涉及班级", rows.length + " 个" + (totalHead ? " · " + totalHead + " 人次" : "")) +
      "</div>";

    if (rows.length > 1) {
      body += '<div class="section-head" style="padding:0;margin:16px 2px 6px"><h3>班级与考场</h3></div>' +
        '<div>' + rows.map(function (r) {
          return '<div class="kv"><span style="min-width:0;flex:1">' + esc(r.cls || "多班合考") + "</span>" +
            "<b>" + esc((r.rooms || []).join("、") || "—") + "</b></div>";
        }).join("") + "</div>";
    }

    if (!rows.some(function (r) { return r.cls; })) {
      body += '<div class="notice" style="margin-top:12px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>' +
        "<span>该场次为多班合考，原表未标注具体班级，请结合教学班组成或学院通知确认。</span></div>";
    }

    openSheet({
      title: "考试详情",
      body: body,
      foot: '<button class="btn ' + (starred ? "btn--ghost" : "btn--primary") + ' btn--block" data-sheet-star>' +
        (starred ? "取消关注此课程" : "关注此课程，加进倒数日") + "</button>",
      onFoot: function (foot, close) {
        $("[data-sheet-star]", foot).onclick = async function () {
          if (S.starred.has(g.course)) S.starred.delete(g.course); else S.starred.add(g.course);
          S.profile.starred = Array.from(S.starred);
          await Store.saveProfile(S.profile);
          close(true);
          toast(S.starred.has(g.course) ? "已关注「" + g.course + "」" : "已取消关注");
          renderView();
        };
      }
    });
  }

  /* ============================== 总表管理 ============================== */
  async function openDatasetSheet() {
    if (!S.datasets.length) {
      var go = await confirmSheet({
        title: "导入考试总表",
        message: "当前还没有任何总表。导入教务下发的 Excel 安排表开始使用。",
        okText: "选择文件"
      });
      if (go) $("#fileInput").click();
      return;
    }

    var body = S.datasets.map(function (d) {
      var s = d.summary || {};
      var isActive = d.id === S.activeId;
      return '<div class="ds-row' + (isActive ? " is-active" : "") + '" data-id="' + esc(d.id) + '">' +
        '<button class="ds-row__radio" data-act="pick-ds" data-id="' + esc(d.id) + '" aria-label="选择"></button>' +
        '<button class="ds-row__main" data-act="pick-ds" data-id="' + esc(d.id) + '" style="text-align:left">' +
        "<b>" + esc(d.title) + (d.builtin ? '<span class="ds-row__tag">内置</span>' : "") +
        (d.ownerId !== "shared" ? '<span class="ds-row__tag">私有</span>' : "") + "</b>" +
        "<small>" + (s.recordCount || 0) + " 条 · " + (s.courseCount || 0) + " 门课 · " +
        (s.classCount || 0) + " 个班<br/>" + esc(s.dateFrom || "") + " ~ " + esc(s.dateTo || "") +
        (d.builtin ? "<br/>内置示例数据" : (d.sourceName ? "<br/>" + esc(d.sourceName) : "")) + "</small></button>" +
        '<button class="ds-row__del" data-act="del-ds" data-id="' + esc(d.id) + '" aria-label="删除总表">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"><path d="M4.5 7h15M9.5 7V4.8h5V7M6.5 7l.8 12.2a1.5 1.5 0 0 0 1.5 1.4h6.4a1.5 1.5 0 0 0 1.5-1.4L17.5 7"/></svg>' +
        "</button></div>";
    }).join("");

    body += '<div class="notice" style="margin-top:6px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>' +
      "<span>各总表数据相互独立：切换只改变查询范围，删除只清除该表自身记录，不影响其他总表。</span></div>";

    openSheet({
      title: "考试总表",
      body: body,
      foot: '<button class="btn btn--primary btn--block" data-act="import">导入新的总表</button>',
      onMount: function (root, close) {
        root.addEventListener("click", async function (e) {
          var pick = e.target.closest('[data-act="pick-ds"]');
          var del = e.target.closest('[data-act="del-ds"]');
          if (pick) {
            var id = pick.dataset.id;
            if (id === S.activeId) { close(null); return; }
            await switchDataset(id);
            close(null);
          }
          if (del) {
            e.stopPropagation();
            var did = del.dataset.id;
            var d = S.datasets.find(function (x) { return x.id === did; });
            close(null);
            var ok = await confirmSheet({
              title: "删除总表",
              message: "将删除「" + esc(d.title) + "」及其 " + (d.summary.recordCount || 0) +
                " 条考试记录。其他总表不受影响，此操作不可撤销。",
              okText: "删除",
              danger: true
            });
            if (ok) await removeDataset(did);
          }
        });
      }
    });
  }

  async function switchDataset(id) {
    if (!findVisible(id)) return;
    S.activeId = id;
    await Store.setSetting("activeDataset", id);
    await loadActiveRecords();
    await restoreFilters();
    if (S.filters.classKey && !S.records.some(function (r) { return r.classKey === S.filters.classKey; })) {
      S.filters = { grade: "", major: "", classKey: "", q: "" };
      $("#searchInput").value = "";
    }
    renderChrome();
    renderView();
    toast("已翻到「" + (activeDataset() ? activeDataset().title : "") + "」");
  }

  async function removeDataset(id) {
    var wasActive = id === S.activeId;
    await Store.deleteDataset(id);
    S.cache.delete(id);
    await refreshDatasets();
    if (wasActive) {
      S.activeId = S.datasets[0] ? S.datasets[0].id : null;
      await Store.setSetting("activeDataset", S.activeId || "");
      await loadActiveRecords();
    }
    renderChrome();
    renderView();
    toast("已删除该总表，其他总表未受影响");
  }

  /* ============================== 导入流程 ============================== */
  $("#fileInput").addEventListener("change", async function (e) {
    var files = Array.prototype.slice.call(e.target.files || []);
    e.target.value = "";
    if (!files.length) return;

    if (/\.json$/i.test(files[0].name)) {
      await importBundle(files[0]);
      return;
    }

    var body = '<p style="font-size:12px;color:var(--muted);margin:2px 0 10px">将导入 ' + files.length + " 个文件：</p>" +
      '<div style="margin-bottom:14px">' + files.map(function (f) {
        return '<div class="kv"><b style="text-align:left;font-weight:500;word-break:break-all">' + esc(f.name) + "</b><span>" +
          (f.size / 1024).toFixed(0) + " KB</span></div>";
      }).join("") + "</div>" +
      '<div class="field"><label>数据归属</label>' +
      '<div class="choice-row" id="ownerChoice">' +
      '<button class="choice is-active" data-owner="shared">所有人共用</button>' +
      '<button class="choice" data-owner="private">仅当前档案可见</button>' +
      "</div>" +
      '<p class="hint">共用：本机所有使用者都能查询该表；私有：仅「' + esc(S.profile.name) + "」可见，实现多人配置隔离。</p></div>";

    var owner = "shared";
    openSheet({
      title: "导入考试总表",
      body: body,
      foot: '<button class="btn btn--primary btn--block" data-do-import>开始导入</button>',
      onMount: function (root) {
        root.addEventListener("click", function (ev) {
          var c = ev.target.closest("[data-owner]");
          if (!c) return;
          $$(".choice", root).forEach(function (x) { x.classList.remove("is-active"); });
          c.classList.add("is-active");
          owner = c.dataset.owner;
        });
      },
      onFoot: function (foot, close) {
        $("[data-do-import]", foot).onclick = function () { close({ owner: owner }); };
      }
    }).then(async function (res) {
      if (!res) return;
      await runImport(files, res.owner);
    });
  });

  async function runImport(files, owner) {
    loading(true, "正在展卷解析…", "首次使用需加载解析组件");
    try {
      var XLSX = await ensureXLSX();
      var okCount = 0, lastId = null, allWarnings = [];

      for (var i = 0; i < files.length; i++) {
        var file = files[i];
        loading(true, "正在解析 " + (i + 1) + " / " + files.length, file.name);
        var buf = await file.arrayBuffer();
        var wb;
        try {
          wb = XLSX.read(new Uint8Array(buf), { type: "array" });
        } catch (err) {
          allWarnings.push(file.name + "：文件无法读取（" + err.message + "）");
          continue;
        }

        var parsed = window.QingJianParser.parseWorkbook(wb, XLSX, { sourceName: file.name, title: "" });
        if (!parsed.ok) {
          allWarnings.push(file.name + "：" + parsed.error);
          continue;
        }

        var meta = Object.assign({}, parsed.dataset, {
          ownerId: owner === "private" ? S.profile.id : "shared",
          builtin: false,
          importedAt: new Date().toISOString()
        });
        meta.campus = meta.campus || "";
        meta.term = meta.term || "";

        await Store.saveDataset(meta, parsed.records);
        S.cache.set(meta.id, parsed.records);
        lastId = meta.id;
        okCount++;
        allWarnings = allWarnings.concat((parsed.warnings || []).map(function (w) { return file.name + "：" + w; }));
      }

      await refreshDatasets();
      if (lastId) {
        S.activeId = lastId;
        await Store.setSetting("activeDataset", lastId);
        await loadActiveRecords();
        S.filters = { grade: "", major: "", classKey: "", q: "" };
        $("#searchInput").value = "";
      }
      renderChrome();
      renderView();

      loading(false);
      if (!okCount) {
        toast("导入失败，请检查表格格式", "error");
        if (allWarnings.length) openWarnings("导入失败", allWarnings);
        return;
      }

      toast("已导入 " + okCount + " 份总表 · " + S.records.length + " 条记录");
      if (allWarnings.length) setTimeout(function () { openWarnings("导入说明", allWarnings); }, 400);
    } catch (err) {
      console.error(err);
      loading(false);
      toast(err.message || "导入失败", "error");
    }
  }

  function openWarnings(title, list) {
    openSheet({
      title: title,
      body: '<div class="notice" style="margin-bottom:12px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>' +
        "<span>以下情况已自动处理或需要留意：</span></div>" +
        list.map(function (w) {
          return '<div class="kv"><b style="text-align:left;font-weight:400;font-size:11px;line-height:1.65">' + esc(w) + "</b></div>";
        }).join("")
    });
  }

  /* ============================== 使用者档案 ============================== */
  function openProfileForm() {
    var p = S.profile;
    var grades = gradeOptions();
    if (!grades.length) {
      grades = ["24级", "25级", "26级"].map(function (g) { return { value: g, label: g }; });
    }
    var body =
      '<div class="field"><label for="pfName">称呼</label>' +
      '<input id="pfName" maxlength="12" placeholder="如：小林" value="' + esc(p.name || "") + '" /></div>' +
      '<div class="field"><label for="pfGrade">年级</label>' +
      '<select id="pfGrade"><option value="">未设置</option>' +
      grades.map(function (g) { return '<option value="' + esc(g.value) + '"' + (p.grade === g.value ? " selected" : "") + ">" + esc(g.label) + "</option>"; }).join("") +
      "</select></div>" +
      '<div class="field"><label for="pfMajor">专业</label>' +
      '<select id="pfMajor"><option value="">未设置</option>' +
      majorOptions(p.grade).map(function (m) { return '<option value="' + esc(m.value) + '"' + (p.major === m.value ? " selected" : "") + ">" + esc(m.label) + "</option>"; }).join("") +
      "</select></div>" +
      '<div class="field"><label for="pfClass">班级</label>' +
      '<select id="pfClass"><option value="">未设置</option>' +
      classOptions(p.grade, p.major).map(function (c) {
        return '<option value="' + esc(c.value) + '"' + (p.classKey === c.value ? " selected" : "") + ">" + esc(c.label) + "</option>";
      }).join("") +
      "</select><p class=\"hint\">切换年级会同步更新可选专业与班级。</p></div>";

    openSheet({
      title: "我的档案",
      body: body,
      foot: '<button class="btn btn--primary btn--block" data-save>保存</button>',
      onMount: function (root) {
        var gSel = $("#pfGrade", root), mSel = $("#pfMajor", root), cSel = $("#pfClass", root);

        gSel.addEventListener("change", function () {
          var majors = majorOptions(gSel.value);
          mSel.innerHTML = '<option value="">未设置</option>' + majors.map(function (m) {
            return '<option value="' + esc(m.value) + '">' + esc(m.label) + "</option>";
          }).join("");
          var classes = classOptions(gSel.value, "");
          cSel.innerHTML = '<option value="">未设置</option>' + classes.map(function (c) {
            return '<option value="' + esc(c.value) + '">' + esc(c.label) + "</option>";
          }).join("");
        });
        mSel.addEventListener("change", function () {
          var classes = classOptions(gSel.value, mSel.value);
          cSel.innerHTML = '<option value="">未设置</option>' + classes.map(function (c) {
            return '<option value="' + esc(c.value) + '">' + esc(c.label) + "</option>";
          }).join("");
        });
      },
      onFoot: function (foot, close) {
        $("[data-save]", foot).onclick = async function () {
          var name = $("#pfName").value.trim() || "我";
          var grade = $("#pfGrade").value;
          var major = $("#pfMajor").value;
          var classKey = $("#pfClass").value;
          var parts = classKey ? classKey.split("|") : [];
          S.profile.name = name;
          S.profile.grade = grade;
          S.profile.major = major;
          S.profile.classKey = classKey;
          S.profile.classNo = parts[2] || "";
          S.profile.cls = classKey ? parts[0] + parts[1] + parts[2] + "班" : "";
          await Store.saveProfile(S.profile);
          if (classKey && !S.filters.grade && !S.filters.major && !S.filters.classKey) {
            S.filters.grade = parts[0];
            S.filters.major = parts[1];
            S.filters.classKey = classKey;
            persistFilters();
          }
          close(true);
          renderChrome();
          renderView();
          toast("已保存档案");
        };
      }
    });
  }

  function openProfileList() {
    var body = S.profiles.map(function (p) {
      var isActive = p.id === S.profile.id;
      return '<div class="ds-row' + (isActive ? " is-active" : "") + '" data-pid="' + esc(p.id) + '">' +
        '<span class="ds-row__radio"></span>' +
        '<button class="ds-row__main" data-pick="' + esc(p.id) + '" style="text-align:left">' +
        "<b>" + esc(p.name) + (isActive ? '<span class="ds-row__tag">当前</span>' : "") + "</b>" +
        "<small>" + esc(p.classKey ? [p.grade, p.major, p.classNo + "班"].filter(Boolean).join(" · ") : "未设置班级") +
        (p.starred && p.starred.length ? " · 关注 " + p.starred.length + " 门" : "") + "</small></button>" +
        (isActive ? "" : '<button class="ds-row__del" data-del="' + esc(p.id) + '" aria-label="删除档案"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"><path d="M4.5 7h15M9.5 7V4.8h5V7M6.5 7l.8 12.2a1.5 1.5 0 0 0 1.5 1.4h6.4a1.5 1.5 0 0 0 1.5-1.4L17.5 7"/></svg></button>') +
        "</div>";
    }).join("");

    body += '<p class="me-note" style="text-align:left;margin:10px 2px 0">同一台设备可建立多个档案，各自保存年级、班级、壁纸与关注科目；标记为「仅本人可见」的总表只对导入者显示。</p>';

    openSheet({
      title: "切换使用者",
      body: body,
      foot: '<button class="btn btn--primary btn--block" data-new-profile>新建档案</button>',
      onMount: function (root, close) {
        root.addEventListener("click", async function (e) {
          var pick = e.target.closest("[data-pick]");
          var del = e.target.closest("[data-del]");
          if (pick) {
            await switchProfile(pick.dataset.pick);
            close(null);
          }
          if (del) {
            e.stopPropagation();
            var pid = del.dataset.del;
            if (S.profiles.length <= 1) { toast("至少保留一个档案"); return; }
            close(null);
            var ok = await confirmSheet({
              title: "删除档案",
              message: "该档案的年级、班级、壁纸与关注配置将被清除，私有总表也会一并移除。",
              okText: "删除", danger: true
            });
            if (ok) {
              await Store.deleteProfile(pid);
              var priv = S.datasets.filter(function (d) { return d.ownerId === pid; });
              for (var i = 0; i < priv.length; i++) await Store.deleteDataset(priv[i].id);
              S.profiles = S.profiles.filter(function (p) { return p.id !== pid; });
              if (S.profile.id === pid) await switchProfile(S.profiles[0].id);
              else { await refreshDatasets(); renderChrome(); renderView(); }
              toast("档案已删除");
            }
          }
        });
      },
      onFoot: function (foot, close) {
        $("[data-new-profile]", foot).onclick = async function () {
          var np = {
            id: "p-" + Date.now().toString(36),
            name: "同学 " + (S.profiles.length + 1),
            grade: "", major: "", classKey: "", cls: "",
            starred: [], scope: "class",
            createdAt: new Date().toISOString()
          };
          await Store.saveProfile(np);
          S.profiles.push(np);
          close(null);
          await switchProfile(np.id);
          openProfileForm();
        };
      }
    });
  }

  async function switchProfile(id) {
    var p = S.profiles.find(function (x) { return x.id === id; });
    if (!p) return;
    S.profile = p;
    S.starred = new Set(p.starred || []);
    S.scope = p.scope || "class";
    S.entries = await Store.listEntries(id);
    await loadWallpaper();
    await Store.setSetting("activeProfile", id);
    await refreshDatasets();
    if (!findVisible(S.activeId)) {
      S.activeId = S.datasets[0] ? S.datasets[0].id : null;
      await Store.setSetting("activeDataset", S.activeId || "");
    }
    await loadActiveRecords();
    await restoreFilters();
    renderChrome();
    renderView();
    toast("已切换到「" + p.name + "」");
  }

  /* ============================== 导出与分享 ============================== */
  function exportCtx() {
    var list = filteredRecords();
    var custom = customEvents(S.filters.q).map(function (e) {
      return {
        custom: true,
        course: e.course,
        code: "",
        date: e.date,
        start: e.start,
        end: e.end,
        rooms: e.rooms,
        grade: "",
        major: "",
        cls: e.entryType === "lab" ? "自建实验" : "自建考试",
        headcount: null,
        college: "",
        note: e.note || ""
      };
    });
    var merged = list.concat(custom).sort(function (a, b) {
      return (a.date + (a.start || "99:99")).localeCompare(b.date + (b.start || "99:99"));
    });

    var ds = activeDataset();
    var f = S.filters;
    var scopeKey = f.classKey
      ? f.classKey.split("|")[0] + f.classKey.split("|")[1] + f.classKey.split("|")[2] + "班"
      : ([f.grade, f.major].filter(Boolean).join("") || "全部");
    return {
      records: merged,
      datasetTitle: ds ? ds.title : "",
      docTitle: "考试纪事 · 考程表",
      scopeLabel: filterScopeLabel(),
      filenameBase: ("青简叙_考程_" + scopeKey).replace(/[\\/:*?"<>|\s·]/g, ""),
      count: merged.length
    };
  }

  async function openExportSheet() {
    var ctx = exportCtx();
    var list = ctx.records;
    if (!list.length) { toast("当前筛选结果为空"); return; }
    var courses = uniqueBy(list, function (r) { return r.course; }).length;

    var body = '<div class="notice" style="margin-bottom:12px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
      '<circle cx="12" cy="12" r="9"/><path d="M12 8.5v5M12 16.5h.01"/></svg>' +
      "<span>将导出当前筛选结果：<b>" + list.length + "</b> 场 · <b>" + courses + "</b> 个科目（" +
      esc(filterScopeLabel()) + "）</span></div>" +
      '<div class="export-grid">' +
      tile("xls", "Excel", ".xlsx · 可继续编辑", "XLS") +
      tile("pdf", "PDF", ".pdf · 适合打印分享", "PDF") +
      tile("csv", "CSV", ".csv · 通用导入", "CSV") +
      tile("doc", "Word", ".doc · 可改写排版", "DOC") +
      "</div>" +
      '<p class="export-note">PDF 按 A4 逐页渲染，中文无需额外字体；文件数量较多时请保持页面在前台。</p>';

    openSheet({ title: "导出考程", body: body });
  }

  function tile(kind, name, desc, tag) {
    return '<button class="export-tile export-tile--' + kind + '" data-export="' + kind + '">' +
      '<span class="export-tile__icon">' + tag + "</span>" +
      "<b>" + name + "</b><small>" + desc + "</small></button>";
  }

  document.addEventListener("click", async function (e) {
    var t = e.target.closest("[data-export]");
    if (!t) return;
    var kind = t.dataset.export;
    var ctx = exportCtx();
    if (!ctx.records.length) { toast("没有可导出的记录"); return; }

    try {
      if (kind === "csv") {
        Exporter.toCSV(ctx.records, ctx);
        toast("CSV 已开始下载");
        closeSheet(null);
      } else if (kind === "xls") {
        loading(true, "正在生成 Excel…");
        await ensureXLSX();
        Exporter.toExcel(ctx.records, ctx);
        loading(false);
        closeSheet(null);
        toast("Excel 已开始下载");
      } else if (kind === "doc") {
        Exporter.toWord(ctx.records, ctx);
        toast("Word 已开始下载");
        closeSheet(null);
      } else if (kind === "pdf") {
        var pages = Math.ceil(ctx.records.length / Exporter.ROWS_PER_PAGE);
        loading(true, "正在生成 PDF…", "共 " + ctx.records.length + " 场 · " + pages + " 页");
        await ensurePDF();
        await Exporter.toPDF(ctx.records, ctx, function (cur, total) {
          $("#loadingSub").textContent = "第 " + cur + " / " + total + " 页";
        });
        loading(false);
        closeSheet(null);
        toast("PDF 已开始下载");
      }
    } catch (err) {
      console.error(err);
      loading(false);
      toast(err.message || "导出失败", "error");
    }
  });

  async function shareApp() {
    // 保留 location.search：托管平台的访问凭证可能放在查询串里
    var base = location.origin + location.pathname + location.search;
    var url = base + (S.profile.classKey ? "#c=" + encodeURIComponent(S.profile.classKey) : "");
    var text = S.profile.classKey
      ? "青简叙｜已按「" + S.profile.cls + "」筛选，打开就能看考试倒数日与考场"
      : "青简叙｜考试纪事与倒数日，选好班级即可使用";

    if (navigator.share) {
      try {
        await navigator.share({ title: "青简叙 · 考试纪事与倒数日", text: text, url: url });
        return;
      } catch (err) {
        if (err && err.name === "AbortError") return;
      }
    }
    try {
      await navigator.clipboard.writeText(text + "\n" + url);
      toast("链接已复制，去粘贴给同学吧");
    } catch (err) {
      openSheet({
        title: "分享链接",
        body: '<p style="font-size:12px;color:var(--muted);margin:0 0 10px">长按或复制下面的链接：</p>' +
          '<div style="padding:12px;border-radius:10px;background:var(--paper-2);font-size:11px;word-break:break-all;line-height:1.7">' + esc(url) + "</div>"
      });
    }
  }

  async function exportBundle() {
    var ds = activeDataset();
    if (!ds) { toast("没有可导出的总表"); return; }
    loading(true, "正在打包…");
    try {
      var records = await Store.getRecords(ds.id);
      var bundle = { type: "qingjianxu-bundle", version: 1, exportedAt: new Date().toISOString(), dataset: ds, records: records };
      var blob = new Blob([JSON.stringify(bundle)], { type: "application/json" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "青简叙数据包_" + ds.title.replace(/[\\/:*?"<>|]/g, "").slice(0, 18) + ".json";
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1500);
      loading(false);
      toast("数据包已生成，可发给同学离线导入");
    } catch (err) {
      loading(false);
      toast("打包失败", "error");
    }
  }

  async function importBundle(file) {
    loading(true, "正在导入数据包…");
    try {
      var text = await file.text();
      var bundle = JSON.parse(text);
      if (bundle.type !== "qingjianxu-bundle" || !bundle.dataset || !Array.isArray(bundle.records)) {
        throw new Error("不是有效的数据包文件");
      }
      var meta = Object.assign({}, bundle.dataset, {
        id: "ds-" + Date.now().toString(36),
        builtin: false,
        ownerId: "shared",
        importedAt: new Date().toISOString()
      });
      var records = bundle.records.map(function (r, i) {
        r.id = meta.id + "#" + (i + 1);
        r.datasetId = meta.id;
        return r;
      });
      await Store.saveDataset(meta, records);
      S.cache.set(meta.id, records);
      await refreshDatasets();
      S.activeId = meta.id;
      await Store.setSetting("activeDataset", meta.id);
      await loadActiveRecords();
      renderChrome(); renderView();
      loading(false);
      toast("已导入数据包 · " + records.length + " 条记录");
    } catch (err) {
      loading(false);
      toast(err.message || "导入失败", "error");
    }
  }

  function openStarSheet() {
    if (!S.starred.size) {
      openSheet({
        title: "已关注课程",
        body: emptyState("还没有关注任何课程", "在考程表里点击课程右侧的星标，即可把它加入倒数日的「关注」范围。")
      });
      return;
    }
    var list = Array.from(S.starred);
    openSheet({
      title: "已关注课程",
      body: '<div class="opt-list">' + list.map(function (c) {
        return '<div class="opt"><span class="opt__text"><b>' + esc(c) + "</b></span>" +
          '<button class="btn btn--ghost btn--sm" data-unstar="' + esc(c) + '">移除</button></div>';
      }).join("") + "</div>",
      onMount: function (root, close) {
        root.addEventListener("click", async function (e) {
          var b = e.target.closest("[data-unstar]");
          if (!b) return;
          S.starred.delete(b.dataset.unstar);
          S.profile.starred = Array.from(S.starred);
          await Store.saveProfile(S.profile);
          close(null);
          renderView();
          toast("已移除关注");
        });
      }
    });
  }

  function openAbout() {
    var ds = activeDataset();
    var s = ds ? ds.summary : {};
    var body =
      '<div style="padding:2px 0 8px"><h3 style="margin:0 0 6px;font-size:16px;font-family:var(--serif);letter-spacing:1px">青简叙</h3>' +
      '<p style="margin:0;font-size:11.5px;line-height:1.85;color:var(--muted)">以青简记考程，叙一季光阴。把教务 Excel 总表变成可筛选、可倒数、可补录、可导出的手机页面。开源、纯前端、无需登录。</p></div>' +
      '<div style="margin-top:10px">' +
      kv("当前总表", ds ? ds.title : "—") +
      kv("数据来源", ds ? (ds.builtin ? "内置示例数据（虚构，不含真实课表）" : (ds.sourceName || "本机导入")) : "内置示例数据") +
      kv("考试场次", (s.recordCount || 0) + " 条") +
      kv("覆盖范围", (s.dateFrom || "") + " ~ " + (s.dateTo || "")) +
      kv("班级数量", (s.classCount || 0) + " 个") +
      "</div>" +
      '<div class="notice" style="margin-top:14px"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>' +
      "<span>本应用为查询辅助工具，考试时间与地点请以教务系统及学院通知为准。若发现数据差异，请以官方安排为准。</span></div>" +
      '<div style="margin-top:14px">' +
      kv("版本", "3.0（青简叙 · 自建条目 · 自定义壁纸）") +
      kv("制作者", "叶盗") +
      kv("隐私", "所有数据存于本机 IndexedDB，不上传服务器") +
      kv("与 2.0 关系", "独立目录、独立数据库、独立入口，互不影响") +
      kv("开源协议", "MIT") +
      kv("技术栈", "原生 JS · SheetJS · jsPDF · html2canvas") +
      "</div>";
    openSheet({ title: "关于与数据说明", body: body });
  }

  async function wipeAll() {
    var ok = await confirmSheet({
      title: "清空本机数据",
      message: "将删除本机保存的所有总表、使用者档案、自建条目与壁纸配置，并恢复内置示例数据。此操作不可撤销，建议先导出数据包备份。",
      okText: "全部清空",
      danger: true
    });
    if (!ok) return;
    await Store.clearAll();
    S.cache.clear();
    location.reload();
  }

  /* ============================== 启动 ============================== */
  function registerSW() {
    if (!("serviceWorker" in navigator)) return;
    if (location.protocol !== "http:" && location.protocol !== "https:") return;
    navigator.serviceWorker.register("sw.js").catch(function () { /* 离线为增强项，失败不影响使用 */ });
  }

  document.addEventListener("DOMContentLoaded", function () {
    boot().catch(function (err) {
      console.error(err);
      $("#countdownBody").innerHTML = emptyState(
        "初始化失败",
        esc(err.message || "浏览器存储不可用") + "。请尝试刷新页面，或更换浏览器打开。",
        "重新加载",
        "reload"
      );
    });
  });
})();
