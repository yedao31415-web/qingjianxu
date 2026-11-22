/*!
 * 青简叙 · QingJianXu Store — 本机数据层（IndexedDB）
 *
 * 与 2.0「考试雷达」完全隔离：
 *   - 数据库名独立（qingjianxu），不打开、不升级、不写入 2.0 的数据库；
 *   - 缓存前缀独立（qingjianxu-*），同一域名下共存互不覆盖；
 *   - 所有配置只存在本机 IndexedDB，不上传服务器。
 *
 * 结构：
 *   datasets  一份考试总表 = 一条记录（ownerId: shared / 某档案 id）
 *   records   主键 datasetId#序号，按 datasetId 建索引，删除只清本表主键区间
 *   profiles  使用者档案（年级 / 专业 / 班级 / 关注）
 *   entries   手动录入的实验与考试（按档案隔离）
 *   settings  键值配置（当前总表、壁纸、筛选状态等）
 */
(function (root) {
  "use strict";

  var DB_NAME = "qingjianxu";
  var DB_VERSION = 1;
  var dbPromise = null;
  var memory = null; // IndexedDB 不可用时的内存降级容器

  function open() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      if (!root.indexedDB) {
        reject(new Error("当前环境不支持 IndexedDB"));
        return;
      }
      var req = root.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function (e) {
        var db = e.target.result;
        if (!db.objectStoreNames.contains("datasets")) {
          var ds = db.createObjectStore("datasets", { keyPath: "id" });
          ds.createIndex("ownerId", "ownerId", { unique: false });
        }
        if (!db.objectStoreNames.contains("records")) {
          var rs = db.createObjectStore("records", { keyPath: "id" });
          rs.createIndex("datasetId", "datasetId", { unique: false });
          rs.createIndex("classKey", "classKey", { unique: false });
        }
        if (!db.objectStoreNames.contains("profiles")) {
          db.createObjectStore("profiles", { keyPath: "id" });
        }
        if (!db.objectStoreNames.contains("entries")) {
          var es = db.createObjectStore("entries", { keyPath: "id" });
          es.createIndex("ownerId", "ownerId", { unique: false });
        }
        if (!db.objectStoreNames.contains("settings")) {
          db.createObjectStore("settings", { keyPath: "key" });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error || new Error("数据库打开失败")); };
      req.onblocked = function () { reject(new Error("数据库被其他标签页占用，请关闭后重试")); };
    }).catch(function (err) {
      console.warn("[青简叙] IndexedDB 不可用，降级为内存存储：", err && err.message);
      memory = {
        datasets: new Map(), records: new Map(), profiles: new Map(),
        settings: new Map(), entries: new Map()
      };
      return null;
    });
    return dbPromise;
  }

  function tx(db, stores, mode) {
    var t = db.transaction(stores, mode || "readonly");
    return {
      t: t,
      store: function (name) { return t.objectStore(name); },
      done: new Promise(function (resolve, reject) {
        t.oncomplete = function () { resolve(); };
        t.onerror = function () { reject(t.error); };
        t.onabort = function () { reject(t.error || new Error("事务已中止")); };
      })
    };
  }

  function req2p(req) {
    return new Promise(function (resolve, reject) {
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  /* ------------------------------- datasets ------------------------------- */
  async function listDatasets() {
    var db = await open();
    if (!db) return Array.from(memory.datasets.values());
    var t = tx(db, ["datasets"]);
    return req2p(t.store("datasets").getAll());
  }

  async function getDataset(id) {
    var db = await open();
    if (!db) return memory.datasets.get(id) || null;
    var t = tx(db, ["datasets"]);
    return req2p(t.store("datasets").get(id));
  }

  async function saveDataset(meta, records) {
    var db = await open();
    var stamped = records.map(function (r) {
      r.datasetId = meta.id;
      return r;
    });

    if (!db) {
      memory.datasets.set(meta.id, meta);
      stamped.forEach(function (r) { memory.records.set(r.id, r); });
      return meta;
    }

    var t = tx(db, ["datasets", "records"], "readwrite");
    t.store("datasets").put(meta);
    var rs = t.store("records");
    stamped.forEach(function (r) { rs.put(r); });
    await t.done;
    return meta;
  }

  async function getRecords(datasetId) {
    var db = await open();
    if (!db) {
      return Array.from(memory.records.values()).filter(function (r) { return r.datasetId === datasetId; });
    }
    var t = tx(db, ["records"]);
    var idx = t.store("records").index("datasetId");
    return req2p(idx.getAll(IDBKeyRange.only(datasetId)));
  }

  async function deleteDataset(id) {
    var db = await open();
    if (!db) {
      memory.datasets.delete(id);
      Array.from(memory.records.keys()).forEach(function (k) {
        if (memory.records.get(k).datasetId === id) memory.records.delete(k);
      });
      return;
    }
    var t = tx(db, ["datasets", "records"], "readwrite");
    t.store("datasets").delete(id);
    // 只删本表主键区间，不触及其他总表
    t.store("records").delete(IDBKeyRange.bound(id + "#", id + "#\uffff"));
    await t.done;
  }

  async function updateDataset(id, patch) {
    var cur = await getDataset(id);
    if (!cur) return null;
    var next = Object.assign({}, cur, patch);
    var db = await open();
    if (!db) { memory.datasets.set(id, next); return next; }
    var t = tx(db, ["datasets"], "readwrite");
    t.store("datasets").put(next);
    await t.done;
    return next;
  }

  /* ------------------------------- profiles ------------------------------- */
  async function listProfiles() {
    var db = await open();
    if (!db) return Array.from(memory.profiles.values());
    var t = tx(db, ["profiles"]);
    return req2p(t.store("profiles").getAll());
  }

  async function saveProfile(profile) {
    var db = await open();
    if (!db) { memory.profiles.set(profile.id, profile); return profile; }
    var t = tx(db, ["profiles"], "readwrite");
    t.store("profiles").put(profile);
    await t.done;
    return profile;
  }

  async function deleteProfile(id) {
    var db = await open();
    if (!db) { memory.profiles.delete(id); return; }
    var t = tx(db, ["profiles"], "readwrite");
    t.store("profiles").delete(id);
    await t.done;
  }

  /* -------------------------- entries（手动录入） -------------------------- */
  async function listEntries(ownerId) {
    var db = await open();
    if (!db) {
      return Array.from(memory.entries.values()).filter(function (e) { return e.ownerId === ownerId; });
    }
    var t = tx(db, ["entries"]);
    var all = await req2p(t.store("entries").getAll());
    return all.filter(function (e) { return e.ownerId === ownerId; });
  }

  async function saveEntry(entry) {
    var db = await open();
    if (!db) { memory.entries.set(entry.id, entry); return entry; }
    var t = tx(db, ["entries"], "readwrite");
    t.store("entries").put(entry);
    await t.done;
    return entry;
  }

  async function deleteEntry(id) {
    var db = await open();
    if (!db) { memory.entries.delete(id); return; }
    var t = tx(db, ["entries"], "readwrite");
    t.store("entries").delete(id);
    await t.done;
  }

  /* ------------------------------- settings ------------------------------- */
  async function getSetting(key, fallback) {
    var db = await open();
    if (!db) {
      return memory.settings.has(key) ? memory.settings.get(key) : fallback;
    }
    var t = tx(db, ["settings"]);
    var row = await req2p(t.store("settings").get(key));
    return row ? row.value : fallback;
  }

  async function setSetting(key, value) {
    var db = await open();
    if (!db) { memory.settings.set(key, value); return value; }
    var t = tx(db, ["settings"], "readwrite");
    t.store("settings").put({ key: key, value: value });
    await t.done;
    return value;
  }

  async function clearAll() {
    var db = await open();
    if (!db) {
      memory.datasets.clear(); memory.records.clear(); memory.profiles.clear();
      memory.settings.clear(); memory.entries.clear();
      return;
    }
    var t = tx(db, ["datasets", "records", "profiles", "settings", "entries"], "readwrite");
    ["datasets", "records", "profiles", "settings", "entries"].forEach(function (s) { t.store(s).clear(); });
    await t.done;
  }

  async function usage() {
    if (!root.navigator || !navigator.storage || !navigator.storage.estimate) return null;
    try {
      var est = await navigator.storage.estimate();
      return { usage: est.usage || 0, quota: est.quota || 0 };
    } catch (e) {
      return null;
    }
  }

  root.QingJianStore = {
    open: open,
    listDatasets: listDatasets,
    getDataset: getDataset,
    saveDataset: saveDataset,
    getRecords: getRecords,
    deleteDataset: deleteDataset,
    updateDataset: updateDataset,
    listProfiles: listProfiles,
    saveProfile: saveProfile,
    deleteProfile: deleteProfile,
    listEntries: listEntries,
    saveEntry: saveEntry,
    deleteEntry: deleteEntry,
    getSetting: getSetting,
    setSetting: setSetting,
    clearAll: clearAll,
    usage: usage
  };
})(typeof window !== "undefined" ? window : globalThis);
