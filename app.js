(function(){
  'use strict';

  var DB_NAME = 'archiveBoxesDB';
  var DB_VERSION = 1;
  var STORE_NAME = 'store';
  var DATA_KEY = 'archiveData';

  var fileInput = document.getElementById('fileInput');
  var dataStatus = document.getElementById('dataStatus');
  var customerInput = document.getElementById('customerInput');
  var boxInput = document.getElementById('boxInput');
  var checkBtn = document.getElementById('checkBtn');
  var resultEl = document.getElementById('result');

  var HEADER_MAP = {
    customer: ["מס' לקוח", "מספר לקוח"],
    box: ["מס' תיבה לקוח", "מספר תיבה לקוח", "מס' תיבה", "מספר תיבה"],
    intakeDate: ["תאריך הכנסה"],
    shredDate: ["תאריך גריסה"],
    shredded: ["נגרס"]
  };

  // In-memory cache so repeated searches don't re-read IndexedDB every time.
  var cachedPayload = null; // { records, filename, updatedAt }
  var cachedIndex = null;   // Map "customer|box" -> record

  function norm(v){
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function normKey(v){
    // strip leading zeros / whitespace so "050007" matches "50007"
    var s = norm(v);
    if (/^\d+$/.test(s)) return String(parseInt(s, 10));
    return s;
  }

  function formatDate(v){
    if (!v) return '';
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return norm(v);
    return d.toLocaleDateString('he-IL');
  }

  function findKey(rowKeys, candidates){
    for (var i = 0; i < candidates.length; i++){
      for (var j = 0; j < rowKeys.length; j++){
        if (rowKeys[j].trim() === candidates[i]) return rowKeys[j];
      }
    }
    return null;
  }

  function parseWorkbook(workbook){
    var sheet = workbook.Sheets[workbook.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
    if (!rows.length) return [];

    var keys = Object.keys(rows[0]);
    var kCustomer = findKey(keys, HEADER_MAP.customer);
    var kBox = findKey(keys, HEADER_MAP.box);
    var kIntake = findKey(keys, HEADER_MAP.intakeDate);
    var kShredDate = findKey(keys, HEADER_MAP.shredDate);
    var kShredded = findKey(keys, HEADER_MAP.shredded);

    return rows.map(function(r){
      return {
        customer: normKey(kCustomer ? r[kCustomer] : ''),
        box: normKey(kBox ? r[kBox] : ''),
        intakeDate: kIntake ? r[kIntake] : null,
        shredDate: kShredDate ? r[kShredDate] : null,
        shredded: kShredded ? !!r[kShredded] : false
      };
    }).filter(function(r){ return r.customer !== '' && r.box !== ''; });
  }

  // --- IndexedDB storage (handles far more data than localStorage's ~5-10MB cap) ---

  function openDB(){
    return new Promise(function(resolve, reject){
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function(){
        req.result.createObjectStore(STORE_NAME);
      };
      req.onsuccess = function(){ resolve(req.result); };
      req.onerror = function(){ reject(req.error); };
    });
  }

  function idbPut(key, value){
    return openDB().then(function(db){
      return new Promise(function(resolve, reject){
        var tx = db.transaction(STORE_NAME, 'readwrite');
        tx.objectStore(STORE_NAME).put(value, key);
        tx.oncomplete = function(){ resolve(); };
        tx.onerror = function(){ reject(tx.error); };
      });
    });
  }

  function idbGet(key){
    return openDB().then(function(db){
      return new Promise(function(resolve, reject){
        var tx = db.transaction(STORE_NAME, 'readonly');
        var req = tx.objectStore(STORE_NAME).get(key);
        req.onsuccess = function(){ resolve(req.result || null); };
        req.onerror = function(){ reject(req.error); };
      });
    });
  }

  function buildIndex(records){
    var map = new Map();
    for (var i = 0; i < records.length; i++){
      var r = records[i];
      map.set(r.customer + '|' + r.box, r);
    }
    return map;
  }

  function saveData(records, filename){
    var payload = { records: records, filename: filename || '', updatedAt: new Date() };
    return idbPut(DATA_KEY, payload).then(function(){
      cachedPayload = payload;
      cachedIndex = buildIndex(records);
      return payload;
    });
  }

  function ensureLoaded(){
    if (cachedPayload) return Promise.resolve(cachedPayload);
    return idbGet(DATA_KEY).then(function(payload){
      cachedPayload = payload;
      cachedIndex = payload ? buildIndex(payload.records) : null;
      return payload;
    });
  }

  function renderStatus(payload){
    if (!payload || !payload.records || !payload.records.length){
      dataStatus.textContent = 'לא נטען מאגר עדיין';
      return;
    }
    var d = new Date(payload.updatedAt);
    dataStatus.innerHTML = 'המאגר עודכן: <b>' + d.toLocaleDateString('he-IL') + ' ' +
      d.toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) + '</b>' +
      ' &nbsp;|&nbsp; <b>' + payload.records.length + '</b> תיבות';
  }

  fileInput.addEventListener('change', function(e){
    var file = e.target.files[0];
    if (!file) return;
    dataStatus.textContent = 'טוען ומעבד קובץ...';
    var reader = new FileReader();
    reader.onload = function(ev){
      try {
        var data = new Uint8Array(ev.target.result);
        var workbook = XLSX.read(data, { type: 'array', cellDates: true });
        var records = parseWorkbook(workbook);
        if (!records.length){
          alert('לא נמצאו רשומות תקינות בקובץ. ודא שהעמודות תואמות (מס\' לקוח, מס\' תיבה לקוח, תאריך הכנסה, תאריך גריסה, נגרס).');
          renderStatus(cachedPayload);
          return;
        }
        saveData(records, file.name).then(function(payload){
          renderStatus(payload);
          alert('נטענו ' + records.length + ' תיבות בהצלחה.');
        }).catch(function(err){
          alert('שגיאה בשמירת הנתונים: ' + err.message);
          renderStatus(cachedPayload);
        });
      } catch (err){
        alert('שגיאה בקריאת הקובץ: ' + err.message);
        renderStatus(cachedPayload);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  function showResult(state, title, sub){
    resultEl.className = state;
    resultEl.innerHTML = '<div class="title">' + title + '</div>' +
      (sub ? '<div class="sub">' + sub + '</div>' : '');
  }

  function doCheck(){
    ensureLoaded().then(function(payload){
      if (!payload || !payload.records.length){
        showResult('not-found', '⚠️ לא נטען מאגר', 'יש לטעון קודם קובץ אקסל');
        return;
      }
      var cust = normKey(customerInput.value);
      var box = normKey(boxInput.value);
      if (!cust || !box){
        showResult('not-found', '⚠️ יש להזין מספר לקוח ומספר תיבה', '');
        return;
      }
      var match = cachedIndex.get(cust + '|' + box);
      if (!match){
        showResult('not-found', '❓ לא נמצאה תיבה כזו', 'בדוק את מספר הלקוח ומספר התיבה');
        return;
      }
      if (match.shredded){
        showResult('found-shredded', '✗ תיבה לגריסה',
          match.shredDate ? 'תאריך גריסה: ' + formatDate(match.shredDate) : '');
      } else {
        showResult('found-active', '✔ התיבה פעילה במאגר',
          match.intakeDate ? 'תאריך הכנסה: ' + formatDate(match.intakeDate) : '');
      }
    });
  }

  checkBtn.addEventListener('click', doCheck);

  customerInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); boxInput.focus(); }
  });
  boxInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); doCheck(); }
  });

  ensureLoaded().then(renderStatus);

  if ('serviceWorker' in navigator){
    window.addEventListener('load', function(){
      navigator.serviceWorker.register('sw.js').catch(function(){});
    });
  }
})();
