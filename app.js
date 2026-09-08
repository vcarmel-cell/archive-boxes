(function(){
  'use strict';

  var DB_NAME = 'archiveBoxesDB2';
  var DB_VERSION = 1;
  var RECORDS_STORE = 'records';
  var META_STORE = 'meta';
  var META_KEY = 'info';
  var BATCH_SIZE = 2000;

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

  // Small metadata cache only (filename / updatedAt / count) - the records
  // themselves stay in IndexedDB and are looked up by key, never held in bulk
  // in memory. This avoids Safari/iOS's QuotaExceededError, which happens
  // when a single IndexedDB value is one huge array of many records - even
  // when the actual data size is well within the device's free storage.
  var cachedMeta = null;

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

  // --- IndexedDB storage: one small row per box, not one giant array ---

  function openDB(){
    return new Promise(function(resolve, reject){
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function(){
        var db = req.result;
        if (!db.objectStoreNames.contains(RECORDS_STORE)) db.createObjectStore(RECORDS_STORE);
        if (!db.objectStoreNames.contains(META_STORE)) db.createObjectStore(META_STORE);
      };
      req.onsuccess = function(){ resolve(req.result); };
      req.onerror = function(){ reject(req.error); };
    });
  }

  function runTx(storeName, mode, fn){
    return openDB().then(function(db){
      return new Promise(function(resolve, reject){
        var tx = db.transaction(storeName, mode);
        var store = tx.objectStore(storeName);
        var result = fn(store);
        tx.oncomplete = function(){ resolve(result); };
        tx.onerror = function(){ reject(tx.error); };
        tx.onabort = function(){ reject(tx.error || new Error('Transaction aborted')); };
      });
    });
  }

  function clearStore(storeName){
    return runTx(storeName, 'readwrite', function(store){ store.clear(); });
  }

  function putBatch(records, startIdx, endIdx){
    return runTx(RECORDS_STORE, 'readwrite', function(store){
      for (var i = startIdx; i < endIdx; i++){
        var r = records[i];
        store.put(r, r.customer + '|' + r.box);
      }
    });
  }

  function saveData(records, filename, onProgress){
    return clearStore(RECORDS_STORE).then(function(){
      var i = 0;
      function next(){
        if (i >= records.length) return Promise.resolve();
        var end = Math.min(i + BATCH_SIZE, records.length);
        return putBatch(records, i, end).then(function(){
          i = end;
          if (onProgress) onProgress(i, records.length);
          return next();
        });
      }
      return next();
    }).then(function(){
      var meta = { filename: filename || '', updatedAt: new Date(), count: records.length };
      return runTx(META_STORE, 'readwrite', function(store){
        store.put(meta, META_KEY);
      }).then(function(){
        cachedMeta = meta;
        return meta;
      });
    });
  }

  function getRecord(customer, box){
    return runTx(RECORDS_STORE, 'readonly', function(store){
      return store.get(customer + '|' + box);
    }).then(function(req){ return req.result; });
  }

  function loadMeta(){
    if (cachedMeta) return Promise.resolve(cachedMeta);
    return runTx(META_STORE, 'readonly', function(store){
      return store.get(META_KEY);
    }).then(function(req){
      cachedMeta = req.result || null;
      return cachedMeta;
    });
  }

  function renderStatus(meta){
    if (!meta || !meta.count){
      dataStatus.textContent = 'לא נטען מאגר עדיין';
      return;
    }
    var d = new Date(meta.updatedAt);
    dataStatus.innerHTML = 'המאגר עודכן: <b>' + d.toLocaleDateString('he-IL') + ' ' +
      d.toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) + '</b>' +
      ' &nbsp;|&nbsp; <b>' + meta.count + '</b> תיבות';
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
          renderStatus(cachedMeta);
          return;
        }
        saveData(records, file.name, function(done, total){
          dataStatus.textContent = 'שומר במאגר... ' + done + ' מתוך ' + total;
        }).then(function(meta){
          renderStatus(meta);
          alert('נטענו ' + records.length + ' תיבות בהצלחה.');
        }).catch(function(err){
          alert('שגיאה בשמירת הנתונים: ' + err.message);
          renderStatus(cachedMeta);
        });
      } catch (err){
        alert('שגיאה בקריאת הקובץ: ' + err.message);
        renderStatus(cachedMeta);
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
    loadMeta().then(function(meta){
      if (!meta || !meta.count){
        showResult('not-found', '⚠️ לא נטען מאגר', 'יש לטעון קודם קובץ אקסל');
        return;
      }
      var cust = normKey(customerInput.value);
      var box = normKey(boxInput.value);
      if (!cust || !box){
        showResult('not-found', '⚠️ יש להזין מספר לקוח ומספר תיבה', '');
        return;
      }
      return getRecord(cust, box).then(function(match){
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
    });
  }

  checkBtn.addEventListener('click', doCheck);

  customerInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); boxInput.focus(); }
  });
  boxInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); doCheck(); }
  });

  loadMeta().then(renderStatus);

  if ('serviceWorker' in navigator){
    window.addEventListener('load', function(){
      navigator.serviceWorker.register('sw.js').catch(function(){});
    });
  }
})();
