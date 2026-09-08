(function(){
  'use strict';

  var STORAGE_KEY = 'archiveBoxesData';

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

  function saveData(records, filename){
    var payload = {
      records: records,
      filename: filename || '',
      updatedAt: new Date().toISOString()
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    return payload;
  }

  function loadData(){
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (e) {
      return null;
    }
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
    var reader = new FileReader();
    reader.onload = function(ev){
      try {
        var data = new Uint8Array(ev.target.result);
        var workbook = XLSX.read(data, { type: 'array', cellDates: true });
        var records = parseWorkbook(workbook);
        if (!records.length){
          alert('לא נמצאו רשומות תקינות בקובץ. ודא שהעמודות תואמות (מס\' לקוח, מס\' תיבה לקוח, תאריך הכנסה, תאריך גריסה, נגרס).');
          return;
        }
        var payload = saveData(records, file.name);
        renderStatus(payload);
        alert('נטענו ' + records.length + ' תיבות בהצלחה.');
      } catch (err){
        alert('שגיאה בקריאת הקובץ: ' + err.message);
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
    var payload = loadData();
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
    var match = payload.records.find(function(r){
      return r.customer === cust && r.box === box;
    });
    if (!match){
      showResult('not-found', '❓ לא נמצאה תיבה כזו', 'בדוק את מספר הלקוח ומספר התיבה');
      return;
    }
    if (match.shredded){
      showResult('found-shredded', '✗ התיבה נגרסה',
        match.shredDate ? 'תאריך גריסה: ' + formatDate(match.shredDate) : '');
    } else {
      showResult('found-active', '✔ התיבה פעילה במאגר',
        match.intakeDate ? 'תאריך הכנסה: ' + formatDate(match.intakeDate) : '');
    }
  }

  checkBtn.addEventListener('click', doCheck);

  customerInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); boxInput.focus(); }
  });
  boxInput.addEventListener('keydown', function(e){
    if (e.key === 'Enter'){ e.preventDefault(); doCheck(); }
  });

  renderStatus(loadData());

  if ('serviceWorker' in navigator){
    window.addEventListener('load', function(){
      navigator.serviceWorker.register('sw.js').catch(function(){});
    });
  }
})();
