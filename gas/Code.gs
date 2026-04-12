/**
 * Apps Script API for GitHub dashboard.
 * Reads only these tabs:
 * - F_피벗(요일), F_피벗(본부), F_피벗
 * - I_피벗(요일), I_피벗(본부), I_피벗
 * - 대시보드
 */
const CONFIG = {
  spreadsheetId: '1FPiXIDJHPjd4-96MKYXoUe-9Bvu70mD37noA55dXR74',
  cacheKey: 'ai_dashboard_payload_v2',
  cacheTtlSec: 180,
  gidMap: {
    fDow: 361980346,
    fDept: 2079191863,
    fPivot: 2091751942,
    iDow: 616006980,
    iDept: 1713667830,
    iPivot: 1036774965,
    dashboard: 0,
  },
};

function doGet(e) {
  const callback = e && e.parameter ? String(e.parameter.callback || '') : '';
  try {
    const payload = getPayload_();
    return output_(payload, callback);
  } catch (err) {
    return output_({
      error: String(err && err.message ? err.message : err),
      generatedAt: new Date().toISOString(),
    }, callback);
  }
}

function output_(payload, callback) {
  const body = JSON.stringify(payload);
  if (callback) {
    const safe = callback.replace(/[^\w.$]/g, '');
    return ContentService.createTextOutput(safe + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function getPayload_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const matrices = readMatrices_();
  const kpi = extractKpi_(matrices.dashboard);
  const corpTotals = extractCorpTotals_(matrices.dashboard);

  const fast = parsePlatform_(matrices.fPivot, matrices.fDept, matrices.fDow);
  const inf = parsePlatform_(matrices.iPivot, matrices.iDept, matrices.iDow);

  const corpNames = unique_([
    Object.keys(corpTotals),
    Object.keys(fast.corpUsers),
    Object.keys(inf.corpUsers),
  ]);

  const fcCorps = buildCorpRows_(corpNames, corpTotals, fast.corpUsers);
  const infCorps = buildCorpRows_(corpNames, corpTotals, inf.corpUsers);

  const totalHeadcount = kpi.totalHeadcount || sum_(Object.keys(corpTotals).map(function (name) {
    return corpTotals[name];
  }));
  const totalUsers = kpi.totalUsers || 0;
  const usageRate = kpi.usageRate || (totalHeadcount > 0 ? totalUsers / totalHeadcount * 100 : 0);

  const payload = {
    updatedAt: kpi.updatedAt || '',
    periodLabel: kpi.periodLabel || '',
    kpi: {
      totalHeadcount: Math.round(totalHeadcount),
      totalUsers: Math.round(totalUsers),
      usageRate: round1_(usageRate),
    },
    corpTotals: corpTotals,
    fcDaily: fast.daily,
    infDaily: inf.daily,
    fcCorps: fcCorps,
    infCorps: infCorps,
    fcDept: fast.depts,
    infDept: inf.depts,
    fcDow: {
      dow: fast.dow,
      topCourses: fast.topCourses,
    },
    infDow: {
      dow: inf.dow,
      topCourses: inf.topCourses,
    },
    generatedAt: new Date().toISOString(),
  };

  cache.put(CONFIG.cacheKey, JSON.stringify(payload), CONFIG.cacheTtlSec);
  return payload;
}

function readMatrices_() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const byId = {};
  ss.getSheets().forEach(function (s) {
    byId[s.getSheetId()] = s;
  });

  function valuesByGid(gid) {
    const sheet = byId[gid];
    return sheet ? sheet.getDataRange().getDisplayValues() : [];
  }

  return {
    fDow: valuesByGid(CONFIG.gidMap.fDow),
    fDept: valuesByGid(CONFIG.gidMap.fDept),
    fPivot: valuesByGid(CONFIG.gidMap.fPivot),
    iDow: valuesByGid(CONFIG.gidMap.iDow),
    iDept: valuesByGid(CONFIG.gidMap.iDept),
    iPivot: valuesByGid(CONFIG.gidMap.iPivot),
    dashboard: valuesByGid(CONFIG.gidMap.dashboard),
  };
}

function parsePlatform_(pivotRows, deptRows, dowRows) {
  return {
    daily: parseDailyTrend_(pivotRows),
    corpUsers: parseCorpUsers_(deptRows),
    depts: parseDeptRows_(deptRows),
    dow: parseDow_(dowRows),
    topCourses: parseTopCourses_(dowRows),
  };
}

function parseDailyTrend_(rows) {
  if (!rows || rows.length < 2) {
    return [];
  }
  const out = [];
  for (var i = 1; i < rows.length; i += 1) {
    var rawDay = clean_(rows[i][2]);
    if (!/^\d{1,2}-\d{1,2}월$/.test(rawDay)) {
      continue;
    }
    var dayLabel = formatPivotDayLabel_(rawDay);
    out.push({
      dateLabel: dayLabel,
      hours: round2_(toNumber_(rows[i][3])),
      users: Math.round(toNumber_(rows[i][4])),
    });
  }
  return out;
}

function parseCorpUsers_(rows) {
  var map = {};
  if (!rows || rows.length < 2) {
    return map;
  }
  var currentCorp = '';
  for (var i = 1; i < rows.length; i += 1) {
    var corp = clean_(rows[i][0]);
    var dept = clean_(rows[i][1]);
    var users = Math.round(toNumber_(rows[i][3]));
    if (corp) {
      currentCorp = corp;
    }
    if (!currentCorp || isTotalLabel_(currentCorp)) {
      continue;
    }
    if (!dept) {
      map[currentCorp] = users;
    }
  }
  return map;
}

function parseDeptRows_(rows) {
  var out = [];
  if (!rows || rows.length < 2) {
    return out;
  }
  var currentCorp = '';
  for (var i = 1; i < rows.length; i += 1) {
    var corp = clean_(rows[i][0]);
    var dept = clean_(rows[i][1]);
    if (corp) {
      currentCorp = corp;
    }
    if (!currentCorp || !dept || isTotalLabel_(currentCorp) || isTotalLabel_(dept)) {
      continue;
    }
    var hours = round2_(toNumber_(rows[i][2]));
    var users = Math.round(toNumber_(rows[i][3]));
    out.push({
      corp: currentCorp,
      dept: dept,
      deptLabel: shortCorp_(currentCorp) + ' | ' + dept,
      hours: hours,
      users: users,
    });
  }
  // Defensive fallback: when corp/dept structured parsing fails,
  // recover from generic [corp, dept, hours, users] rows.
  if (!out.length) {
    for (var j = 1; j < rows.length; j += 1) {
      var corp2 = clean_(rows[j][0]);
      var dept2 = clean_(rows[j][1]);
      var hours2 = round2_(toNumber_(rows[j][2]));
      var users2 = Math.round(toNumber_(rows[j][3]));
      if (!corp2 || !dept2 || isTotalLabel_(corp2) || isTotalLabel_(dept2)) {
        continue;
      }
      out.push({
        corp: corp2,
        dept: dept2,
        deptLabel: shortCorp_(corp2) + ' | ' + dept2,
        hours: hours2,
        users: users2,
      });
    }
  }

  out = out.filter(function (row) {
    return row.hours > 0 || row.users > 0;
  });
  out.sort(function (a, b) { return b.hours - a.hours; });
  return out.slice(0, 30);
}

function parseDow_(rows) {
  var order = ['일', '월', '화', '수', '목', '금', '토'];
  var byPct = {};
  var byHours = {};

  if (!rows || rows.length < 2) {
    return order.map(function (day) { return { day: day, pct: 0 }; });
  }

  for (var i = 1; i < rows.length; i += 1) {
    var dayA = clean_(rows[i][8]);
    var pct = toPercent_(rows[i][9]);
    if (order.indexOf(dayA) >= 0 && pct > 0) {
      byPct[dayA] = pct;
    }

    var dayB = clean_(rows[i][5]);
    var h = toNumber_(rows[i][6]);
    if (order.indexOf(dayB) >= 0 && h > 0) {
      byHours[dayB] = (byHours[dayB] || 0) + h;
    }
  }

  if (Object.keys(byPct).length < 5 && Object.keys(byHours).length > 0) {
    var total = sum_(Object.keys(byHours).map(function (key) { return byHours[key]; }));
    order.forEach(function (day) {
      byPct[day] = total > 0 ? round1_(byHours[day] / total * 100) : 0;
    });
  }

  return order.map(function (day) {
    return { day: day, pct: round1_(byPct[day] || 0) };
  });
}

function parseTopCourses_(rows) {
  var out = [];
  if (!rows || rows.length < 2) {
    return out;
  }
  for (var i = 1; i < rows.length; i += 1) {
    var name = clean_(rows[i][0]);
    if (!name || /course_name|강의명|총계|합계/i.test(name)) {
      continue;
    }
    var users = Math.round(toNumber_(rows[i][1]));
    var hours = round2_(toNumber_(rows[i][2]));
    if (users <= 0 && hours <= 0) {
      continue;
    }
    out.push({ name: name, users: users, hours: hours });
  }
  out.sort(function (a, b) { return b.users - a.users; });
  return out.slice(0, 5);
}

function extractKpi_(rows) {
  var result = {
    updatedAt: '',
    periodLabel: '',
    totalHeadcount: 0,
    totalUsers: 0,
    usageRate: 0,
  };
  if (!rows || !rows.length) {
    return result;
  }

  var upd = findLabeledValue_(rows, /(update|업데이트)/i);
  var total = findLabeledValue_(rows, /(대상\s*총원|대상총원|총원)/i);
  var users = findLabeledValue_(rows, /(수강인원|학습인원|이용인원)/i);
  var rate = findLabeledValue_(rows, /(사용률|활용률|참여율)/i);
  var period = findLabeledValue_(rows, /(기간)/i);

  result.updatedAt = clean_(upd);
  result.periodLabel = clean_(period);
  result.totalHeadcount = toNumber_(total);
  result.totalUsers = toNumber_(users);
  result.usageRate = toPercent_(rate);

  if (!result.updatedAt) {
    result.updatedAt = extractDateLike_(rows);
  }
  return result;
}

function extractCorpTotals_(rows) {
  var map = {};
  if (!rows || !rows.length) {
    return map;
  }
  for (var i = 0; i < rows.length; i += 1) {
    var corp = clean_(rows[i][28]);
    var total = Math.round(toNumber_(rows[i][29]));
    if (!corp || total <= 0 || isTotalLabel_(corp)) {
      continue;
    }
    map[corp] = total;
  }
  return map;
}

function buildCorpRows_(corpNames, corpTotals, userMap) {
  var rows = corpNames.map(function (name) {
    var total = corpTotals[name] || 0;
    var users = userMap[name] || 0;
    var rate = total > 0 ? users / total * 100 : 0;
    return {
      name: name,
      total: total,
      users: users,
      rate: round1_(rate),
    };
  });
  rows.sort(function (a, b) {
    if (b.total !== a.total) {
      return b.total - a.total;
    }
    return a.name.localeCompare(b.name);
  });
  return rows;
}

function findLabeledValue_(rows, regex) {
  for (var r = 0; r < rows.length; r += 1) {
    for (var c = 0; c < rows[r].length; c += 1) {
      var label = clean_(rows[r][c]);
      if (!label || !regex.test(label)) {
        continue;
      }
      for (var n = c + 1; n < Math.min(rows[r].length, c + 6); n += 1) {
        var right = clean_(rows[r][n]);
        if (right) {
          return right;
        }
      }
      for (var m = r + 1; m < Math.min(rows.length, r + 4); m += 1) {
        var down = clean_(rows[m][c]);
        if (down) {
          return down;
        }
      }
    }
  }
  return '';
}

function extractDateLike_(rows) {
  var re = /\d{4}[-./]\d{1,2}[-./]\d{1,2}/;
  for (var r = 0; r < rows.length; r += 1) {
    for (var c = 0; c < rows[r].length; c += 1) {
      var cell = clean_(rows[r][c]);
      if (re.test(cell)) {
        return cell;
      }
    }
  }
  return '';
}

function toNumber_(v) {
  if (v === null || v === undefined) {
    return 0;
  }
  var s = String(v).replace(/,/g, '').replace(/[^0-9.\-]/g, '').trim();
  if (!s) {
    return 0;
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function toPercent_(v) {
  if (v === null || v === undefined) {
    return 0;
  }
  var raw = String(v).trim();
  if (!raw) {
    return 0;
  }
  if (raw.indexOf('%') >= 0) {
    return round1_(toNumber_(raw));
  }
  var n = toNumber_(raw);
  if (n <= 1) {
    n = n * 100;
  }
  return round1_(n);
}

function clean_(v) {
  return String(v || '').trim();
}

function isTotalLabel_(text) {
  return /(총|합계|전체)/.test(clean_(text));
}

function shortCorp_(name) {
  if (name === 'KRAFTON HQ') {
    return 'HQ';
  }
  if (name === 'PUBG STUDIOS') {
    return 'PUBG';
  }
  return name;
}

function sum_(arr) {
  return arr.reduce(function (acc, n) { return acc + (n || 0); }, 0);
}

function unique_(listOfLists) {
  var s = {};
  listOfLists.forEach(function (arr) {
    arr.forEach(function (v) {
      var key = clean_(v);
      if (key) {
        s[key] = true;
      }
    });
  });
  return Object.keys(s);
}

function round1_(n) {
  return Math.round((n || 0) * 10) / 10;
}

function round2_(n) {
  return Math.round((n || 0) * 100) / 100;
}

function formatPivotDayLabel_(value) {
  var s = clean_(value);
  var m = /^(\d{1,2})-(\d{1,2})월$/.exec(s);
  if (!m) {
    return s;
  }
  var day = parseInt(m[1], 10);
  var month = parseInt(m[2], 10);
  if (isNaN(day) || isNaN(month)) {
    return s;
  }
  return month + '월 ' + day + '일';
}
