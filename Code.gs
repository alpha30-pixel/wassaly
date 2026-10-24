/**
 * وصّلي — سيستم شركة الشحن (Google Apps Script + Google Sheets)
 * ------------------------------------------------------------
 * الخطوات:
 * 1) اعمل Google Sheet جديد وحط الـ ID بتاعه في SHEET_ID، واعمل مشروع Apps Script (script.new) والصق الكود ده.
 * 2) شغّل الدالة setup مرة واحدة (هتعمل الشيتات وحساب المدير admin، والباسورد هيظهر في Execution log وفي شيت Users).
 * 3) Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone ← Deploy.
 * 4) خد لينك الـ Web app وحطه في API_URL في صفحة الموظفين وصفحة العميل.
 */

var SHEET_ID = '17lSFsHoKWhye74TXLCZrRlaiYAIdp00_jIVyFtGrvzc'; // رقم جوجل شيت (من اللينك)
var RETURN_FEE_PCT = 50;                    // نسبة مصاريف الشحن على المرتجع (%)
var TOKEN_HOURS = 12;

var SHEETS = {
  Users: ['id', 'username', 'name', 'role', 'phone', 'salt', 'hash', 'active', 'created'],
  Merchants: ['id', 'name', 'phone', 'address', 'governorate', 'notes', 'created'],
  Shipments: ['code', 'created', 'createdBy', 'merchantId', 'merchantName', 'senderPhone', 'receiverName', 'receiverPhone',
    'governorate', 'city', 'address', 'pieces', 'weight', 'content', 'cod', 'fee', 'status', 'courierId', 'courierName',
    'attempts', 'notes', 'updated', 'deliveredAt', 'collected', 'settled'],
  History: ['ts', 'code', 'status', 'note', 'by'],
  Ledger: ['id', 'ts', 'type', 'amount', 'party', 'partyName', 'ref', 'note', 'by'],
  Rates: ['governorate', 'fee', 'days']
};

var STATUS = {
  new: 'تم تسجيل الشحنة', picked: 'تم الاستلام من التاجر', inbound: 'وصلت المخزن',
  outbound: 'خرجت للتوصيل', delivered: 'تم التسليم', failed: 'محاولة توصيل فاشلة',
  returning: 'مرتجع في الطريق للتاجر', returned: 'تم الإرجاع للتاجر', cancelled: 'ملغية'
};

var GOVS = ['القاهرة', 'الجيزة', 'الإسكندرية', 'القليوبية', 'الشرقية', 'الدقهلية', 'الغربية', 'المنوفية', 'البحيرة',
  'كفر الشيخ', 'دمياط', 'بورسعيد', 'الإسماعيلية', 'السويس', 'الفيوم', 'بني سويف', 'المنيا', 'أسيوط', 'سوهاج',
  'قنا', 'الأقصر', 'أسوان', 'البحر الأحمر', 'مطروح', 'الوادي الجديد', 'شمال سيناء', 'جنوب سيناء'];

/* صلاحيات كل دور */
var ROLES = {
  admin: ['*'],
  accountant: ['dashboard', 'shipments.view', 'merchants', 'accounting', 'reports'],
  ops: ['dashboard', 'shipments.view', 'shipments.edit', 'warehouse', 'dispatch', 'merchants', 'reports'],
  courier: ['courier']
};

/* ===================== التجهيز ===================== */
function setup() {
  var ss = ss_();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  var pass = '';
  if (!rows_('Users').length) {
    var salt = uid_(); pass = uid_().slice(0, 8);
    append_('Users', { id: uid_(), username: 'admin', name: 'المدير', role: 'admin', phone: '', salt: salt,
      hash: hash_(pass, salt), active: true, created: now_() });
  }
  if (!rows_('Rates').length) {
    GOVS.forEach(function (g) {
      var near = ['القاهرة', 'الجيزة', 'القليوبية'].indexOf(g) >= 0;
      append_('Rates', { governorate: g, fee: near ? 50 : 70, days: near ? 1 : 2 });
    });
  }
  var msg = pass ? 'تم التجهيز ✔ — اسم المستخدم: admin  —  الباسورد: ' + pass + '  (غيّره من الإعدادات بعد أول دخول)' : 'الشيتات جاهزة ✔';
  Logger.log(msg);
  if (pass) sheet_('Users').getRange(1, SHEETS.Users.length + 2).setValue('أول باسورد للمدير: ' + pass + ' — امسح الخانة دي بعد ما تغيّره');
  return msg;
}

/* ===================== نقطة الدخول ===================== */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'track') return json_(handle_({ action: 'track', code: p.code || p.q }));
  return json_({ ok: true, app: 'wassaly', time: now_() });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'بيانات غلط' }); }
  return json_(handle_(body));
}

function handle_(b) {
  try {
    var a = b.action;
    if (a === 'track') return ok_(trackAny_(b.code || b.q));
    if (a === 'login') return ok_(login_(b.username, b.password));
    if (a === 'rates') return ok_(rows_('Rates'));
    var u = auth_(b.token);
    var lock = LockService.getScriptLock();
    var write = WRITE_ACTIONS.indexOf(a) >= 0;
    if (write) lock.waitLock(20000);
    try {
      var fn = API[a];
      if (!fn) throw new Error('أمر غير معروف: ' + a);
      return ok_(fn(u, b));
    } finally { if (write) lock.releaseLock(); }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

var WRITE_ACTIONS = ['createShipment', 'updateShipment', 'setStatus', 'scan', 'assign', 'deliver', 'fail', 'saveMerchant',
  'addLedger', 'handover', 'settleMerchant', 'saveUser', 'saveRates', 'changePassword', 'logout'];

/* ===================== الأوامر ===================== */
var API = {
  me: function (u) { var p = pub_(u); p.rates = rows_('Rates'); return p; },
  logout: function (u, b) { CacheService.getScriptCache().remove('t_' + b.token); return true; },
  changePassword: function (u, b) {
    var users = rows_('Users'), me = find_(users, 'id', u.id);
    if (hash_(b.old || '', me.salt) !== me.hash) throw new Error('الباسورد القديم غلط');
    if (!b.password || String(b.password).length < 6) throw new Error('الباسورد لازم 6 حروف على الأقل');
    var salt = uid_(); update_('Users', 'id', u.id, { salt: salt, hash: hash_(b.password, salt) }); return true;
  },

  dashboard: function (u) {
    need_(u, 'dashboard');
    var sh = rows_('Shipments'), today = day_(new Date()), by = {}, codToday = 0, feeToday = 0, newToday = 0;
    sh.forEach(function (s) {
      by[s.status] = (by[s.status] || 0) + 1;
      if (day_(s.created) === today) newToday++;
      if (s.status === 'delivered' && day_(s.deliveredAt) === today) { codToday += num_(s.cod); feeToday += num_(s.fee); }
    });
    var hist = rows_('History').slice(-12).reverse();
    var daily = {};
    for (var i = 13; i >= 0; i--) { var d = new Date(); d.setDate(d.getDate() - i); daily[day_(d)] = { day: day_(d), created: 0, delivered: 0 }; }
    sh.forEach(function (s) {
      var c = day_(s.created), dd = day_(s.deliveredAt);
      if (daily[c]) daily[c].created++;
      if (s.status === 'delivered' && daily[dd]) daily[dd].delivered++;
    });
    return { total: sh.length, byStatus: by, newToday: newToday, codToday: codToday, feeToday: feeToday,
      custody: courierBalances_().reduce(function (t, c) { return t + c.balance; }, 0),
      merchantDue: merchantDues_().reduce(function (t, m) { return t + m.due; }, 0),
      recent: hist, daily: Object.keys(daily).map(function (k) { return daily[k]; }) };
  },

  listShipments: function (u, b) {
    need_(u, 'shipments.view');
    var f = b.filter || {}, q = String(f.q || '').trim().toLowerCase();
    var list = rows_('Shipments').filter(function (s) {
      if (f.status && s.status !== f.status) return false;
      if (f.merchantId && s.merchantId !== f.merchantId) return false;
      if (f.courierId && s.courierId !== f.courierId) return false;
      if (f.governorate && s.governorate !== f.governorate) return false;
      if (f.from && day_(s.created) < f.from) return false;
      if (f.to && day_(s.created) > f.to) return false;
      if (q) {
        var qd = phoneKey_(q), hay = [s.code, s.receiverName, s.receiverPhone, s.merchantName, s.city, s.senderPhone].join(' ').toLowerCase();
        var phoneHit = qd.length >= 6 && (phoneKey_(s.receiverPhone).indexOf(qd) >= 0 || phoneKey_(s.senderPhone).indexOf(qd) >= 0);
        if (hay.indexOf(q) < 0 && !phoneHit) return false;
      }
      return true;
    });
    list.reverse();
    return { total: list.length, items: list.slice(0, num_(f.limit) || 300) };
  },

  getShipment: function (u, b) {
    if (!can_(u, 'shipments.view') && !can_(u, 'courier')) throw new Error('مش مسموح');
    var s = find_(rows_('Shipments'), 'code', norm_(b.code));
    if (!s) throw new Error('الشحنة مش موجودة');
    if (u.role === 'courier' && s.courierId !== u.id) throw new Error('الشحنة دي مش معاك');
    s.history = rows_('History').filter(function (h) { return h.code === s.code; });
    return s;
  },

  createShipment: function (u, b) {
    need_(u, 'shipments.edit');
    var d = b.data || {};
    if (!d.receiverName || !d.receiverPhone || !d.governorate) throw new Error('اسم المستلم وتليفونه والمحافظة مطلوبين');
    var m = d.merchantId ? find_(rows_('Merchants'), 'id', d.merchantId) : null;
    var rate = find_(rows_('Rates'), 'governorate', d.governorate);
    var s = {
      code: newCode_(), created: now_(), createdBy: u.name, merchantId: m ? m.id : '', merchantName: m ? m.name : (d.merchantName || ''),
      senderPhone: m ? m.phone : (d.senderPhone || ''), receiverName: d.receiverName, receiverPhone: String(d.receiverPhone),
      governorate: d.governorate, city: d.city || '', address: d.address || '', pieces: num_(d.pieces) || 1, weight: num_(d.weight) || 0,
      content: d.content || '', cod: num_(d.cod), fee: d.fee !== '' && d.fee != null ? num_(d.fee) : (rate ? num_(rate.fee) : 0),
      status: 'new', courierId: '', courierName: '', attempts: 0, notes: d.notes || '', updated: now_(), deliveredAt: '', collected: false, settled: false
    };
    append_('Shipments', s);
    log_(s.code, 'new', d.notes || '', u.name);
    return s;
  },

  updateShipment: function (u, b) {
    need_(u, 'shipments.edit');
    var code = norm_(b.code), d = b.data || {}, keep = {};
    ['receiverName', 'receiverPhone', 'governorate', 'city', 'address', 'pieces', 'weight', 'content', 'cod', 'fee', 'notes'].forEach(function (k) {
      if (d[k] !== undefined) keep[k] = ['pieces', 'weight', 'cod', 'fee'].indexOf(k) >= 0 ? num_(d[k]) : d[k];
    });
    keep.updated = now_();
    if (!update_('Shipments', 'code', code, keep)) throw new Error('الشحنة مش موجودة');
    log_(code, 'edit', 'تعديل بيانات', u.name);
    return find_(rows_('Shipments'), 'code', code);
  },

  setStatus: function (u, b) {
    need_(u, 'shipments.edit');
    if (!STATUS[b.status]) throw new Error('حالة غير معروفة');
    return setStatus_(norm_(b.code), b.status, b.note || '', u);
  },

  scan: function (u, b) {
    need_(u, 'warehouse');
    var map = { pickup: 'picked', in: 'inbound', out: 'outbound', ret: 'returning', retdone: 'returned' };
    var st = map[b.kind]; if (!st) throw new Error('نوع مسح غلط');
    var codes = [].concat(b.codes || b.code).map(norm_).filter(String), res = [];
    codes.forEach(function (c) {
      try {
        if (st === 'outbound' && b.courierId) assign_(c, b.courierId);
        setStatus_(c, st, b.note || '', u); res.push({ code: c, ok: true });
      } catch (e) { res.push({ code: c, ok: false, error: e.message }); }
    });
    return res;
  },

  assign: function (u, b) {
    need_(u, 'dispatch');
    var res = [];
    [].concat(b.codes).map(norm_).forEach(function (c) {
      try { assign_(c, b.courierId); res.push({ code: c, ok: true }); } catch (e) { res.push({ code: c, ok: false, error: e.message }); }
    });
    return res;
  },

  myShipments: function (u) {
    need_(u, 'courier');
    return rows_('Shipments').filter(function (s) {
      return s.courierId === u.id && ['outbound', 'failed', 'inbound', 'picked'].indexOf(s.status) >= 0;
    });
  },

  deliver: function (u, b) {
    if (!can_(u, 'courier') && !can_(u, 'shipments.edit')) throw new Error('مش مسموح');
    var code = norm_(b.code), s = find_(rows_('Shipments'), 'code', code);
    if (!s) throw new Error('الشحنة مش موجودة');
    if (u.role === 'courier' && s.courierId !== u.id) throw new Error('الشحنة دي مش معاك');
    if (s.status === 'delivered') throw new Error('الشحنة متسلّمة قبل كده');
    var amount = b.amount !== undefined && b.amount !== '' ? num_(b.amount) : num_(s.cod);
    update_('Shipments', 'code', code, { status: 'delivered', deliveredAt: now_(), updated: now_(), cod: amount });
    log_(code, 'delivered', b.note || '', u.name);
    if (amount > 0) {
      var cid = s.courierId || u.id, cname = s.courierName || u.name;
      append_('Ledger', { id: uid_(), ts: now_(), type: 'cod', amount: amount, party: cid, partyName: cname, ref: code, note: 'تحصيل عند التسليم', by: u.name });
    }
    return true;
  },

  fail: function (u, b) {
    if (!can_(u, 'courier') && !can_(u, 'shipments.edit')) throw new Error('مش مسموح');
    var code = norm_(b.code), s = find_(rows_('Shipments'), 'code', code);
    if (!s) throw new Error('الشحنة مش موجودة');
    if (u.role === 'courier' && s.courierId !== u.id) throw new Error('الشحنة دي مش معاك');
    update_('Shipments', 'code', code, { status: 'failed', attempts: num_(s.attempts) + 1, updated: now_() });
    log_(code, 'failed', b.reason || '', u.name);
    return true;
  },

  listMerchants: function (u) {
    if (!can_(u, 'merchants') && !can_(u, 'shipments.edit')) throw new Error('مش مسموح');
    var sh = rows_('Shipments');
    return rows_('Merchants').map(function (m) {
      m.count = sh.filter(function (s) { return s.merchantId === m.id; }).length; return m;
    });
  },

  saveMerchant: function (u, b) {
    if (!can_(u, 'merchants')) throw new Error('مش مسموح');
    var d = b.data || {};
    if (!d.name) throw new Error('اسم التاجر مطلوب');
    if (d.id && update_('Merchants', 'id', d.id, { name: d.name, phone: d.phone || '', address: d.address || '', governorate: d.governorate || '', notes: d.notes || '' })) return d;
    var m = { id: uid_(), name: d.name, phone: d.phone || '', address: d.address || '', governorate: d.governorate || '', notes: d.notes || '', created: now_() };
    append_('Merchants', m); return m;
  },

  listUsers: function (u) {
    return rows_('Users').map(pub_).filter(function (x) { return can_(u, 'users') || x.role === 'courier'; });
  },

  saveUser: function (u, b) {
    need_(u, 'users');
    var d = b.data || {};
    if (!ROLES[d.role]) throw new Error('دور غلط');
    var users = rows_('Users');
    if (d.id) {
      var patch = { name: d.name, role: d.role, phone: d.phone || '', active: d.active !== false && d.active !== 'false' };
      if (d.password) { if (String(d.password).length < 6) throw new Error('الباسورد 6 حروف على الأقل'); patch.salt = uid_(); patch.hash = hash_(d.password, patch.salt); }
      update_('Users', 'id', d.id, patch); return true;
    }
    if (!d.username || !d.password) throw new Error('اسم المستخدم والباسورد مطلوبين');
    if (String(d.password).length < 6) throw new Error('الباسورد 6 حروف على الأقل');
    if (find_(users, 'username', String(d.username).toLowerCase())) throw new Error('اسم المستخدم موجود قبل كده');
    var salt = uid_();
    append_('Users', { id: uid_(), username: String(d.username).toLowerCase(), name: d.name || d.username, role: d.role, phone: d.phone || '',
      salt: salt, hash: hash_(d.password, salt), active: true, created: now_() });
    return true;
  },

  saveRates: function (u, b) {
    need_(u, 'users');
    var sh = sheet_('Rates'); sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).clearContent();
    var vals = (b.rates || []).map(function (r) { return [r.governorate, num_(r.fee), num_(r.days)]; });
    if (vals.length) sh.getRange(2, 1, vals.length, 3).setValues(vals);
    dirty_('Rates');
    return true;
  },

  /* ---------- المحاسبة ---------- */
  accounting: function (u) {
    need_(u, 'accounting');
    var led = rows_('Ledger'), cash = 0, income = 0, expense = 0, payouts = 0, handover = 0;
    led.forEach(function (l) {
      var a = num_(l.amount);
      if (l.type === 'handover') { cash += a; handover += a; }
      if (l.type === 'payout') { cash -= a; payouts += a; }
      if (l.type === 'expense') { cash -= a; expense += a; }
      if (l.type === 'income') { cash += a; income += a; }
    });
    var fees = 0;
    rows_('Shipments').forEach(function (s) {
      if (s.status === 'delivered') fees += num_(s.fee);
      if (s.status === 'returned') fees += num_(s.fee) * RETURN_FEE_PCT / 100;
    });
    return { cash: cash, income: income, expense: expense, payouts: payouts, handover: handover, feesEarned: fees,
      profit: fees + income - expense, couriers: courierBalances_(), merchants: merchantDues_(), ledger: led.slice(-200).reverse() };
  },

  handover: function (u, b) {
    need_(u, 'accounting');
    var c = find_(rows_('Users'), 'id', b.courierId); if (!c) throw new Error('المندوب مش موجود');
    var amt = num_(b.amount); if (amt <= 0) throw new Error('المبلغ لازم أكبر من صفر');
    append_('Ledger', { id: uid_(), ts: now_(), type: 'handover', amount: amt, party: c.id, partyName: c.name, ref: '', note: b.note || 'تسليم عهدة', by: u.name });
    return true;
  },

  settleMerchant: function (u, b) {
    need_(u, 'accounting');
    var d = merchantDues_().filter(function (m) { return m.id === b.merchantId; })[0];
    if (!d || !d.codes.length) throw new Error('مفيش مستحقات للتاجر ده');
    var sh = sheet_('Shipments'), data = sh.getDataRange().getValues(), H = data[0], ci = H.indexOf('code'), si = H.indexOf('settled');
    for (var i = 1; i < data.length; i++) if (d.codes.indexOf(String(data[i][ci])) >= 0) sh.getRange(i + 1, si + 1).setValue(true);
    dirty_('Shipments');
    append_('Ledger', { id: uid_(), ts: now_(), type: 'payout', amount: d.due, party: d.id, partyName: d.name, ref: d.codes.join(','), note: b.note || 'تسوية تاجر', by: u.name });
    return d;
  },

  addLedger: function (u, b) {
    need_(u, 'accounting');
    if (['expense', 'income'].indexOf(b.type) < 0) throw new Error('نوع غلط');
    var amt = num_(b.amount); if (amt <= 0) throw new Error('المبلغ لازم أكبر من صفر');
    append_('Ledger', { id: uid_(), ts: now_(), type: b.type, amount: amt, party: '', partyName: b.party || '', ref: '', note: b.note || '', by: u.name });
    return true;
  },

  report: function (u, b) {
    need_(u, 'reports');
    var from = b.from || '0000', to = b.to || '9999';
    var sh = rows_('Shipments').filter(function (s) { var d = day_(s.created); return d >= from && d <= to; });
    var by = {}, gov = {}, cour = {}, cod = 0, fee = 0;
    sh.forEach(function (s) {
      by[s.status] = (by[s.status] || 0) + 1;
      gov[s.governorate] = (gov[s.governorate] || 0) + 1;
      if (s.courierName) { cour[s.courierName] = cour[s.courierName] || { name: s.courierName, total: 0, delivered: 0 }; cour[s.courierName].total++; if (s.status === 'delivered') cour[s.courierName].delivered++; }
      if (s.status === 'delivered') { cod += num_(s.cod); fee += num_(s.fee); }
    });
    var led = rows_('Ledger').filter(function (l) { var d = day_(l.ts); return d >= from && d <= to; });
    var exp = led.filter(function (l) { return l.type === 'expense'; }).reduce(function (t, l) { return t + num_(l.amount); }, 0);
    var inc = led.filter(function (l) { return l.type === 'income'; }).reduce(function (t, l) { return t + num_(l.amount); }, 0);
    return { count: sh.length, byStatus: by, byGov: gov, couriers: Object.keys(cour).map(function (k) { return cour[k]; }),
      cod: cod, fees: fee, expense: exp, income: inc, net: fee + inc - exp, rows: sh };
  }
};

/* ===================== منطق داخلي ===================== */
/* تتبّع بالكود أو برقم الموبايل */
function trackAny_(q) {
  q = String(q || '').trim();
  if (!q) throw new Error('اكتب رقم موبايلك أو كود الشحنة');
  if (/^ws/i.test(q)) return { items: [track_(q)] };
  var key = phoneKey_(q);
  if (key.length < 9) throw new Error('اكتب رقم الموبايل كامل');
  var list = rows_('Shipments').filter(function (s) { return phoneKey_(s.receiverPhone) === key; }).reverse().slice(0, 20);
  if (!list.length) throw new Error('مفيش شحنات على الرقم ده');
  return { items: list.map(function (s) { return track_(s.code); }) };
}
function phoneKey_(p) {
  var d = String(p || '').replace(/[٠-٩]/g, function (c) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(c); }).replace(/\D/g, '');
  if (d.indexOf('0020') === 0) d = d.slice(4); else if (d.indexOf('20') === 0 && d.length >= 12) d = d.slice(2);
  if (d.charAt(0) === '0') d = d.slice(1);
  return d;
}

function track_(code) {
  code = norm_(code);
  if (!code) throw new Error('اكتب كود الشحنة');
  var s = find_(rows_('Shipments'), 'code', code);
  if (!s) throw new Error('مفيش شحنة بالكود ده');
  var hist = rows_('History').filter(function (h) { return h.code === code && STATUS[h.status]; })
    .map(function (h) { return { ts: h.ts, status: h.status, label: STATUS[h.status], note: h.status === 'failed' ? h.note : '' }; });
  return { code: s.code, status: s.status, label: STATUS[s.status], governorate: s.governorate, city: s.city,
    receiver: mask_(s.receiverName), merchant: s.merchantName, cod: num_(s.cod), created: s.created,
    deliveredAt: s.deliveredAt, updated: s.updated, courier: s.status === 'outbound' ? (String(s.courierName).split(' ')[0] || '') : '', history: hist };
}

function setStatus_(code, status, note, u) {
  var s = find_(rows_('Shipments'), 'code', code);
  if (!s) throw new Error('الشحنة مش موجودة: ' + code);
  if (s.status === 'delivered' && status !== 'returning') throw new Error('الشحنة متسلّمة: ' + code);
  if (status === 'delivered') return API.deliver(u, { code: code, note: note });
  var patch = { status: status, updated: now_() };
  if (status === 'failed') patch.attempts = num_(s.attempts) + 1;
  update_('Shipments', 'code', code, patch);
  log_(code, status, note, u.name);
  return true;
}

function assign_(code, courierId) {
  var c = find_(rows_('Users'), 'id', courierId);
  if (!c || c.role !== 'courier') throw new Error('اختار مندوب صحيح');
  if (!update_('Shipments', 'code', code, { courierId: c.id, courierName: c.name, updated: now_() })) throw new Error('الشحنة مش موجودة: ' + code);
  log_(code, 'assign', 'اتسلّمت للمندوب ' + c.name, '');
}

function courierBalances_() {
  var bal = {};
  rows_('Users').forEach(function (u) { if (u.role === 'courier') bal[u.id] = { id: u.id, name: u.name, collected: 0, handed: 0, balance: 0, active: 0 }; });
  rows_('Ledger').forEach(function (l) {
    if (!bal[l.party]) return;
    if (l.type === 'cod') bal[l.party].collected += num_(l.amount);
    if (l.type === 'handover') bal[l.party].handed += num_(l.amount);
  });
  rows_('Shipments').forEach(function (s) { if (bal[s.courierId] && ['outbound', 'failed'].indexOf(s.status) >= 0) bal[s.courierId].active++; });
  return Object.keys(bal).map(function (k) { var b = bal[k]; b.balance = b.collected - b.handed; return b; });
}

function merchantDues_() {
  var m = {};
  rows_('Merchants').forEach(function (x) { m[x.id] = { id: x.id, name: x.name, phone: x.phone, cod: 0, fees: 0, due: 0, codes: [] }; });
  rows_('Shipments').forEach(function (s) {
    if (!m[s.merchantId] || truthy_(s.settled)) return;
    if (s.status === 'delivered') { m[s.merchantId].cod += num_(s.cod); m[s.merchantId].fees += num_(s.fee); m[s.merchantId].codes.push(s.code); }
    if (s.status === 'returned') { m[s.merchantId].fees += num_(s.fee) * RETURN_FEE_PCT / 100; m[s.merchantId].codes.push(s.code); }
  });
  return Object.keys(m).map(function (k) { var x = m[k]; x.due = x.cod - x.fees; return x; });
}

function login_(username, password) {
  var u = find_(rows_('Users'), 'username', String(username || '').toLowerCase().trim());
  if (!u || !truthy_(u.active) || hash_(password || '', u.salt) !== u.hash) throw new Error('اسم المستخدم أو الباسورد غلط');
  var token = uid_() + uid_();
  CacheService.getScriptCache().put('t_' + token, u.id, TOKEN_HOURS * 3600);
  var p = pub_(u); p.rates = rows_('Rates');
  return { token: token, user: p };
}

function auth_(token) {
  if (!token) throw new Error('سجّل دخول الأول');
  var id = CacheService.getScriptCache().get('t_' + token);
  if (!id) throw new Error('انتهت الجلسة، سجّل دخول تاني');
  var u = find_(rows_('Users'), 'id', id);
  if (!u || !truthy_(u.active)) throw new Error('الحساب موقوف');
  return u;
}

function can_(u, perm) { var p = ROLES[u.role] || []; return p.indexOf('*') >= 0 || p.indexOf(perm) >= 0; }
function need_(u, perm) { if (!can_(u, perm)) throw new Error('مش مسموحلك بالجزء ده'); }
function pub_(u) { return { id: u.id, username: u.username, name: u.name, role: u.role, phone: u.phone, active: truthy_(u.active) }; }

function newCode_() {
  var d = new Date(), y = String(d.getFullYear()).slice(2), m = ('0' + (d.getMonth() + 1)).slice(-2), dd = ('0' + d.getDate()).slice(-2);
  var codes = {}; rows_('Shipments').forEach(function (s) { codes[s.code] = 1; });
  var c; do { c = 'WS' + y + m + dd + Math.floor(1000 + Math.random() * 9000); } while (codes[c]);
  return c;
}

function log_(code, status, note, by) { append_('History', { ts: now_(), code: code, status: status, note: note || '', by: by || '' }); }

/* ===================== أدوات الشيت ===================== */
/* كاش داخل الطلب الواحد — بيقلل القراية من الشيت ويسرّع الرد */
var SS_ = null, SH_ = {}, RAW_ = {}, ROWS_ = {};
function ss_() { return SS_ || (SS_ = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet()); }
function sheet_(name) {
  if (SH_[name]) return SH_[name];
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('شغّل setup الأول');
  return (SH_[name] = sh);
}
function raw_(name) { return RAW_[name] || (RAW_[name] = sheet_(name).getDataRange().getValues()); }
function dirty_(name) { delete RAW_[name]; delete ROWS_[name]; }
function rows_(name) {
  if (!ROWS_[name]) {
    var v = raw_(name), H = v[0], out = [];
    for (var i = 1; i < v.length; i++) {
      if (v[i].join('') === '') continue;
      var o = {}; for (var j = 0; j < H.length; j++) o[H[j]] = v[i][j] instanceof Date ? v[i][j].toISOString() : v[i][j];
      out.push(o);
    }
    ROWS_[name] = out;
  }
  return ROWS_[name].map(function (o) { var c = {}; for (var k in o) c[k] = o[k]; return c; });
}
var TEXT_COLS = ['receiverPhone', 'senderPhone', 'phone', 'username'];
function cell_(k, v) { return TEXT_COLS.indexOf(k) >= 0 && v !== '' && v != null ? "'" + String(v) : v; }
function append_(name, obj) {
  sheet_(name).appendRow(SHEETS[name].map(function (k) { return obj[k] === undefined ? '' : cell_(k, obj[k]); }));
  dirty_(name);
}
function update_(name, key, val, patch) {
  var sh = sheet_(name), v = raw_(name), H = v[0], ki = H.indexOf(key);
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][ki]) === String(val)) {
      var row = v[i].slice();
      Object.keys(patch).forEach(function (k) { var c = H.indexOf(k); if (c >= 0) row[c] = patch[k]; });
      TEXT_COLS.forEach(function (k) { var c = H.indexOf(k); if (c >= 0 && row[c] !== '' && String(row[c]).charAt(0) !== "'") row[c] = "'" + String(row[c]); });
      sh.getRange(i + 1, 1, 1, H.length).setValues([row]);
      dirty_(name);
      return true;
    }
  }
  return false;
}
function find_(list, key, val) { for (var i = 0; i < list.length; i++) if (String(list[i][key]) === String(val)) return list[i]; return null; }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ok_(data) { return { ok: true, data: data }; }
function now_() { return new Date().toISOString(); }
function day_(v) { if (!v) return ''; var d = new Date(v); if (isNaN(d)) return ''; return Utilities.formatDate(d, 'Africa/Cairo', 'yyyy-MM-dd'); }
function num_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function norm_(c) { return String(c || '').trim().toUpperCase(); }
function truthy_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function hash_(pw, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function mask_(name) { name = String(name || ''); var p = name.split(' '); return p[0] + (p[1] ? ' ' + p[1].charAt(0) + '.' : ''); }
