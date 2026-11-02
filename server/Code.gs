/* CHÉRIE Care · máy chủ (Google Apps Script)
   Dán vào: Sheet "CHÉRIE Care · Quản lý" → Tiện ích mở rộng → Apps Script (xoá hết code mẫu, dán toàn bộ file này, bấm Lưu).
   Triển khai: Triển khai → Triển khai mới → Loại: Ứng dụng web · Thực thi với tư cách: Tôi · Người có quyền truy cập: Bất kỳ ai.
   Lần đầu Google hỏi quyền (Sheet + Drive để lưu ảnh tình trạng): chọn tài khoản → Nâng cao → Đi tới dự án → Cho phép.
   Dữ liệu và đăng nhập RIÊNG của Care (tab NhanSu), không dùng chung với cherie.io.vn hay Office. */

var SHEET_ID = "14q-LysI0cKpZFfei0DTEiu3R8ygfF5ZdLDJq7XV8FfA";
var TZ = "Asia/Ho_Chi_Minh";
var TOKEN_DAYS = 30;
var ROLES = ["Chủ tiệm", "Quầy", "Kỹ thuật viên", "Thợ may"];
var MON_TT = ["Chờ làm", "Đang làm", "Gửi xưởng", "Chờ kiểm", "Sẵn sàng trả", "Đã trả", "Có sự cố"];
var PHIEU_TT = ["Đang làm", "Sẵn sàng trả", "Đã trả", "Hủy"];
var LICH_TT = ["Mới", "Đã xác nhận", "Đã nhận đồ", "Hủy"];
var LOAI_KHACH = ["Khách ngoài", "Doanh nghiệp", "Thuê đầm", "Office", "Seconde"];
var SITE = "https://care.cherie.io.vn/";

function doGet() { return out({ ok: true, app: "CHÉRIE Care" }); }
function doPost(e) {
  var r;
  try { r = handle(JSON.parse((e && e.postData && e.postData.contents) || "{}")); }
  catch (err) { r = { err: String(err && err.message || err), code: (err && err.code) || "error" }; }
  return out(r);
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fail(msg, code) { var e = new Error(msg); e.code = code || "error"; throw e; }
function ss() { return SpreadsheetApp.openById(SHEET_ID); }
function props() { return PropertiesService.getScriptProperties(); }
function now(f) { return Utilities.formatDate(new Date(), TZ, f || "yyyy-MM-dd HH:mm"); }
function txt(v, n) { v = v == null ? "" : String(v).replace(/[\u0000-\u001f]/g, " ").trim(); return v.slice(0, n || 200); }
function num(v) { var n = Math.round(+String(v == null ? "" : v).replace(/[^\d\-]/g, "") || 0); return isFinite(n) ? n : 0; }
function sdtOk(v) { v = txt(v, 15).replace(/[^\d]/g, ""); if (/^84\d{9}$/.test(v)) v = "0" + v.slice(2); return /^0\d{9}$/.test(v) ? v : ""; }
function pick(v, list, d) { return list.indexOf(v) > -1 ? v : d; }
/* chữ ghi vào ô: giữ số 0 đầu, giữ ngày dạng chữ, chặn công thức */
function cell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string" && (/^0\d+$/.test(v) || /^\d{4}-\d{2}-\d{2}/.test(v) || /^[=+\-@]/.test(v))) return "'" + v;
  return v;
}

function handle(q) {
  switch (q.fn) {
    case "webInfo": return webInfo();
    case "webDatLich": return webDatLich(q);
    case "webTraDon": return webTraDon(q);
    case "staff": return { staff: staffList() };
    case "login": return login(q.nv, q.pin);
    case "logout": props().deleteProperty("tk_" + q.token); return { ok: true };
    case "me": return { me: auth(q.token) };
    case "load": var me = auth(q.token); return { me: me, data: load(me), luc: now() };
    case "taoPhieu": return taoPhieu(auth(q.token), q);
    case "capNhatMon": return capNhatMon(auth(q.token), q);
    case "capNhatPhieu": return capNhatPhieu(auth(q.token), q);
    case "traDo": return traDo(auth(q.token), q);
    case "capNhatLich": return capNhatLich(auth(q.token), q);
    case "taoLich": return taoLich(auth(q.token), q);
    case "upAnh": return upAnh(auth(q.token), q);
    case "luuGia": return luuGia(auth(q.token), q);
    case "luuNhanSu": return luuNhanSu(auth(q.token), q);
  }
  fail("Không rõ yêu cầu", "bad_request");
}

/* ---------- đọc ghi bảng theo tên cột ---------- */
function tab(name) {
  var sh = ss().getSheetByName(name); if (!sh) fail("Thiếu tab " + name, "setup");
  var lc = sh.getLastColumn(), lr = sh.getLastRow();
  var head = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  var rows = [];
  if (lr > 1) sh.getRange(2, 1, lr - 1, lc).getDisplayValues().forEach(function (v, i) {
    var o = { _r: i + 2 }; head.forEach(function (k, j) { if (k) o[k] = v[j]; }); if (v[0]) rows.push(o);
  });
  return { sh: sh, head: head, rows: rows };
}
function find(t, key, val) { for (var i = 0; i < t.rows.length; i++) if (t.rows[i][key] === val) return t.rows[i]; return null; }
function writeObj(t, r, obj) {
  Object.keys(obj).forEach(function (k) { var j = t.head.indexOf(k); if (j < 0 || k.charAt(0) === "_") return; t.sh.getRange(r, j + 1).setValue(cell(obj[k])); });
}
function appendObjs(t, objs) {
  if (!objs.length) return;
  var vals = objs.map(function (o) { return t.head.map(function (k) { return cell(o[k]); }); });
  t.sh.getRange(t.sh.getLastRow() + 1, 1, vals.length, t.head.length).setValues(vals);
}
function log(me, viec, ma, ct) {
  try { var t = tab("NhatKy"); appendObjs(t, [{ Luc: now("yyyy-MM-dd HH:mm:ss"), MaNV: me ? me.id : "web", Ten: me ? me.ten : "Website", Viec: viec, Ma: ma || "", ChiTiet: txt(ct, 300) }]); } catch (e) {}
}
function withLock(f) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return f(); } finally { l.releaseLock(); } }
/* mã tăng dần theo tiền tố, ví dụ PC-0211-01, CA-0211-014 */
function nextMa(rows, key, prefix, width) {
  var mx = 0; rows.forEach(function (r) { var v = String(r[key] || ""); if (v.indexOf(prefix) === 0) { var n = +v.slice(prefix.length); if (n > mx) mx = n; } });
  return function () { mx++; return prefix + ("0000" + mx).slice(-width); };
}

/* ---------- nhân sự & đăng nhập ---------- */
function staffList() {
  return tab("NhanSu").rows.filter(function (n) { return n.Ten && n.Nghi !== "Có"; })
    .map(function (n) { return { id: n.MaNV, ten: n.Ten, vaiTro: n.VaiTro, taoPin: !String(n.PIN || "").trim() && n.VaiTro === "Chủ tiệm" }; });
}
function salt() { var s = props().getProperty("SALT"); if (!s) { s = Utilities.getUuid(); props().setProperty("SALT", s); } return s; }
function hashPin(pin) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt() + "|" + pin).map(function (b) { return ("0" + (b & 255).toString(16)).slice(-2); }).join("");
}
function login(nv, pin) {
  pin = String(pin || "").trim();
  if (!/^\d{4,8}$/.test(pin)) fail("Mã PIN gồm 4–8 chữ số", "bad_pin");
  var cache = CacheService.getScriptCache(), key = "fail_" + nv, n = +(cache.get(key) || 0);
  if (n >= 5) fail("Sai quá 5 lần. Đợi 15 phút rồi thử lại.", "locked");
  return withLock(function () {
    var t = tab("NhanSu"), u = t.rows.filter(function (x) { return x.MaNV === nv && x.Nghi !== "Có"; })[0];
    if (!u) fail("Không tìm thấy nhân sự", "bad_pin");
    var stored = String(u.PIN || "").trim(), ok = false;
    if (!stored && u.VaiTro === "Chủ tiệm") { ok = true; t.sh.getRange(u._r, t.head.indexOf("PIN") + 1).setNumberFormat("@").setValue(hashPin(pin)); }
    else if (/^[0-9a-f]{64}$/.test(stored)) ok = stored === hashPin(pin);
    else if (stored) { ok = stored === pin; if (ok) t.sh.getRange(u._r, t.head.indexOf("PIN") + 1).setNumberFormat("@").setValue(hashPin(pin)); }
    if (!ok) { cache.put(key, String(n + 1), 900); fail("Sai mã PIN", "bad_pin"); }
    cache.remove(key);
    var tk = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    props().setProperty("tk_" + tk, JSON.stringify({ nv: nv, exp: Date.now() + TOKEN_DAYS * 864e5 }));
    log(meOf(u), "Đăng nhập", nv, "");
    return { token: tk, me: meOf(u) };
  });
}
function meOf(u) { return { id: u.MaNV, ten: u.Ten, vaiTro: pick(u.VaiTro, ROLES, "Quầy") }; }
function auth(token) {
  if (!token) fail("Chưa đăng nhập", "revoked");
  var raw = props().getProperty("tk_" + token); if (!raw) fail("Phiên đăng nhập đã hết", "revoked");
  var t = JSON.parse(raw);
  if (t.exp < Date.now()) { props().deleteProperty("tk_" + token); fail("Phiên đăng nhập đã hết", "revoked"); }
  var u = tab("NhanSu").rows.filter(function (x) { return x.MaNV === t.nv && x.Nghi !== "Có"; })[0];
  if (!u) { props().deleteProperty("tk_" + token); fail("Tài khoản đã bị tắt", "revoked"); }
  return meOf(u);
}
function isOwner(me) { return me.vaiTro === "Chủ tiệm"; }
function isDesk(me) { return me.vaiTro === "Chủ tiệm" || me.vaiTro === "Quầy"; }
function need(ok, msg) { if (!ok) fail(msg || "Vai trò của bạn chưa được làm việc này", "forbidden"); }
function laSua(dv) { return /sửa/i.test(String(dv || "")); }

/* ---------- tải dữ liệu theo vai trò ---------- */
function strip(o) { var x = {}; Object.keys(o).forEach(function (k) { if (k.charAt(0) !== "_") x[k] = o[k]; }); return x; }
function load(me) {
  var phieu = tab("Phieu").rows.map(strip), mon = tab("Mon").rows.map(strip), gia = tab("BangGia").rows.map(strip);
  var res = { gia: gia };
  if (isDesk(me)) {
    res.phieu = phieu; res.mon = mon; res.lich = tab("DatLich").rows.map(strip);
    if (isOwner(me)) res.nhanSu = tab("NhanSu").rows.map(function (n) { var x = strip(n); x.coPin = !!String(x.PIN || "").trim(); delete x.PIN; return x; });
    return res;
  }
  /* kỹ thuật viên, thợ may: chỉ thấy món cần làm và mã phiếu, hẹn trả; không thấy số điện thoại, tiền */
  if (me.vaiTro === "Thợ may") mon = mon.filter(function (m) { return laSua(m.DichVu); });
  var can = {}; mon.forEach(function (m) { can[m.MaPhieu] = 1; });
  res.mon = mon.map(function (m) { var x = Object.assign({}, m); if (me.vaiTro !== "Thợ may") { delete x.Gia; } delete x.GiaTriKhai; delete x.DenBu; delete x.GiaXuong; return x; });
  res.phieu = phieu.filter(function (p) { return can[p.MaPhieu]; }).map(function (p) { return { MaPhieu: p.MaPhieu, NgayNhan: p.NgayNhan, HenTra: p.HenTra, TrangThai: p.TrangThai, LoaiKhach: p.LoaiKhach, TenKhach: String(p.TenKhach || "").split(" ").slice(-1)[0], GhiChu: p.GhiChu }; });
  return res;
}

/* ---------- nhận đồ: tạo phiếu và các món ---------- */
function taoPhieu(me, q) {
  need(isDesk(me), "Chỉ quầy hoặc chủ tiệm được nhận đồ");
  var p = q.phieu || {}, ds = Array.isArray(q.mon) ? q.mon.slice(0, 40) : [];
  if (!ds.length) fail("Phiếu chưa có món nào", "bad_request");
  var loai = pick(p.LoaiKhach, LOAI_KHACH, "Khách ngoài");
  var sdt = sdtOk(p.SDT);
  if (!sdt && (loai === "Khách ngoài" || loai === "Doanh nghiệp")) fail("Số Zalo của khách chưa đúng (10 số, bắt đầu bằng 0)", "bad_request");
  return withLock(function () {
    var tp = tab("Phieu"), tm = tab("Mon"), d = now("ddMM");
    var maP = nextMa(tp.rows, "MaPhieu", "PC-" + d + "-", 2)();
    var gen = nextMa(tm.rows, "MaMon", "CA-" + d + "-", 3), luc = now(), tong = 0;
    var mons = ds.map(function (m) {
      var gia = Math.max(0, num(m.Gia)); tong += gia;
      var a = Array.isArray(m.Anh) ? m.Anh : [];
      return { MaMon: gen(), MaPhieu: maP, TenMon: txt(m.TenMon, 80) || "Món", Loai: txt(m.Loai, 40), DichVu: txt(m.DichVu, 80), ChiTiet: txt(m.ChiTiet, 300),
        Gia: gia, GiaTriKhai: Math.max(0, num(m.GiaTriKhai)) || "", TinhTrangNhan: txt(m.TinhTrangNhan, 300),
        Anh1: idOk(a[0]), Anh2: idOk(a[1]), Anh3: idOk(a[2]), Anh4: idOk(a[3]), TrangThai: "Chờ làm", NguoiLam: "", NoiLam: txt(m.NoiLam, 30) || "Tại tiệm", CapNhat: luc };
    });
    var giam = Math.min(tong, Math.max(0, num(p.GiamGia))), phai = tong - giam, thu = Math.min(phai, Math.max(0, num(p.DaThu)));
    var row = { MaPhieu: maP, NgayNhan: luc, SDT: sdt, TenKhach: txt(p.TenKhach, 80), LoaiKhach: loai, CachNhan: txt(p.CachNhan, 30) || "Tại quầy",
      DiaChi: txt(p.DiaChi, 200), HenTra: txt(p.HenTra, 16), SoMon: mons.length, TongTien: tong, GiamGia: giam || "", PhaiThu: phai, DaThu: thu,
      CachThu: thu ? txt(p.CachThu, 30) || "Tiền mặt" : "", TrangThai: "Đang làm", NguoiNhan: me.ten, MaLich: txt(p.MaLich, 20), GhiChu: txt(p.GhiChu, 300), CapNhat: luc };
    appendObjs(tp, [row]); appendObjs(tm, mons);
    if (row.MaLich) { var tl = tab("DatLich"), l = find(tl, "MaLich", row.MaLich); if (l) writeObj(tl, l._r, { TrangThai: "Đã nhận đồ", MaPhieu: maP, XuLy: me.ten }); }
    log(me, "Nhận đồ", maP, mons.length + " món, " + phai + "đ");
    return { ok: true, phieu: row, mon: mons, traCuu: SITE + "?tra=" + maP };
  });
}
function idOk(v) { v = txt(v, 80); return /^[A-Za-z0-9_\-]{20,80}$/.test(v) ? v : ""; }

/* ---------- cập nhật món ở xưởng ---------- */
var F_XUONG = ["TrangThai", "NguoiLam", "NoiLam", "XuongDoiTac", "GhiChuXuong", "SuCo"];
var F_QUAY = F_XUONG.concat(["TenMon", "Loai", "DichVu", "ChiTiet", "Gia", "GiaTriKhai", "TinhTrangNhan", "GiaXuong", "DenBu", "Anh1", "Anh2", "Anh3", "Anh4"]);
function capNhatMon(me, q) {
  return withLock(function () {
    var tm = tab("Mon"), m = find(tm, "MaMon", txt(q.ma, 20)); if (!m) fail("Không tìm thấy món " + q.ma, "not_found");
    if (me.vaiTro === "Thợ may") need(laSua(m.DichVu), "Món này không có việc sửa");
    var allow = isDesk(me) ? F_QUAY : F_XUONG, o = q.obj || {}, w = {};
    Object.keys(o).forEach(function (k) { if (allow.indexOf(k) > -1) w[k] = o[k]; });
    if (w.TrangThai != null) { w.TrangThai = pick(w.TrangThai, MON_TT, m.TrangThai); if (w.TrangThai === "Đã trả") need(isDesk(me), "Chỉ quầy được trả đồ"); }
    ["Gia", "GiaTriKhai", "GiaXuong", "DenBu"].forEach(function (k) { if (w[k] != null) w[k] = Math.max(0, num(w[k])) || ""; });
    ["Anh1", "Anh2", "Anh3", "Anh4"].forEach(function (k) { if (w[k] != null) w[k] = idOk(w[k]); });
    ["TenMon", "Loai", "DichVu", "NguoiLam", "NoiLam", "XuongDoiTac"].forEach(function (k) { if (w[k] != null) w[k] = txt(w[k], 80); });
    ["ChiTiet", "TinhTrangNhan", "GhiChuXuong", "SuCo"].forEach(function (k) { if (w[k] != null) w[k] = txt(w[k], 300); });
    if (w.TrangThai && w.TrangThai !== m.TrangThai && !w.NguoiLam && /Đang làm|Gửi xưởng/.test(w.TrangThai)) w.NguoiLam = me.ten;
    w.CapNhat = now();
    writeObj(tm, m._r, w); Object.keys(w).forEach(function (k) { m[k] = w[k]; });
    var p = dongBoPhieu(m.MaPhieu, w.Gia != null);
    log(me, "Cập nhật món", m.MaMon, JSON.stringify(w).slice(0, 280));
    return { ok: true, mon: strip(m), phieu: p };
  });
}
/* tính lại tổng tiền và trạng thái phiếu theo các món */
function dongBoPhieu(maP, tinhTien) {
  var tp = tab("Phieu"), p = find(tp, "MaPhieu", maP); if (!p) return null;
  var ms = tab("Mon").rows.filter(function (x) { return x.MaPhieu === maP; }), w = {};
  if (tinhTien) { var tong = 0; ms.forEach(function (x) { tong += num(x.Gia); }); w.TongTien = tong; w.PhaiThu = tong - Math.min(tong, num(p.GiamGia)); }
  if (p.TrangThai !== "Đã trả" && p.TrangThai !== "Hủy") {
    var con = ms.filter(function (x) { return x.TrangThai !== "Đã trả"; });
    var tt = con.length && con.every(function (x) { return x.TrangThai === "Sẵn sàng trả"; }) ? "Sẵn sàng trả" : "Đang làm";
    if (tt !== p.TrangThai) w.TrangThai = tt;
  }
  if (Object.keys(w).length) { w.CapNhat = now(); writeObj(tp, p._r, w); Object.keys(w).forEach(function (k) { p[k] = w[k]; }); }
  return strip(p);
}

/* ---------- phiếu: sửa thông tin, thu tiền, trả đồ ---------- */
function capNhatPhieu(me, q) {
  need(isDesk(me));
  return withLock(function () {
    var tp = tab("Phieu"), p = find(tp, "MaPhieu", txt(q.ma, 20)); if (!p) fail("Không tìm thấy phiếu " + q.ma, "not_found");
    var o = q.obj || {}, w = {};
    if (o.SDT != null) { var s = sdtOk(o.SDT); if (o.SDT && !s) fail("Số Zalo chưa đúng", "bad_request"); w.SDT = s; }
    ["TenKhach", "CachNhan", "CachThu"].forEach(function (k) { if (o[k] != null) w[k] = txt(o[k], 80); });
    ["DiaChi", "GhiChu"].forEach(function (k) { if (o[k] != null) w[k] = txt(o[k], 300); });
    if (o.HenTra != null) w.HenTra = txt(o.HenTra, 16);
    if (o.LoaiKhach != null) w.LoaiKhach = pick(o.LoaiKhach, LOAI_KHACH, p.LoaiKhach);
    if (o.GiamGia != null) { var tong = num(p.TongTien); w.GiamGia = Math.min(tong, Math.max(0, num(o.GiamGia))) || ""; w.PhaiThu = tong - num(w.GiamGia); }
    if (o.ThuThem != null) { var phai = w.PhaiThu != null ? w.PhaiThu : num(p.PhaiThu); w.DaThu = Math.min(phai, num(p.DaThu) + Math.max(0, num(o.ThuThem))); if (num(o.ThuThem) && !w.CachThu) w.CachThu = txt(o.CachThu, 30) || "Tiền mặt"; }
    if (o.TrangThai === "Hủy") { need(isOwner(me), "Chỉ chủ tiệm được huỷ phiếu"); w.TrangThai = "Hủy"; }
    w.CapNhat = now();
    writeObj(tp, p._r, w); Object.keys(w).forEach(function (k) { p[k] = w[k]; });
    log(me, "Sửa phiếu", p.MaPhieu, JSON.stringify(w).slice(0, 280));
    return { ok: true, phieu: strip(p) };
  });
}
function traDo(me, q) {
  need(isDesk(me), "Chỉ quầy được trả đồ");
  return withLock(function () {
    var tp = tab("Phieu"), p = find(tp, "MaPhieu", txt(q.ma, 20)); if (!p) fail("Không tìm thấy phiếu " + q.ma, "not_found");
    if (p.TrangThai === "Đã trả") fail("Phiếu này đã trả rồi", "bad_request");
    var tm = tab("Mon"), luc = now(), chon = Array.isArray(q.mon) ? q.mon : null;
    var cua = tm.rows.filter(function (m) { return m.MaPhieu === p.MaPhieu && m.TrangThai !== "Đã trả"; });
    var tra = cua.filter(function (m) { return !chon || chon.indexOf(m.MaMon) > -1; }), n = tra.length;
    if (!n) fail("Chưa chọn món nào để trả", "bad_request");
    var con = cua.length - n;
    var phai = num(p.PhaiThu), thu = Math.min(phai, num(p.DaThu) + Math.max(0, num(q.thuThem)));
    /* kiểm tiền trước khi ghi, để báo lỗi không làm lệch dữ liệu */
    if (!con && thu < phai && !q.choNo) fail("Khách còn thiếu " + (phai - thu) + "đ. Thu đủ rồi mới trả hết đồ.", "con_no");
    tra.forEach(function (m) { writeObj(tm, m._r, { TrangThai: "Đã trả", CapNhat: luc }); m.TrangThai = "Đã trả"; });
    var w = { DaThu: thu, CapNhat: luc };
    if (num(q.thuThem)) w.CachThu = txt(q.cachThu, 30) || "Tiền mặt";
    if (!con) { w.TrangThai = "Đã trả"; w.NgayTra = luc; w.NguoiTra = me.ten; }
    writeObj(tp, p._r, w); Object.keys(w).forEach(function (k) { p[k] = w[k]; });
    log(me, "Trả đồ", p.MaPhieu, n + " món, thu thêm " + num(q.thuThem) + "đ");
    return { ok: true, phieu: strip(p), mon: tm.rows.filter(function (m) { return m.MaPhieu === p.MaPhieu; }).map(strip) };
  });
}

/* ---------- lịch hẹn nhận đồ ---------- */
function capNhatLich(me, q) {
  need(isDesk(me));
  return withLock(function () {
    var tl = tab("DatLich"), l = find(tl, "MaLich", txt(q.ma, 20)); if (!l) fail("Không tìm thấy lịch " + q.ma, "not_found");
    var o = q.obj || {}, w = {};
    if (o.TrangThai != null) w.TrangThai = pick(o.TrangThai, LICH_TT, l.TrangThai);
    if (o.NgayHen != null) w.NgayHen = txt(o.NgayHen, 10);
    if (o.KhungGio != null) w.KhungGio = txt(o.KhungGio, 30);
    if (o.GhiChu != null) w.GhiChu = txt(o.GhiChu, 500);
    w.XuLy = me.ten;
    writeObj(tl, l._r, w); Object.keys(w).forEach(function (k) { l[k] = w[k]; });
    log(me, "Lịch hẹn", l.MaLich, JSON.stringify(w));
    return { ok: true, lich: strip(l) };
  });
}
function taoLich(me, q) { need(isDesk(me)); q.nguon = "Quầy · " + me.ten; return datLich(q, me); }

/* ---------- ảnh tình trạng: lưu vào Drive, ai có link mới xem được ---------- */
function folder() {
  var id = props().getProperty("ANH_FOLDER");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = DriveApp.getFileById(SHEET_ID).getParents(), f = (parent.hasNext() ? parent.next() : DriveApp.getRootFolder()).createFolder("CHÉRIE Care · Ảnh tình trạng");
  props().setProperty("ANH_FOLDER", f.getId()); return f;
}
function upAnh(me, q) {
  var b64 = String(q.data || "").replace(/^data:image\/\w+;base64,/, "");
  if (!b64 || b64.length > 2800000) fail("Ảnh trống hoặc quá lớn", "bad_request");
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/jpeg", txt(q.ten, 60).replace(/[^\w\-\.]/g, "_") || ("anh-" + now("yyyyMMdd-HHmmss") + ".jpg"));
  var f = folder().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: f.getId() };
}

/* ---------- bảng giá, nhân sự (chủ tiệm) ---------- */
function luuGia(me, q) {
  need(isOwner(me), "Chỉ chủ tiệm được sửa bảng giá");
  return withLock(function () {
    var t = tab("BangGia"), o = q.obj || {}, ma = txt(o.MaGia, 10), r = find(t, "MaGia", ma);
    var w = { Nhom: txt(o.Nhom, 40), DichVu: txt(o.DichVu, 120), GiaTu: num(o.GiaTu), GiaDen: num(o.GiaDen) || num(o.GiaTu), GiaCaoCap: num(o.GiaCaoCap) || "", DonVi: txt(o.DonVi, 20) || "món",
      SoNgay: Math.max(0, num(o.SoNgay)), ThoiGianLam: txt(o.ThoiGianLam, 60), GhiChu: txt(o.GhiChu, 200), HienWeb: o.HienWeb === "Không" ? "Không" : "Có" };
    if (q.xoa && r) { t.sh.deleteRow(r._r); log(me, "Xoá giá", ma, ""); return { ok: true }; }
    if (r) writeObj(t, r._r, w);
    else { w.MaGia = nextMa(t.rows, "MaGia", "X", 2)(); w.NguonGia = "Chủ tiệm nhập " + now("dd/MM/yyyy"); appendObjs(t, [w]); }
    log(me, "Bảng giá", w.MaGia || ma, w.DichVu + " " + w.GiaTu + "–" + w.GiaDen);
    return { ok: true, gia: tab("BangGia").rows.map(strip) };
  });
}
function luuNhanSu(me, q) {
  need(isOwner(me), "Chỉ chủ tiệm được sửa nhân sự");
  return withLock(function () {
    var t = tab("NhanSu"), o = q.obj || {}, r = o.MaNV ? find(t, "MaNV", txt(o.MaNV, 10)) : null;
    var w = { Ten: txt(o.Ten, 60), VaiTro: pick(o.VaiTro, ROLES, "Quầy"), SDT: sdtOk(o.SDT), Nghi: o.Nghi === "Có" ? "Có" : "" };
    if (!w.Ten) fail("Nhập tên nhân sự", "bad_request");
    if (r && r.MaNV === me.id && (w.VaiTro !== "Chủ tiệm" || w.Nghi)) fail("Không tự đổi vai trò hoặc tự cho mình nghỉ", "bad_request");
    if (o.PIN) { if (!/^\d{4,8}$/.test(String(o.PIN))) fail("PIN gồm 4–8 chữ số", "bad_pin"); w.PIN = hashPin(String(o.PIN)); }
    if (r) writeObj(t, r._r, w);
    else { w.MaNV = nextMa(t.rows, "MaNV", "NV", 2)(); if (!w.PIN) fail("Nhân sự mới cần mã PIN", "bad_pin"); appendObjs(t, [w]); }
    log(me, "Nhân sự", w.MaNV || o.MaNV, w.Ten + " · " + w.VaiTro + (w.Nghi ? " · nghỉ" : ""));
    return { ok: true, nhanSu: tab("NhanSu").rows.map(function (n) { var x = strip(n); x.coPin = !!String(x.PIN || "").trim(); delete x.PIN; return x; }) };
  });
}

/* ---------- công khai: bảng giá, đặt lịch, tra đơn (không cần đăng nhập) ---------- */
function rate(key, max, sec) {
  var c = CacheService.getScriptCache(), n = +(c.get(key) || 0), n0 = +(c.get("web_all") || 0);
  if (n >= max) fail("Bạn gửi hơi nhiều lần. Nhắn Zalo 0377 159 189 giúp tụi mình nhé.", "rate");
  if (n0 >= 600) fail("Hệ thống đang bận, thử lại sau ít phút.", "rate");
  c.put(key, String(n + 1), sec); c.put("web_all", String(n0 + 1), 3600);
}
function webInfo() {
  var c = CacheService.getScriptCache(), hit = c.get("web_gia");
  if (hit) return JSON.parse(hit);
  var gia = tab("BangGia").rows.filter(function (g) { return g.HienWeb !== "Không"; })
    .map(function (g) { return { ma: g.MaGia, nhom: g.Nhom, dv: g.DichVu, tu: num(g.GiaTu), den: num(g.GiaDen), cc: num(g.GiaCaoCap), dvi: g.DonVi, tg: g.ThoiGianLam, gc: g.GhiChu }; });
  var r = { ok: true, gia: gia }; c.put("web_gia", JSON.stringify(r), 300); return r;
}
function webDatLich(q) {
  var sdt = sdtOk(q.sdt); if (!sdt) fail("Số Zalo chưa đúng (10 số, bắt đầu bằng 0)", "bad_request");
  if (q.web2) return { ok: true, ma: "DL-0000" }; /* ô bẫy chống máy gửi tự động */
  rate("dl_" + sdt, 5, 21600);
  return datLich(q, null);
}
function datLich(q, me) {
  var sdt = sdtOk(q.sdt); if (!sdt) fail("Số Zalo chưa đúng (10 số, bắt đầu bằng 0)", "bad_request");
  var ten = txt(q.ten, 80); if (!ten) fail("Nhập tên để tụi mình gọi cho đúng", "bad_request");
  var ngay = txt(q.ngayHen, 10); if (ngay && !/^\d{4}-\d{2}-\d{2}$/.test(ngay)) ngay = "";
  return withLock(function () {
    var t = tab("DatLich"), ma = nextMa(t.rows, "MaLich", "DL-", 4)();
    var row = { MaLich: ma, ThoiGian: now(), Ten: ten, SDT: sdt, DichVu: txt(Array.isArray(q.dichVu) ? q.dichVu.join(", ") : q.dichVu, 400), SoMon: Math.min(99, Math.max(1, num(q.soMon) || 1)),
      CachNhan: q.cachNhan === "Giao nhận tận nơi" ? "Giao nhận tận nơi" : "Mang đến quầy", NgayHen: ngay, KhungGio: txt(q.khungGio, 30), DiaChi: txt(q.diaChi, 200),
      GhiChu: txt(q.ghiChu, 500), TrangThai: "Mới", XuLy: "", MaPhieu: "", Nguon: txt(q.nguon, 60) || "Website",
      Hang: q.hang === "Signature" ? "Signature" : "Chuẩn", TamTinh: Math.max(0, num(q.tamTinh)) || "" };
    appendObjs(t, [row]);
    log(me, "Đặt lịch", ma, row.DichVu + " · " + row.NgayHen + " " + row.KhungGio);
    return { ok: true, ma: ma, lich: row };
  });
}
/* tra đơn: cần mã phiếu (hoặc mã món) và 4 số cuối Zalo. Không trả tên, số điện thoại, địa chỉ. */
function webTraDon(q) {
  var ma = txt(q.ma, 20).toUpperCase().replace(/\s/g, ""), so = txt(q.so, 4).replace(/\D/g, "");
  if (!/^(PC|CA)-\d{4}-\d{2,4}$/.test(ma) || so.length !== 4) fail("Nhập mã phiếu (ví dụ PC-0211-01) và 4 số cuối Zalo", "bad_request");
  rate("tra_" + ma, 30, 3600);
  var tp = tab("Phieu"), tm = tab("Mon"), maP = ma;
  if (ma.indexOf("CA-") === 0) { var m0 = find(tm, "MaMon", ma); maP = m0 ? m0.MaPhieu : ""; }
  var p = maP ? find(tp, "MaPhieu", maP) : null;
  if (!p || String(p.SDT || "").slice(-4) !== so) fail("Không tìm thấy phiếu. Kiểm tra lại mã và 4 số cuối Zalo.", "not_found");
  var mons = tm.rows.filter(function (m) { return m.MaPhieu === p.MaPhieu; }).map(function (m) { return { ma: m.MaMon, ten: m.TenMon, dv: m.DichVu, tt: m.TrangThai === "Có sự cố" ? "Đang làm" : m.TrangThai }; });
  var phai = num(p.PhaiThu), thu = num(p.DaThu);
  return { ok: true, phieu: { ma: p.MaPhieu, ngayNhan: p.NgayNhan, henTra: p.HenTra, tt: p.TrangThai, ngayTra: p.NgayTra, phai: phai, thu: thu, con: Math.max(0, phai - thu), ten: String(p.TenKhach || "").split(" ").slice(-1)[0] }, mon: mons };
}

/* Chạy tay trong trình chỉnh sửa nếu muốn đăng xuất mọi người */
function dangXuatTatCa() { var p = props(), all = p.getProperties(); Object.keys(all).forEach(function (k) { if (k.indexOf("tk_") === 0) p.deleteProperty(k); }); }
