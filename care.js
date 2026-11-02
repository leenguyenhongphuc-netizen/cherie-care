/* CHÉRIE Care · hàm dùng chung cho trang khách và app quản lý */
(function () {
  "use strict";
  var KEY = "care-token", ME = "care-me";
  var token = "", me = null;
  try { token = localStorage.getItem(KEY) || ""; me = JSON.parse(localStorage.getItem(ME) || "null"); } catch (e) {}

  /* gọi máy chủ: POST thân JSON dạng text/plain (không cần hỏi trước CORS), tự thử lại khi mạng rớt */
  function call(fn, args) {
    return new Promise(function (ok, no) {
      var api = window.CARE_API;
      if (!api) { var x = new Error("App chưa kết nối máy chủ. Chủ tiệm cần triển khai Apps Script và dán link vào care/config.js."); x.code = "no_api"; return no(x); }
      var body = JSON.stringify(Object.assign({ fn: fn, token: token }, args || {})), lan = 0, cho = [700, 1600, 3200], done = false;
      var tm = setTimeout(function () { if (done) return; done = true; var x = new Error("Máy chủ Google phản hồi quá 45 giây. Bấm thử lại."); x.code = "unavailable"; no(x); }, 45000);
      var thu = function () {
        fetch(api, { method: "POST", redirect: "follow", body: body }).then(function (r) { return r.json(); }).then(function (r) {
          if (done) return; done = true; clearTimeout(tm);
          if (r && r.err) { var e = new Error(r.err); e.code = r.code || "error"; if (e.code === "revoked") saveAuth("", null); no(e); } else ok(r);
        }, function () {
          if (done) return;
          if (lan < cho.length) { setTimeout(thu, cho[lan++]); return; }
          done = true; clearTimeout(tm); var x = new Error("Không kết nối được máy chủ. Kiểm tra mạng rồi thử lại."); x.code = "unavailable"; no(x);
        });
      };
      thu();
    });
  }
  function saveAuth(t, m) {
    token = t || ""; me = m || null;
    try { if (token) { localStorage.setItem(KEY, token); localStorage.setItem(ME, JSON.stringify(me)); } else { localStorage.removeItem(KEY); localStorage.removeItem(ME); } } catch (e) {}
  }

  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function num(v) { var n = Math.round(+String(v == null ? "" : v).replace(/[^\d\-]/g, "") || 0); return isFinite(n) ? n : 0; }
  function tien(v) { return num(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".") + "đ"; }
  function khoangGia(tu, den) { tu = num(tu); den = num(den) || tu; return tu === den ? tien(tu) : tien(tu) + " – " + tien(den); }
  function pad(n) { return ("0" + n).slice(-2); }
  /* giờ Việt Nam, không phụ thuộc múi giờ của máy */
  function vnNow() { var d = new Date(Date.now() + 7 * 36e5); return d; }
  function isoVN(d) { d = d || vnNow(); return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate()) + " " + pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes()); }
  function homNay() { return isoVN().slice(0, 10); }
  function congNgay(iso, n) { var d = new Date(iso.slice(0, 10) + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
  function ngayDep(iso) {
    if (!iso) return "";
    var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/); if (!m) return String(iso);
    var thu = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"][new Date(m[1] + "-" + m[2] + "-" + m[3] + "T00:00:00Z").getUTCDay()];
    return thu + " " + m[3] + "/" + m[2] + (m[4] ? " · " + m[4] + ":" + m[5] : "");
  }
  function sdtOk(v) { v = String(v || "").replace(/[^\d]/g, ""); if (/^84\d{9}$/.test(v)) v = "0" + v.slice(2); return /^0\d{9}$/.test(v) ? v : ""; }
  function sdtDep(v) { v = String(v || ""); return v.length === 10 ? v.slice(0, 4) + " " + v.slice(4, 7) + " " + v.slice(7) : v; }
  function zalo(sdt) { return "https://zalo.me/" + String(sdt || "").replace(/^0/, "84"); }
  function anh(id, w) { return id ? "https://drive.google.com/thumbnail?id=" + encodeURIComponent(id) + "&sz=w" + (w || 400) : ""; }

  /* nén ảnh trước khi gửi: cạnh dài tối đa 1280px, JPEG 0,8 */
  function nenAnh(file, max, q) {
    max = max || 1280; q = q || 0.8;
    return new Promise(function (ok, no) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var s = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight)), c = document.createElement("canvas");
        c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
        ok(c.toDataURL("image/jpeg", q));
      };
      img.onerror = function () { URL.revokeObjectURL(url); no(new Error("Không đọc được ảnh")); };
      img.src = url;
    });
  }
  function copy(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    return new Promise(function (ok) { var t = document.createElement("textarea"); t.value = text; t.style.cssText = "position:fixed;opacity:0"; document.body.appendChild(t); t.select(); try { document.execCommand("copy"); } catch (e) {} t.remove(); ok(); });
  }
  var toastT = 0;
  function toast(msg, loi) {
    var n = document.getElementById("toast");
    if (!n) { n = document.createElement("div"); n.id = "toast"; n.setAttribute("role", "status"); document.body.appendChild(n); }
    n.textContent = msg; n.className = "toast show" + (loi ? " err" : "");
    clearTimeout(toastT); toastT = setTimeout(function () { n.className = "toast"; }, loi ? 5000 : 2600);
  }

  /* bảng giá mặc định khi chưa tải được từ máy chủ (giá tham khảo, khớp tab BangGia) */
  var GIA_MAU = [
    { ma: "G01", nhom: "Hấp lấy liền", dv: "Hấp một món (đồ sạch, chỉ nhăn)", tu: 40000, den: 60000, cc: 0, dvi: "món", tg: "Lấy trong 60 phút" },
    { ma: "G02", nhom: "Giặt và hấp", dv: "Áo, chân váy, quần", tu: 40000, den: 40000, cc: 90000, dvi: "món", tg: "2 ngày" },
    { ma: "G03", nhom: "Giặt và hấp", dv: "Đầm ngắn, đầm công sở", tu: 90000, den: 90000, cc: 150000, dvi: "món", tg: "2 ngày" },
    { ma: "G04", nhom: "Giặt và hấp", dv: "Đầm dài, đầm tiệc", tu: 100000, den: 100000, cc: 220000, dvi: "món", tg: "3 ngày" },
    { ma: "G05", nhom: "Giặt và hấp", dv: "Đầm dạ hội", tu: 150000, den: 200000, cc: 320000, dvi: "món", tg: "3 ngày" },
    { ma: "G06", nhom: "Giặt và hấp", dv: "Áo dài", tu: 90000, den: 90000, cc: 180000, dvi: "bộ", tg: "2 ngày" },
    { ma: "G07", nhom: "Giặt và hấp", dv: "Vest, blazer", tu: 90000, den: 90000, cc: 180000, dvi: "món", tg: "2 ngày" },
    { ma: "G08", nhom: "Giặt và hấp", dv: "Áo khoác dạ, măng tô", tu: 150000, den: 150000, cc: 250000, dvi: "món", tg: "3 ngày" },
    { ma: "W01", nhom: "Váy cưới", dv: "Đầm cưới ngắn, đầm tiệc cưới", tu: 300000, den: 300000, cc: 0, dvi: "chiếc", tg: "5 ngày" },
    { ma: "W02", nhom: "Váy cưới", dv: "Dáng A, đuôi cá · vải tổng hợp", tu: 350000, den: 350000, cc: 0, dvi: "chiếc", tg: "5 ngày" },
    { ma: "W03", nhom: "Váy cưới", dv: "Dáng A, đuôi cá · lụa, satin", tu: 400000, den: 400000, cc: 0, dvi: "chiếc", tg: "7 ngày" },
    { ma: "W04", nhom: "Váy cưới", dv: "Dáng công chúa, xoè phồng", tu: 400000, den: 400000, cc: 0, dvi: "chiếc", tg: "7 ngày" },
    { ma: "W05", nhom: "Váy cưới", dv: "Dáng công chúa · lụa, satin", tu: 450000, den: 450000, cc: 0, dvi: "chiếc", tg: "7 ngày" },
    { ma: "W11", nhom: "Phụ thu váy cưới", dv: "Đuôi váy dưới 1m", tu: 50000, den: 50000, cc: 0, dvi: "", tg: "" },
    { ma: "W12", nhom: "Phụ thu váy cưới", dv: "Đuôi váy 1–2m", tu: 120000, den: 120000, cc: 0, dvi: "", tg: "" },
    { ma: "W13", nhom: "Phụ thu váy cưới", dv: "Đuôi váy trên 2m", tu: 200000, den: 200000, cc: 0, dvi: "", tg: "" },
    { ma: "W14", nhom: "Phụ thu váy cưới", dv: "Đính kết dày (cườm, pha lê, ren nổi)", tu: 100000, den: 100000, cc: 0, dvi: "", tg: "" },
    { ma: "W15", nhom: "Phụ thu váy cưới", dv: "Voan cưới", tu: 50000, den: 50000, cc: 0, dvi: "", tg: "" },
    { ma: "W21", nhom: "Phụ thu váy cưới", dv: "Gói cất giữ: hút chân không, hộp cứng, túi hút ẩm", tu: 350000, den: 350000, cc: 0, dvi: "", tg: "+3 ngày" },
    { ma: "S01", nhom: "Sửa đồ", dv: "Lên lai", tu: 25000, den: 40000, cc: 0, dvi: "món", tg: "1 ngày" },
    { ma: "S02", nhom: "Sửa đồ", dv: "Bóp eo", tu: 40000, den: 80000, cc: 0, dvi: "món", tg: "2 ngày" },
    { ma: "S03", nhom: "Sửa đồ", dv: "Thu nhỏ đầm", tu: 80000, den: 200000, cc: 0, dvi: "món", tg: "3 ngày" },
    { ma: "S04", nhom: "Sửa đồ", dv: "Thay khoá kéo", tu: 50000, den: 80000, cc: 0, dvi: "món", tg: "1 ngày" },
    { ma: "C01", nhom: "Chăm thêm", dv: "Tẩy vết khó (rượu vang, son, dầu)", tu: 30000, den: 80000, cc: 0, dvi: "vết", tg: "" },
    { ma: "C02", nhom: "Chăm thêm", dv: "Chống mốc, khử khuẩn", tu: 30000, den: 30000, cc: 0, dvi: "món", tg: "" },
    { ma: "C03", nhom: "Chăm thêm", dv: "Túi vải bọc đồ", tu: 40000, den: 40000, cc: 0, dvi: "túi", tg: "" },
    { ma: "C04", nhom: "Chăm thêm", dv: "Làm gấp trong ngày", tu: 0, den: 0, cc: 0, dvi: "", tg: "", gc: "Cộng 50% giá dịch vụ" },
    { ma: "C05", nhom: "Chăm thêm", dv: "Khử mùi khói, mồ hôi, nước hoa", tu: 30000, den: 30000, cc: 0, dvi: "món", tg: "" },
    { ma: "C06", nhom: "Chăm thêm", dv: "Chống thấm, chống bám bẩn (áo khoác, đầm sáng màu)", tu: 50000, den: 50000, cc: 0, dvi: "món", tg: "+1 ngày" },
    { ma: "C07", nhom: "Chăm thêm", dv: "Làm trắng đồ trắng ngả vàng", tu: 50000, den: 100000, cc: 0, dvi: "món", tg: "+1 ngày" },
    { ma: "C08", nhom: "Chăm thêm", dv: "Đính lại cườm, sequin, nút trang trí rơi", tu: 30000, den: 100000, cc: 0, dvi: "món", tg: "+1 ngày", gc: "Nút thường lỏng: đính lại miễn phí" },
    { ma: "C09", nhom: "Chăm thêm", dv: "Gấp vali chống nhăn (giấy lụa, túi mỏng)", tu: 20000, den: 20000, cc: 0, dvi: "món", tg: "", gc: "Cho khách mang đồ đi xa" },
    { ma: "C10", nhom: "Chăm thêm", dv: "Móc gỗ giữ vai", tu: 30000, den: 30000, cc: 0, dvi: "cái", tg: "", gc: "Đã gồm trong hạng Signature" },
    { ma: "T01", nhom: "Dịch vụ tận nơi", dv: "Hấp tại khách sạn, nhà riêng (nội thành)", tu: 150000, den: 150000, cc: 0, dvi: "lượt", tg: "Hẹn trước 3 giờ", gc: "Cộng giá hấp từng món" },
    { ma: "T02", nhom: "Dịch vụ tận nơi", dv: "Sẵn sàng đi tiệc: hấp, sửa nhanh lai hoặc eo, giao trước giờ hẹn", tu: 250000, den: 250000, cc: 0, dvi: "bộ", tg: "Trong ngày" },
    { ma: "T03", nhom: "Dịch vụ tận nơi", dv: "Hấp, chỉnh váy áo tại tiệc cưới, buổi chụp ảnh", tu: 2500000, den: 2500000, cc: 0, dvi: "sự kiện", tg: "4–6 giờ", gc: "1 kỹ thuật viên + máy hấp di động" },
    { ma: "K01", nhom: "Gói", dv: "Gói Tết 5 món", tu: 390000, den: 390000, cc: 0, dvi: "gói", tg: "", gc: "Giặt hấp, chống mốc. Nhận đến 20 tháng Chạp" },
    { ma: "K02", nhom: "Gói", dv: "Gói Tết 10 món", tu: 720000, den: 720000, cc: 0, dvi: "gói", tg: "", gc: "Giặt hấp, chống mốc. Nhận đến 20 tháng Chạp" },
    { ma: "K03", nhom: "Gói", dv: "Tủ Chérie: giặt và giữ 5 món qua mùa nồm", tu: 690000, den: 690000, cc: 0, dvi: "gói", tg: "", gc: "Giữ tối đa 6 tháng trong phòng hút ẩm" },
    { ma: "K04", nhom: "Gói", dv: "Gói sau cưới: váy cưới dáng A, áo dài ăn hỏi, vest chú rể, cất giữ váy", tu: 790000, den: 790000, cc: 0, dvi: "gói", tg: "7 ngày", gc: "Mua lẻ 880.000đ" },
    { ma: "K05", nhom: "Gói", dv: "Thẻ Care trả trước: nạp 1.000.000đ dùng 1.100.000đ", tu: 1000000, den: 1000000, cc: 0, dvi: "thẻ", tg: "", gc: "Dùng trong 12 tháng, cho mọi dịch vụ" },
    { ma: "K06", nhom: "Gói", dv: "Thẻ Care trả trước: nạp 2.000.000đ dùng 2.300.000đ", tu: 2000000, den: 2000000, cc: 0, dvi: "thẻ", tg: "", gc: "Dùng trong 12 tháng, cho mọi dịch vụ" }
  ];

  window.Care = {
    call: call, saveAuth: saveAuth, token: function () { return token; }, me: function () { return me; },
    esc: esc, num: num, tien: tien, khoangGia: khoangGia, isoVN: isoVN, homNay: homNay, congNgay: congNgay, ngayDep: ngayDep,
    sdtOk: sdtOk, sdtDep: sdtDep, zalo: zalo, anh: anh, nenAnh: nenAnh, copy: copy, toast: toast, GIA_MAU: GIA_MAU,
    ZALO_TIEM: "0377159189", SITE: "https://care.cherie.io.vn/tra-cuu/"
  };
})();
