/* CHÉRIE Care · phần dùng chung cho các trang: menu, bảng giá */
(function () {
  "use strict";
  var C = window.Care, E = C.esc;
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return [].slice.call((r || document).querySelectorAll(s)); }

  /* menu điện thoại và menu thả xuống */
  document.addEventListener("click", function (e) {
    var t = e.target;
    if (t.closest("[data-open-menu]")) { $("#drawer").classList.add("open"); document.body.style.overflow = "hidden"; var f = $("#drawer a"); if (f) f.focus(); return; }
    if (t.closest("[data-close-menu]") || t.id === "drawer") { $("#drawer").classList.remove("open"); document.body.style.overflow = ""; return; }
    var dd = t.closest(".dd>button");
    $$(".dd.open").forEach(function (d) { if (!dd || d !== dd.parentNode) { d.classList.remove("open"); d.querySelector("button").setAttribute("aria-expanded", "false"); } });
    if (dd) { var o = dd.parentNode.classList.toggle("open"); dd.setAttribute("aria-expanded", o); }
  });
  document.addEventListener("keydown", function (e) { if (e.key === "Escape") { $$(".dd.open").forEach(function (d) { d.classList.remove("open"); }); var d = $("#drawer"); if (d) { d.classList.remove("open"); document.body.style.overflow = ""; } } });

  /* bảng giá: hiện ngay bản có sẵn, rồi cập nhật từ Sheet nếu có máy chủ */
  var GIA = C.GIA_MAU.slice(), nghe = [];
  try { var c = JSON.parse(sessionStorage.getItem("care-gia") || "null"); if (c && c.length) GIA = c; } catch (e) {}
  function gia(cb) { nghe.push(cb); cb(GIA); }
  C.call("webInfo").then(function (r) {
    if (r.gia && r.gia.length) { GIA = r.gia; try { sessionStorage.setItem("care-gia", JSON.stringify(GIA)); } catch (e) {} nghe.forEach(function (f) { f(GIA); }); }
  }, function () {});

  function nhom(ds, n) { return ds.filter(function (g) { return g.nhom === n; }); }
  function giaTxt(g, hai) {
    if (!C.num(g.tu) && !C.num(g.den)) return "";
    if (hai && C.num(g.den) > C.num(g.tu)) return "từ " + C.tien(g.tu);
    return C.khoangGia(g.tu, g.den);
  }
  /* vẽ bảng giá cho một hoặc nhiều nhóm; opts.cong: thêm dấu + trước giá (phụ thu) */
  function bang(el, ds, opts) {
    opts = opts || {};
    if (!ds.length) { el.innerHTML = ""; return; }
    var hai = ds.some(function (g) { return C.num(g.cc); });
    el.innerHTML = '<div class="ptable' + (hai ? "" : " one") + '" role="table" aria-label="' + E(opts.ten || "Bảng giá") + '"><div class="pth" role="row"><span role="columnheader">' + E(opts.cot || "Món") + '</span>' +
      (hai ? '<span role="columnheader">Chuẩn</span><span role="columnheader">Signature</span>' : '<span role="columnheader">Giá</span>') + '</div>' +
      ds.map(function (g) {
        var s = [g.tg, g.gc].filter(Boolean).join(" · "), p = giaTxt(g, hai);
        return '<div class="ptr" role="row"><span class="n" role="cell">' + E(g.dv) + (s ? '<span class="s">' + E(s) + '</span>' : '') + '</span><span class="p" role="cell">' + (p ? (opts.cong ? "+" : "") + p : "–") + '</span>' +
          (hai ? '<span class="p" role="cell">' + (C.num(g.cc) ? C.tien(g.cc) : "–") + '</span>' : '') + '</div>';
      }).join("") + '</div>';
  }
  function tuGia(ds, n) { var m = nhom(ds, n).map(function (g) { return C.num(g.tu); }).filter(Boolean); return m.length ? Math.min.apply(null, m) : 0; }

  window.Site = { gia: gia, nhom: nhom, bang: bang, tuGia: tuGia, $: $, $$: $$ };
})();
