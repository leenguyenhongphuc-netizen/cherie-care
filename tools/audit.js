/* Kiểm tra toàn bộ website CHÉRIE Care trước khi giao (theo skill ui-ux-pro-max).
   Chạy: node tools/audit.js [thư mục ảnh]   (cần playwright; máy chủ được giả lập, không ghi dữ liệu thật)
   Kiểm: lỗi JS, tràn ngang, nút bấm dưới 44px, vị trí nút chính đầu tiên, luồng đặt lịch và tra cứu, ảnh chụp sáng/tối. */
const path = require("path"), fs = require("fs"), http = require("http");
const PW = process.env.PW || "playwright";
const { chromium } = require(PW);
const ROOT = path.resolve(__dirname, ".."), OUT = path.resolve(process.argv[2] || path.join(ROOT, "tools", "shots"));
fs.mkdirSync(OUT, { recursive: true });
const PAGES = ["/", "/dich-vu/giat-hap/", "/dich-vu/vay-cuoi/", "/dich-vu/sua-do/", "/dich-vu/hap-tan-noi/", "/bang-gia/", "/goi/", "/cam-ket/", "/doanh-nghiep/", "/dat-lich/", "/tra-cuu/"];
const API = "https://script.google.com/macros/s/AUDIT/exec";
const TYPES = { ".html": "text/html", ".css": "text/css", ".js": "application/javascript", ".txt": "text/plain" };
const srv = http.createServer((q, r) => { let p = decodeURIComponent(q.url.split("?")[0]); if (p.endsWith("/")) p += "index.html"; const f = path.join(ROOT, p);
  fs.readFile(f, (e, b) => { if (e) { r.writeHead(404); return r.end(); } r.writeHead(200, { "Content-Type": TYPES[path.extname(f)] || "application/octet-stream" }); r.end(b); }); });
function stub(q) {
  if (q.fn === "webDatLich") return { ok: true, ma: "DL-0001" };
  if (q.fn === "webTraDon") return q.so === "3456" ? { ok: true, phieu: { ma: "PC-0211-01", ngayNhan: "2026-11-02 09:08", henTra: "2026-11-05 17:00", tt: "Đang làm", phai: 190000, thu: 100000, con: 90000, ten: "Lan" }, mon: [{ ma: "CA-0211-001", ten: "Đầm lụa đỏ", dv: "Giặt hấp · Đầm dài", tt: "Gửi xưởng" }, { ma: "CA-0211-002", ten: "Quần tây", dv: "Sửa đồ · Lên lai", tt: "Sẵn sàng trả" }] } : { err: "Không tìm thấy phiếu.", code: "not_found" };
  return { err: "stub", code: "stub" };
}
const res = [];
async function checks(page, url, tag) {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, over = document.documentElement.scrollWidth - W;
    const small = [...document.querySelectorAll("a,button,input,select,textarea,summary,[role=button]")].filter(e => { const r = e.getBoundingClientRect(), s = getComputedStyle(e);
      if (!r.width || s.visibility === "hidden" || e.closest(".drawer:not(.open)") || e.closest("footer") || e.closest(".crumb") || e.closest(".sr") || e.classList.contains("sr")) return false;
      if (e.tagName === "A" && e.closest("p")) return false; /* liên kết trong câu văn */
      if (e.type === "radio" || e.type === "checkbox") return false; /* bấm cả nhãn */
      if (e.getAttribute("aria-hidden") === "true" || e.tabIndex < 0) return false; /* ô bẫy chống máy */
      return r.height < 44 || r.width < 44; }).map(e => (e.textContent || e.getAttribute("aria-label") || e.name || e.tagName).trim().slice(0, 30));
    const ys = [...document.querySelectorAll("main .btn:not(.ghost):not(.quiet), .sbar .btn")].filter(e => e.getBoundingClientRect().width).map(e => { const r = e.getBoundingClientRect(); return Math.round(getComputedStyle(e.closest(".sbar") || e).position === "fixed" ? r.top : r.top + scrollY); });
    return { over, small: [...new Set(small)].slice(0, 6), ctaY: ys.length ? Math.min.apply(null, ys) : null, vh: innerHeight };
  });
}
(async () => {
  const b = await chromium.launch({ executablePath: process.env.CHROME || undefined });
  const views = [["p430", { width: 430, height: 932 }, "light"], ["p430d", { width: 430, height: 932 }, "dark"], ["p390", { width: 390, height: 844 }, "light"], ["desk", { width: 1280, height: 860 }, "light"]];
  for (const [tag, vp, scheme] of views) {
    const ctx = await b.newContext({ viewport: vp, deviceScaleFactor: tag === "desk" ? 1 : 2, colorScheme: scheme, locale: "vi-VN" });
    await ctx.route("**/config.js*", r => r.fulfill({ contentType: "application/javascript", body: `window.CARE_API=${JSON.stringify(API)};` }));
    await ctx.route(API, r => r.fulfill({ contentType: "application/json", body: JSON.stringify(stub(JSON.parse(r.request().postData() || "{}"))) }));
    await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
    const page = await ctx.newPage(), errs = [];
    page.on("pageerror", e => errs.push(e.message));
    page.on("console", m => { if (m.type() === "error" && !/ERR_FAILED|net::|Failed to load resource/.test(m.text())) errs.push(m.text()); });
    page.on("dialog", d => d.accept());
    for (const u of PAGES) {
      await page.goto("http://localhost:8790" + u); await page.waitForTimeout(250);
      const c = await checks(page, u, tag);
      const name = tag + "_" + (u === "/" ? "home" : u.replace(/^\/|\/$/g, "").replace(/\//g, "-"));
      await page.screenshot({ path: path.join(OUT, name + ".png"), fullPage: true });
      res.push({ tag, u, ...c, errs: errs.splice(0) });
    }
    if (tag === "p430") {
      /* luồng: trang dịch vụ → đặt lịch (chọn món, bỏ trống, điền, gửi) */
      await page.goto("http://localhost:8790/dich-vu/giat-hap/"); await page.click(".phead .btn:not(.ghost)");
      await page.waitForSelector("#picker .pk"); const open = await page.$$eval("#picker details[open] summary", s => s.map(x => x.childNodes[0].textContent));
      await page.click('.pk[data-m="G04"] [data-d="1"]'); await page.click('.pk[data-m="G04"] [data-d="1"]'); await page.click('#hang [data-h="Signature"]');
      const est = await page.textContent("#estT");
      await page.click("#bLich"); const e1 = await page.textContent("#eLich");
      await page.fill("[name=ten]", "Lan"); await page.fill("[name=sdt]", "0905123456"); await page.click('#cach [data-c="Giao nhận tận nơi"]'); await page.click("#bLich"); const e2 = await page.textContent("#eLich");
      await page.fill("[name=diaChi]", "12 Bạch Đằng"); await page.screenshot({ path: path.join(OUT, "flow1_form.png"), fullPage: true });
      await page.click("#bLich"); await page.waitForSelector("#okLich[style*=flex]"); await page.screenshot({ path: path.join(OUT, "flow2_ok.png") });
      /* luồng váy cưới: tính giá → đặt lịch có ghi chú */
      await page.goto("http://localhost:8790/dich-vu/vay-cuoi/"); await page.waitForSelector("#cDang input");
      await page.check('#cDang input[value="W05"]'); await page.click('#cDuoi label:has(input[value="W13"])'); await page.check('#cThem input[value="W21"]');
      const tong = await page.textContent("#cTong"); await page.locator(".calc").screenshot({ path: path.join(OUT, "flow3_calc.png") });
      await page.click("#cDat"); await page.waitForSelector("#picker details[open] .pk"); const gc = await page.inputValue("[name=ghiChu]"); const wd = await page.textContent("#estT");
      /* tra cứu */
      await page.goto("http://localhost:8790/?tra=PC-0211-01"); await page.waitForURL(/tra-cuu/); await page.fill("[name=so]", "3456"); await page.click("#bTra"); await page.waitForSelector("#res[style*=block]");
      await page.screenshot({ path: path.join(OUT, "flow4_tracuu.png"), fullPage: true });
      /* menu điện thoại */
      await page.goto("http://localhost:8790/"); await page.click("[data-open-menu]"); await page.waitForTimeout(150); await page.screenshot({ path: path.join(OUT, "flow5_menu.png") });
      console.log("FLOW", JSON.stringify({ open, est, e1, e2, tong, gc, wd, errs: errs.splice(0) }));
    }
    await ctx.close();
  }
  await b.close(); srv.close();
  let bad = 0;
  for (const r of res) {
    const warn = [];
    if (r.over > 0) warn.push("tràn ngang " + r.over + "px");
    if (r.small.length) warn.push("nút nhỏ: " + r.small.join(" | "));
    if (r.errs.length) warn.push("lỗi JS: " + r.errs.join(" | "));
    if (r.ctaY != null && r.ctaY > r.vh) warn.push("nút chính ở y=" + r.ctaY + " ngoài màn đầu " + r.vh);
    if (warn.length) bad++;
    console.log((warn.length ? "!! " : "ok ") + r.tag.padEnd(6) + r.u.padEnd(24) + "cta=" + r.ctaY + (warn.length ? "  " + warn.join("; ") : ""));
  }
  console.log(bad ? bad + " trang cần sửa" : "Tất cả trang đạt"); process.exit(bad ? 1 : 0);
})();
srv.listen(8790);
