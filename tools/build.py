#!/usr/bin/env python3
"""Dựng website CHÉRIE Care nhiều trang từ src/*.html + bố cục chung.
Chạy: python3 tools/build.py  (tạo các file index.html theo đường dẫn, kèm số phiên bản để trình duyệt tải bản mới)."""
import os, re, time, html

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
VER = time.strftime("%Y%m%d%H%M", time.gmtime())
SITE = "https://care.cherie.io.vn"
ZALO = "https://zalo.me/84377159189"

DICH_VU = [
    ("giat-hap", "/dich-vu/giat-hap/", "Giặt và hấp", "Đầm, áo dài, vest · 2 mức chăm"),
    ("vay-cuoi", "/dich-vu/vay-cuoi/", "Váy cưới", "Giặt theo dáng váy, cất giữ chống ẩm"),
    ("sua-do", "/dich-vu/sua-do/", "Sửa đồ", "Thợ may ngồi tiệm, hẹn thử đồ"),
    ("hap-tan-noi", "/dich-vu/hap-tan-noi/", "Hấp lấy liền và tận nơi", "60 phút tại quầy, hấp tại khách sạn, tiệc cưới"),
]
NAV = [("bang-gia", "/bang-gia/", "Bảng giá"), ("goi", "/goi/", "Gói và thẻ"), ("cam-ket", "/cam-ket/", "Cam kết"), ("doanh-nghiep", "/doanh-nghiep/", "Doanh nghiệp"), ("tra-cuu", "/tra-cuu/", "Tra cứu đơn")]

PAGES = [
    # key, file src, đường dẫn, tiêu đề thẻ, mô tả, có thanh đặt lịch dưới đáy
    ("home", "home.html", "/", "CHÉRIE Care · Giặt khô, hấp, sửa đồ nữ ở Đà Nẵng", "Chăm đầm tiệc, áo dài, váy cưới theo từng món ở Đà Nẵng. 4 ảnh tình trạng gửi Zalo, đền theo giá trị khai báo, làm lại miễn phí 7 ngày, thợ may ngồi tiệm.", True),
    ("giat-hap", "giat-hap.html", "/dich-vu/giat-hap/", "Giặt và hấp đầm, áo dài, vest · CHÉRIE Care", "Giặt khô, giặt nước nhẹ, hấp form đồ nữ theo từng món ở Đà Nẵng. Hai mức chăm Chuẩn và Signature, bảng giá công khai.", True),
    ("vay-cuoi", "vay-cuoi.html", "/dich-vu/vay-cuoi/", "Giặt và cất giữ váy cưới · CHÉRIE Care", "Giặt váy cưới theo dáng, chất liệu, đuôi váy; gói cất giữ hút chân không chống ẩm cho mùa nồm Đà Nẵng. Tính giá ngay.", True),
    ("sua-do", "sua-do.html", "/dich-vu/sua-do/", "Sửa đồ tại tiệm · CHÉRIE Care", "Lên lai, bóp eo, thu nhỏ đầm, thay khoá kéo, đính lại cườm. Thợ may ngồi tiệm ở Đà Nẵng, làm cùng lúc với giặt hấp.", True),
    ("hap-tan-noi", "hap-tan-noi.html", "/dich-vu/hap-tan-noi/", "Hấp lấy liền và chăm đồ tận nơi · CHÉRIE Care", "Hấp lấy liền 60 phút tại quầy, hấp tại khách sạn, chuẩn bị đồ trước giờ đi tiệc, hấp chỉnh váy áo tại tiệc cưới ở Đà Nẵng.", True),
    ("bang-gia", "bang-gia.html", "/bang-gia/", "Bảng giá · CHÉRIE Care", "Bảng giá giặt hấp, váy cưới, sửa đồ, chăm thêm, dịch vụ tận nơi và gói của CHÉRIE Care Đà Nẵng. Tính theo món.", True),
    ("goi", "goi.html", "/goi/", "Gói và thẻ trả trước · CHÉRIE Care", "Gói Tết, giữ đồ qua mùa nồm, gói sau cưới và thẻ Care trả trước có thưởng.", True),
    ("cam-ket", "cam-ket.html", "/cam-ket/", "Cam kết và đền bù · CHÉRIE Care", "Đền theo giá trị khai báo và tuổi món đồ, làm lại miễn phí 7 ngày, quy trình 6 bước có 4 ảnh tình trạng.", True),
    ("doanh-nghiep", "doanh-nghiep.html", "/doanh-nghiep/", "Doanh nghiệp và đối tác · CHÉRIE Care", "Giặt hấp cho tiệm cho thuê đồ, studio, khách sạn, homestay, wedding planner ở Đà Nẵng. Lịch cố định, giá theo tháng.", False),
    ("dat-lich", "dat-lich.html", "/dat-lich/", "Đặt lịch nhận đồ · CHÉRIE Care", "Chọn món, xem tạm tính và đặt lịch gửi đồ tại quầy hoặc giao nhận tận nơi ở Đà Nẵng.", False),
    ("tra-cuu", "tra-cuu.html", "/tra-cuu/", "Tra cứu đơn · CHÉRIE Care", "Tra tình trạng đồ đang chăm bằng mã phiếu và 4 số cuối Zalo.", False),
    ("theo-doi", "theo-doi.html", "/theo-doi/", "Theo dõi đồ của bạn · CHÉRIE Care", "Xem từng bước chăm đồ, người làm, giờ và ảnh tại xưởng.", False),
    ("404", "404.html", "/404.html", "Không tìm thấy trang · CHÉRIE Care", "Trang không tồn tại.", False),
]

def nav_html(key):
    dv_on = key in [d[0] for d in DICH_VU]
    menu = "".join(f'<a href="{p}"{" aria-current=page" if k == key else ""}>{t}<small>{s}</small></a>' for k, p, t, s in DICH_VU)
    out = f'<div class="dd"><button type="button" aria-expanded="false" aria-haspopup="true"{" aria-current=page" if dv_on else ""}>Dịch vụ <svg width="12" height="12" viewBox="0 0 24 24" aria-hidden="true" style="stroke:currentColor;stroke-width:2;fill:none"><path d="m6 9 6 6 6-6"/></svg></button><div class="menu">{menu}</div></div>'
    out += "".join(f'<a href="{p}"{" aria-current=page" if k == key else ""}>{t}</a>' for k, p, t in NAV)
    return out

def drawer_html(key):
    a = lambda k, p, t, s="": f'<a class="lnk" href="{p}"{" aria-current=page" if k == key else ""}>{t}{f"<span>{s}</span>" if s else ""}</a>'
    return ('<div class="drawer" id="drawer"><div class="pn" role="dialog" aria-modal="true" aria-label="Menu">'
            '<div class="ph"><a class="lk" href="/" style="font-size:17px"><span class="w">CHÉRIE</span><span class="d"></span><span class="b">Care</span></a>'
            '<button class="mbtn" type="button" data-close-menu aria-label="Đóng menu"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18"/></svg></button></div>'
            + a("home", "/", "Trang chủ") + '<div class="grp">Dịch vụ</div>' + "".join(a(k, p, t) for k, p, t, s in DICH_VU)
            + '<div class="grp">Thông tin</div>' + "".join(a(k, p, t) for k, p, t in NAV)
            + f'<div class="grp">Liên hệ</div><a class="lnk" href="{ZALO}" target="_blank" rel="noopener">Nhắn Zalo<span>0377 159 189</span></a>'
            '<a class="btn block" href="/dat-lich/" style="margin-top:10px">Đặt lịch nhận đồ</a></div></div>')

FOOTER = f'''<footer><div class="wrap fg">
<div class="brand"><span class="lk rev" style="font-size:22px"><span class="w">CHÉRIE</span><span class="d"></span><span class="b">Care</span></span><p class="tl">Chăm những bộ đồ cho ngày quan trọng của bạn.</p></div>
<div><h4>Dịch vụ</h4><ul>{"".join(f'<li><a href="{p}">{t}</a></li>' for k, p, t, s in DICH_VU)}</ul></div>
<div><h4>Thông tin</h4><ul>{"".join(f'<li><a href="{p}">{t}</a></li>' for k, p, t in NAV)}</ul></div>
<div class="lh"><h4>Liên hệ</h4><ul><li><a href="{ZALO}" target="_blank" rel="noopener">Zalo 0377 159 189</a></li><li><a href="/dat-lich/">Đặt lịch nhận đồ</a></li><li><a href="https://cherie.io.vn/">Thuê đầm CHÉRIE</a></li></ul></div>
<div class="sub"><span>Một thương hiệu của CHÉRIE · Mỗi lần, một chiếc mới.</span><span>Đà Nẵng</span></div>
</div></footer>'''

DOCK = f'<div class="dock" role="region" aria-label="Đặt lịch nhanh"><a class="btn" href="/dat-lich/">Đặt lịch</a><a class="btn ghost" href="{ZALO}" target="_blank" rel="noopener">Nhắn Zalo</a></div>'

def build():
    for key, src, path, title, desc, dock in PAGES:
        raw = open(os.path.join(SRC, src), encoding="utf-8").read()
        style = "".join(re.findall(r"<style>(.*?)</style>", raw, re.S))
        script = "".join(re.findall(r"<script>(.*?)</script>", raw, re.S))
        body = re.sub(r"<style>.*?</style>|<script>.*?</script>", "", raw, flags=re.S).strip()
        canon = SITE + ("/" if path == "/" else path)
        page = f'''<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>{html.escape(title)}</title>
<meta name="description" content="{html.escape(desc)}">{'<meta name="robots" content="noindex">' if key in ("theo-doi", "404") else ""}
<meta name="theme-color" content="#24476B">
{"" if key == "404" else f'<link rel="canonical" href="{canon}">'}
<meta property="og:title" content="{html.escape(title)}">
<meta property="og:description" content="{html.escape(desc)}">
<meta property="og:url" content="{canon}">
<meta property="og:type" content="website">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bodoni+Moda:opsz,wght@6..96,600&family=Playfair+Display:ital,wght@0,500;0,600;1,500&family=Be+Vietnam+Pro:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<link rel="stylesheet" href="/care.css?v={VER}">
<link rel="stylesheet" href="/site.css?v={VER}">
{f"<style>{style}</style>" if style.strip() else ""}
<script src="/config.js?v={VER}"></script>
<script src="/care.js?v={VER}"></script>
<script src="/site.js?v={VER}"></script>
</head>
<body{' class="has-dock"' if dock else ""}>
<a class="sr" href="#main">Bỏ qua menu</a>
<header class="top"><div class="wrap in">
<a class="lk" href="/" style="font-size:18px" aria-label="CHÉRIE Care, trang chủ"><span class="w">CHÉRIE</span><span class="d"></span><span class="b">Care</span></a>
<nav aria-label="Chính">{nav_html(key)}</nav>
<div class="r"><a class="btn sm" href="/dat-lich/">Đặt lịch</a><button class="mbtn" type="button" data-open-menu aria-label="Mở menu"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M4 12h16M4 17h16"/></svg></button></div>
</div></header>
{drawer_html(key)}
<main id="main">
{body}
</main>
{FOOTER}
{DOCK if dock else ""}
{f"<script>{script}</script>" if script.strip() else ""}
</body>
</html>
'''
        out = os.path.join(ROOT, path.lstrip("/"), "index.html") if path.endswith("/") else os.path.join(ROOT, path.lstrip("/"))
        os.makedirs(os.path.dirname(out), exist_ok=True)
        open(out, "w", encoding="utf-8").write(page)
    urls = "".join(f"<url><loc>{SITE}{p}</loc></url>" for k, s, p, *_ in PAGES if k not in ("404", "theo-doi"))
    open(os.path.join(ROOT, "sitemap.xml"), "w").write(f'<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">{urls}</urlset>\n')
    open(os.path.join(ROOT, "ver.txt"), "w").write(VER + "\n")
    print("đã dựng", len(PAGES), "trang, phiên bản", VER)

if __name__ == "__main__":
    build()
