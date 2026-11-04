#!/usr/bin/env python3
"""لارسی - خزشگر خودکار فروشگاه‌های عطر.

استفاده:
    python crawler/larsi_crawler.py https://example.com --out docs/data --deep

روش‌های کشف (به ترتیب): API فروشگاه ووکامرس، نقشه سایت (sitemap)، پیمایش صفحات فهرست.
خروجی: products.json، products.csv و products.xlsx
"""
import argparse
import csv
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

FIELDS = [
    ("name", "نام محصول"), ("english", "نام انگلیسی / مدل"), ("brand", "برند"),
    ("country", "کشور سازنده"), ("category", "دسته‌بندی"), ("type", "نوع محصول"),
    ("gender", "جنسیت"), ("volume", "حجم"), ("concentration", "غلظت"),
    ("family", "خانواده بویایی"), ("top", "نت ابتدایی"), ("heart", "نت میانی"),
    ("base", "نت پایه"), ("scent", "توصیف رایحه / طبع"), ("season", "فصل مناسب"),
    ("occasion", "مناسبت استفاده"), ("longevity", "ماندگاری"), ("sillage", "پخش بو"),
    ("price", "قیمت"), ("sku", "کد محصول"), ("image", "آدرس تصویر محصول"),
    ("description", "توضیحات"), ("source", "آدرس منبع"),
]
KEYS = [k for k, _ in FIELDS]
WIDTHS = {"name": 34, "english": 26, "brand": 16, "category": 22, "top": 28, "heart": 28,
          "base": 28, "scent": 30, "description": 60, "image": 48, "source": 48, "price": 18}

UA = ("Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/124.0 Mobile Safari/537.36")


class Stopped(Exception):
    pass


# ---------------------------------------------------------------- utils
_DIG = {ord(c): str(i) for i, c in enumerate("۰۱۲۳۴۵۶۷۸۹")}
_DIG.update({ord(c): str(i) for i, c in enumerate("٠١٢٣٤٥٦٧٨٩")})


def norm(s):
    s = str(s or "").translate(_DIG).replace("ي", "ی").replace("ك", "ک")
    s = re.sub(r"[\u200c\u200f\u200e]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def empty():
    p = {k: "" for k in KEYS}
    p["type"] = "نامشخص"
    return p


def fill(a, b):
    for k in KEYS:
        if (not a.get(k) or (k == "type" and a[k] == "نامشخص")) and b.get(k) and b[k] != "نامشخص":
            a[k] = b[k]
    return a


def pkey(p):
    return re.sub(r"\s+", "", norm("%s|%s|%s" % (p["name"] or p["english"], p["brand"], p["volume"])).lower())


def html_text(h):
    if not h:
        return ""
    h = re.sub(r"<\s*(br|/p|/li|/div|/h\d|/tr|/td)[^>]*>", lambda m: " . " + m.group(0), str(h), flags=re.I)
    return norm(BeautifulSoup(h, "html.parser").get_text(" "))


def price_num(s):
    m = re.search(r"\d+(?:\.\d+)?", norm(s).replace(",", "").replace("٬", ""))
    return float(m.group(0)) if m else None


# ---------------------------------------------------------------- field parsing
def parse_gender(t):
    t = norm(t)
    m = re.search(r"(مردانه|آقایان|مردان|\bmen\b|\bmale\b)", t, re.I)
    w = re.search(r"(زنانه|بانوان|خانم|\bwomen\b|\bfemale\b)", t, re.I)
    if re.search(r"(یونیسکس|unisex|مشترک)", t, re.I) or (m and w):
        return "مشترک"
    return "مردانه" if m else ("زنانه" if w else "")


def parse_conc(t):
    t = norm(t)
    for pat, val in [(r"extrait|اکستریت|اکسترایت", "اکستریت دو پرفیوم"),
                     (r"\bedp\b|ادو ?پرفیوم|eau de parfum", "ادو پرفیوم"),
                     (r"\bedt\b|ادو ?تویلت|ادوتویلت|eau de toilette", "ادو تویلت"),
                     (r"\bedc\b|ادو ?کلن|cologne", "ادو کلن"),
                     (r"parfum|پرفیوم|پارفوم", "پرفیوم")]:
        if re.search(pat, t, re.I):
            return val
    return ""


def parse_season(t):
    t = norm(t)
    out = [n for n, r in [("بهار", r"بهار|spring"), ("تابستان", r"تابستان|summer"),
                          ("پاییز", r"پاییز|fall|autumn"), ("زمستان", r"زمستان|winter")]
           if re.search(r, t, re.I)]
    return "، ".join(out)


def parse_type(t):
    t = norm(t)
    for pat, val in [(r"تستر|tester", "تستر"), (r"سمپل|دکانت|decant|sample", "سمپل"),
                     (r"مینیاتور|miniature", "مینیاتوری"), (r"اماراتی|emirat|\buae\b", "اماراتی"),
                     (r"اورجینال|original|اصل", "اورجینال")]:
        if re.search(pat, t, re.I):
            return val
    return ""


def parse_volume(t):
    m = re.search(r"(\d{1,4}(?:\.\d+)?)\s*(میلی ?لیتر|میل|ml)\b", norm(t), re.I)
    return m.group(1) + " میلی‌لیتر" if m else ""


def apply_text(p, text):
    text = norm(text)
    if not text:
        return
    sep, tail = r"\s*[:：\-–]\s*", r"([^.\n؛|]{2,140})"

    def n(pat):
        m = re.search(pat, text, re.I)
        return re.sub(r"[\s،,]+$", "", m.group(1).strip()) if m else ""

    if not p["top"]:
        p["top"] = n(r"(?:نت(?:\s*های)?\s*(?:ابتدایی|بالایی|بالا|سر|اولیه|اول)|top\s*notes?)" + sep + tail)
    if not p["heart"]:
        p["heart"] = n(r"(?:نت(?:\s*های)?\s*(?:میانی|قلب|وسط|میانه)|heart\s*notes?|middle\s*notes?)" + sep + tail)
    if not p["base"]:
        p["base"] = n(r"(?:نت(?:\s*های)?\s*(?:پایه|انتهایی|پایانی|ته)|base\s*notes?)" + sep + tail)
    p["gender"] = p["gender"] or parse_gender(text)
    p["concentration"] = p["concentration"] or parse_conc(text)
    p["season"] = p["season"] or parse_season(text)
    p["volume"] = p["volume"] or parse_volume(text)
    if p["type"] in ("", "نامشخص"):
        p["type"] = parse_type(text) or "نامشخص"
    if not p["country"]:
        p["country"] = n(r"(?:کشور سازنده|ساخت کشور|made in|country of origin)\s*[:：\-–]?\s*([^\s.،,|]{2,20})")


def attrs_of(root):
    m = {}
    for tr in root.find_all("tr"):
        th, td = tr.find("th"), tr.find("td")
        if th and td:
            m[norm(th.get_text(" "))] = norm(td.get_text(" "))
    for li in root.find_all(["li", "dt"]):
        t = norm(li.get_text(" "))
        mm = re.search(r"[:：]", t)
        if mm and 0 < mm.start() < 30 and len(t) < 300:
            m[t[:mm.start()].strip()] = t[mm.start() + 1:].strip()
    return m


def apply_attrs(p, m):
    def setk(k, v):
        if v and not p[k]:
            p[k] = v

    for l, v in m.items():
        if not v:
            continue
        l = norm(l)
        if "نت" in l or re.search(r"notes?", l, re.I):
            if re.search(r"بالا|ابتدا|سر|اولیه|top", l, re.I):
                setk("top", v)
            elif re.search(r"میان|قلب|وسط|heart|middle", l, re.I):
                setk("heart", v)
            elif re.search(r"پایه|انتها|پایان|ته|base", l, re.I):
                setk("base", v)
            continue
        if re.search(r"برند|brand|سازنده", l, re.I) and "کشور" not in l:
            setk("brand", v)
        elif re.search(r"کشور|ساخت|origin", l, re.I):
            setk("country", v)
        elif re.search(r"جنسیت|gender|مناسب برای", l, re.I):
            setk("gender", parse_gender(v))
        elif re.search(r"حجم|volume", l, re.I):
            setk("volume", parse_volume(v) or v)
        elif re.search(r"غلظت|نوع عطر|concentration", l, re.I):
            setk("concentration", parse_conc(v) or v)
        elif re.search(r"خانواده|گروه بویایی|دسته|family", l, re.I):
            setk("family", v)
        elif re.search(r"ماندگاری|longevity", l, re.I):
            setk("longevity", v)
        elif re.search(r"پخش|سیلاژ|sillage", l, re.I):
            setk("sillage", v)
        elif re.search(r"فصل|season", l, re.I):
            setk("season", parse_season(v) or v)
        elif re.search(r"مناسبت|occasion", l, re.I):
            setk("occasion", v)
        elif re.search(r"طبع|رایحه|scent", l, re.I):
            setk("scent", v)
        elif re.search(r"نوع محصول|اصالت", l):
            t = parse_type(v)
            if t and p["type"] in ("", "نامشخص"):
                p["type"] = t


# ---------------------------------------------------------------- page parsing
def jsonld_products(soup, source):
    out = []

    def walk(o):
        if isinstance(o, list):
            for x in o:
                walk(x)
            return
        if not isinstance(o, dict):
            return
        typ = o.get("@type", [])
        typ = typ if isinstance(typ, list) else [typ]
        if "Product" in typ:
            p = empty()
            im = o.get("image")
            im = im[0] if isinstance(im, list) and im else im
            if isinstance(im, dict):
                im = im.get("url")
            brand = o.get("brand")
            brand = brand.get("name") if isinstance(brand, dict) else brand
            offers = o.get("offers")
            offers = offers[0] if isinstance(offers, list) and offers else offers
            offers = offers if isinstance(offers, dict) else {}
            price = offers.get("price") or offers.get("lowPrice")
            p["name"] = norm(o.get("name"))
            p["brand"] = norm(brand)
            p["price"] = (str(price) + (" " + offers["priceCurrency"] if offers.get("priceCurrency") else "")) if price else ""
            p["image"] = urljoin(source, im) if im else ""
            p["description"] = norm(o.get("description"))
            p["sku"] = str(o.get("sku") or "")
            p["source"] = urljoin(source, o["url"]) if o.get("url") else source
            p["_text"] = p["name"] + " . " + p["description"]
            out.append(p)
        for v in o.values():
            if isinstance(v, (dict, list)):
                walk(v)

    for s in soup.find_all("script", type="application/ld+json"):
        try:
            walk(json.loads(s.string or s.get_text()))
        except Exception:
            pass
    return out


def parse_node(node, source):
    p = empty()
    g = lambda sel: norm((node.select_one(sel) or {}).get_text(" ") if node.select_one(sel) else "")
    img = node.find("img")
    p["name"] = g(".woocommerce-loop-product__title,.product-title,h2,h3,h1") or norm(img.get("alt") if img else "")
    p["price"] = g("[class*=price]")
    im = ""
    if img:
        im = img.get("data-src") or img.get("data-lazy-src") or ((img.get("srcset") or "").split() or [""])[0] or img.get("src") or ""
    p["image"] = urljoin(source, im) if im else ""
    a = node.find("a", href=True)
    p["source"] = urljoin(source, a["href"]) if a else source
    p["description"] = g(".description,.product-description")
    p["_text"] = p["name"] + " . " + p["description"]
    apply_attrs(p, attrs_of(node))
    return p


def _meta(soup, name):
    t = soup.find("meta", attrs={"property": name}) or soup.find("meta", attrs={"name": name})
    return (t.get("content") if t else "") or ""


def parse_doc(raw, source):
    """برمی‌گرداند (items, next_url, mode)  mode: detail | list"""
    soup = BeautifulSoup(raw, "html.parser")
    ld = jsonld_products(soup, source)
    nodes = [n for n in soup.select("li.product,.product-small,.product-item,.product-card,div.product,[itemtype*=Product]")
             if not n.select_one("li.product,.product-small,.product-item")]
    nx = soup.select_one("link[rel=next]") or soup.select_one("a.next,a[rel=next],.next.page-numbers")
    next_url = urljoin(source, nx["href"]) if nx and nx.get("href") else ""
    h1 = soup.find("h1")
    is_detail = len(ld) == 1 or (not ld and len(nodes) <= 1 and (h1 or _meta(soup, "og:title")))
    if is_detail:
        p = ld[0] if ld else empty()
        if not p["name"]:
            t = norm(h1.get_text(" ") if h1 else _meta(soup, "og:title") or (soup.title.get_text() if soup.title else ""))
            p["name"] = re.sub(r"\s*[|–-]\s*[^|–-]*$", "", t) or t
        if not p["image"]:
            im = _meta(soup, "og:image")
            if not im:
                t = soup.select_one(".woocommerce-product-gallery img,.product img")
                im = t.get("src") if t else ""
            p["image"] = urljoin(source, im) if im else ""
        if not p["description"]:
            p["description"] = norm(_meta(soup, "og:description") or _meta(soup, "description"))
        p["source"] = p["source"] or source
        if not p["price"]:
            t = soup.select_one("p.price,.price,[itemprop=price]")
            p["price"] = norm(t.get_text(" ")) if t else ""
        short = norm(" . ".join(html_text(str(e)) for e in soup.select(
            ".woocommerce-product-details__short-description,#tab-description,.product-description,[itemprop=description]")))
        orig = p["description"]
        if short and len(short) > len(p["description"]):
            p["description"] = short[:1500]
        root = soup.select_one(".product,main,article") or soup.body or soup
        apply_attrs(p, attrs_of(root))
        apply_text(p, " . ".join([p["name"], orig, p["description"], short]))
        p.pop("_text", None)
        p["english"] = p["english"] or (p["name"] if re.search(r"[a-z]", p["name"], re.I) else "")
        return [p], "", "detail"
    items = ld if len(ld) > 1 else [x for x in (parse_node(n, source) for n in nodes) if x["name"] or x["image"]]
    for p in items:
        apply_text(p, p.pop("_text", p["name"]))
        p["english"] = p["english"] or (p["name"] if re.search(r"[a-z]", p["name"], re.I) else "")
    return items, next_url, "list"


def from_store(x):
    p = empty()
    p["name"] = html_text(x.get("name"))
    p["source"] = x.get("permalink") or ""
    p["sku"] = x.get("sku") or ""
    ims = x.get("images") or []
    p["image"] = (ims[0].get("src") or ims[0].get("thumbnail") or "") if ims else ""
    pr = x.get("prices") or {}
    try:
        if pr.get("price") not in (None, ""):
            v = round(float(pr["price"]) / 10 ** int(pr.get("currency_minor_unit") or 0))
            p["price"] = "%d %s" % (v, pr.get("currency_symbol") or pr.get("currency_code") or "")
    except (ValueError, TypeError):
        pass
    br = x.get("brands") or []
    p["brand"] = html_text(br[0].get("name")) if br else ""
    cats = [html_text(c.get("name")) for c in (x.get("categories") or []) if c.get("name")]
    p["category"] = "، ".join(cats)
    m = {}
    for a in x.get("attributes") or []:
        v = "، ".join(html_text(t.get("name")) for t in (a.get("terms") or []))
        if a.get("name") and v:
            m[norm(a["name"])] = v
    apply_attrs(p, m)
    sd, d = html_text(x.get("short_description")), html_text(x.get("description"))
    p["description"] = (sd + " " + d if sd and d and sd not in d else (d or sd))[:1500]
    apply_text(p, " . ".join([p["name"], sd, d, p["category"]]))
    p["english"] = p["name"] if re.search(r"[a-z]", p["name"], re.I) else ""
    return p


# ---------------------------------------------------------------- crawler
class Crawler:
    def __init__(self, base, limit=0, deep=True, workers=4, delay=0.0, proxy=None, log=print, timeout=30):
        u = urlparse(base if "://" in base else "https://" + base)
        self.base = u.geturl()
        self.origin = "%s://%s" % (u.scheme, u.netloc)
        self.limit, self.deep, self.workers, self.delay, self.log = limit, deep, workers, delay, log
        self.timeout = timeout
        self.stop = False
        self.s = requests.Session()
        retry = Retry(total=3, backoff_factor=1.0, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=("GET",))
        self.s.mount("http://", HTTPAdapter(max_retries=retry, pool_maxsize=16))
        self.s.mount("https://", HTTPAdapter(max_retries=retry, pool_maxsize=16))
        self.s.headers.update({"User-Agent": UA, "Accept-Language": "fa,en;q=0.8",
                               "Accept": "text/html,application/xhtml+xml,application/xml,application/json;q=0.9,*/*;q=0.8"})
        if proxy:
            self.s.proxies = {"http": proxy, "https": proxy}

    def get(self, url):
        if self.stop:
            raise Stopped()
        if self.delay:
            time.sleep(self.delay)
        r = self.s.get(url, timeout=self.timeout)
        r.raise_for_status()
        if not r.encoding or r.encoding.lower() == "iso-8859-1":
            r.encoding = r.apparent_encoding or "utf-8"
        return r.text

    def pool(self, seq, fn):
        with ThreadPoolExecutor(self.workers) as ex:
            list(ex.map(fn, seq))

    # -- discovery
    def via_api(self, items):
        for pg in range(1, 400):
            try:
                t = self.get("%s/wp-json/wc/store/v1/products?per_page=100&page=%d" % (self.origin, pg))
                a = json.loads(t)
            except Stopped:
                raise
            except Exception:
                break
            if not isinstance(a, list) or not a:
                break
            items.extend(from_store(x) for x in a)
            self.log("API فروشگاه: %d محصول" % len(items))
            if len(a) < 100 or (self.limit and len(items) >= self.limit):
                break

    def sitemap_urls(self):
        seeds = []
        try:
            for l in re.findall(r"^sitemap:\s*(\S+)", self.get(self.origin + "/robots.txt"), re.I | re.M):
                seeds.append(l.strip())
        except Stopped:
            raise
        except Exception:
            pass
        seeds += [self.origin + p for p in ("/sitemap_index.xml", "/wp-sitemap.xml", "/sitemap.xml", "/product-sitemap.xml")]
        seen, products, others, q = set(), [], [], list(seeds)
        while q and len(seen) < 80:
            u = q.pop(0)
            if u in seen:
                continue
            seen.add(u)
            try:
                t = self.get(u)
            except Stopped:
                raise
            except Exception:
                continue
            if not re.search(r"<(urlset|sitemapindex)", t, re.I):
                continue
            locs = [re.sub(r"&amp;", "&", m) for m in re.findall(r"<loc>\s*(?:<!\[CDATA\[)?\s*([^<\]\s]+)", t, re.I)]
            if re.search(r"<sitemapindex", t, re.I):
                prod = [l for l in locs if re.search("product", l, re.I) and not re.search("cat|tag|brand", l, re.I)]
                q += prod or [l for l in locs if not re.search("post|blog|author|attachment|category|tag", l, re.I)]
            elif re.search("product", u, re.I) and not re.search("cat|tag|brand", u, re.I):
                products += locs
            else:
                others += locs
            self.log("نقشه سایت: %d آدرس" % (len(products) + len(others)))
        return list(dict.fromkeys(products)), list(dict.fromkeys(others))

    def via_sitemap(self, items):
        products, others = self.sitemap_urls()
        urls = products or [u for u in others if re.search(r"/(product|products)/", urlparse(u).path, re.I)]
        if self.limit:
            urls = urls[:self.limit]
        done = [0]

        def one(u):
            try:
                p = parse_doc(self.get(u), u)[0][0]
                if p["name"] and (p["price"] or p["image"] or p["description"]):
                    p["source"] = p["source"] or u
                    items.append(p)
            except Stopped:
                raise
            except Exception:
                pass
            done[0] += 1
            if done[0] % 10 == 0 or done[0] == len(urls):
                self.log("خواندن صفحات محصول %d/%d" % (done[0], len(urls)))

        self.pool(urls, one)

    def via_listing(self, items):
        for st in dict.fromkeys([self.base, self.origin + "/shop/", self.origin + "/products/"]):
            cur = st
            for pg in range(1, 300):
                if not cur:
                    break
                try:
                    its, nxt, mode = parse_doc(self.get(cur), cur)
                except Stopped:
                    raise
                except Exception:
                    break
                if mode == "list" and its:
                    items.extend(its)
                    self.log("فهرست صفحه %d: %d محصول" % (pg, len(items)))
                    cur = nxt
                    if self.limit and len(items) >= self.limit:
                        break
                else:
                    if not items and its and its[0]["name"]:
                        items.append(its[0])
                    break
            if items:
                return

    def run(self, items=None):
        items = [] if items is None else items
        method = ""
        try:
            self.via_api(items)
            method = "WooCommerce API" if items else ""
            if not items:
                self.via_sitemap(items)
                method = "نقشه سایت" if items else ""
            if not items:
                self.via_listing(items)
                method = "پیمایش فهرست" if items else ""
            if self.deep and items:
                need = [p for p in items if p["source"].startswith(self.origin)
                        and not (p["top"] and p["heart"] and p["base"] and p["gender"] and p["brand"])]
                if self.limit:
                    need = need[:max(self.limit, 1500)]
                done = [0]

                def d(p):
                    try:
                        r = parse_doc(self.get(p["source"]), p["source"])[0][0]
                        im = p["image"]
                        fill(p, r)
                        p["image"] = p["image"] or im
                    except Stopped:
                        raise
                    except Exception:
                        pass
                    done[0] += 1
                    if done[0] % 10 == 0 or done[0] == len(need):
                        self.log("تکمیل جزئیات %d/%d" % (done[0], len(need)))

                if need:
                    self.log("تکمیل جزئیات %d محصول..." % len(need))
                    self.pool(need, d)
        except Stopped:
            self.log("متوقف شد؛ نتایج جزئی ذخیره می‌شود")
        seen, out = set(), []
        for p in items:
            k = p["source"] + "|" + pkey(p)
            if k not in seen:
                seen.add(k)
                out.append(p)
        return out, (method or "نامشخص")


# ---------------------------------------------------------------- export
def quality(items):
    rows = []
    for k, l in FIELDS:
        n = sum(1 for p in items if p.get(k) and not (k == "type" and p["type"] == "نامشخص"))
        rows.append((l, n, len(items) - n, round(n / len(items), 2) if items else 0))
    return rows


def write_xlsx(items, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "محصولات"
    ws.sheet_view.rightToLeft = True
    head = [l for _, l in FIELDS] + ["قیمت (عدد)"]
    ws.append(head)
    for p in items:
        ws.append([p.get(k, "") for k in KEYS] + [price_num(p.get("price"))])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="171717")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 26
    for i, k in enumerate(KEYS + ["_n"], 1):
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(k, 15)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    q = wb.create_sheet("کیفیت داده")
    q.sheet_view.rightToLeft = True
    q.append(["فیلد", "پر شده", "خالی", "درصد تکمیل"])
    for r in quality(items):
        q.append(list(r))
    for c in q[1]:
        c.font = Font(bold=True)
    q.column_dimensions["A"].width = 24
    wb.save(path)


def save_all(items, out_dir, host, method, source_url):
    import os
    d = os.path.join(out_dir, host)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, "products.json"), "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=1)
    with open(os.path.join(d, "products.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([l for _, l in FIELDS])
        for p in items:
            w.writerow([("'" + str(p[k]) if str(p[k])[:1] in "=+-@" and p[k] else p[k]) for k in KEYS])
    write_xlsx(items, os.path.join(d, "products.xlsx"))
    idx_path = os.path.join(out_dir, "index.json")
    try:
        with open(idx_path, encoding="utf-8") as f:
            idx = json.load(f)
    except Exception:
        idx = []
    idx = [x for x in idx if x.get("host") != host]
    idx.insert(0, {"host": host, "url": source_url, "count": len(items), "method": method,
                   "updated": datetime.now(timezone.utc).isoformat(timespec="seconds")})
    with open(idx_path, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)
    return d


def main(argv=None):
    ap = argparse.ArgumentParser(description="خزشگر خودکار فروشگاه عطر")
    ap.add_argument("url")
    ap.add_argument("--out", default="docs/data")
    ap.add_argument("--limit", type=int, default=0, help="حداکثر تعداد محصول (۰ = همه)")
    ap.add_argument("--no-deep", action="store_true", help="صفحهٔ تک‌تک محصولات را نخوان")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--delay", type=float, default=0.2, help="مکث بین درخواست‌ها (ثانیه)")
    ap.add_argument("--proxy", default=None)
    a = ap.parse_args(argv)
    c = Crawler(a.url, a.limit, not a.no_deep, a.workers, a.delay, a.proxy)
    host = urlparse(c.origin).netloc.replace(":", "_")
    print("شروع:", c.origin, flush=True)
    items, method = c.run()
    if not items:
        print("محصولی پیدا نشد. اگر سایت از IP خارج از ایران جواب نمی‌دهد، این اسکریپت را روی دستگاه خودت اجرا کن.", file=sys.stderr)
        return 2
    d = save_all(items, a.out, host, method, a.url)
    print("پایان: %d محصول (%s) -> %s" % (len(items), method, d))
    return 0


if __name__ == "__main__":
    sys.exit(main())
