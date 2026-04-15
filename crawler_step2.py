"""
온라인 가격 모니터링 자동화 - 2단계
게이트웨이 URL에서 상품번호 추출 → 직접 상품 페이지 접속 → 판매자 추출

의존성: pip install beautifulsoup4
"""

import os
import re
import time
import requests
import urllib3
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

_DEBUG = os.getenv("DEBUG_CRAWLER", "false").lower() == "true"

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://search.naver.com/",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}


def detect_mall(url):
    host = urlparse(url).netloc.lower()
    if "coupang" in host:
        return "coupang"
    if "11st" in host:
        return "11st"
    if "gmarket" in host or "link.gmarket" in host:
        return "gmarket"
    if "auction" in host or "link.auction" in host:
        return "auction"
    return "unknown"


def resolve_url(url, mall_type):
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)

    if mall_type == "11st":
        prd_no = qs.get("prdNo", [None])[0]
        if prd_no:
            return f"https://www.11st.co.kr/products/{prd_no}"
        if "/products/" in url:
            return url

    elif mall_type == "gmarket":
        item_no = qs.get("item-no", [None])[0] or qs.get("goodsCode", [None])[0]
        if item_no:
            return f"https://item.gmarket.co.kr/Item?goodscode={item_no}"
        if "goodscode" in url.lower() or "item.gmarket.co.kr" in url.lower():
            return url

    elif mall_type == "auction":
        item_no = qs.get("item-no", [None])[0]
        if item_no:
            return f"https://itempage3.auction.co.kr/DetailView.aspx?itemno={item_no}"
        if "itemno" in url.lower() or "itempage" in url.lower():
            return url

    return url


def _is_blocked_or_nonproduct(resp_text):
    text = (resp_text or "")[:5000].lower()
    block_markers = [
        "access denied",
        "robot",
        "captcha",
        "자동입력",
        "비정상적인 접근",
        "잠시 후 다시",
        "서비스 이용에 불편",
        "로그인이 필요",
    ]
    return any(m in text for m in block_markers)


def fetch_soup(url):
    try:
        s = requests.Session()
        s.verify = False
        s.headers.update(_HEADERS)
        resp = s.get(url, timeout=20, allow_redirects=True)
        if _DEBUG:
            print(f"      [fetch] req={url}")
            print(f"      [fetch] final={resp.url}")
            print(f"      [fetch] status={resp.status_code}")
        if resp.status_code == 200 and not _is_blocked_or_nonproduct(resp.text):
            return BeautifulSoup(resp.text, "html.parser"), resp.url
    except Exception as e:
        if _DEBUG:
            print(f"      [fetch error] {url} -> {e}")
    return None, None


def to_int(val):
    try:
        return int(re.sub(r"[^\d]", "", str(val)))
    except Exception:
        return None


def clean_name(name):
    if not name:
        return None
    name = re.sub(r"[♩♪♫♬★☆◆◇●○]", "", name).strip()
    name = re.sub(r"\s+", " ", name)
    if len(name) < 2 or re.match(r"^\d+$", name):
        return None
    return name


def _find_first_text(soup, selectors):
    for sel in selectors:
        try:
            el = soup.select_one(sel)
            if el:
                text = clean_name(el.get_text(" ", strip=True))
                if text:
                    return text, sel
        except Exception:
            pass
    return None, None


def _find_price(soup):
    price_selectors = [
        "span.total-price strong",
        "em#productPrice",
        "#goods_price strong",
        "em.price_num",
        "strong[class*='price']",
        "em[class*='price']",
        "span[class*='price'] strong",
        ".price-real strong",
        "meta[property='product:price:amount']",
        "meta[property='og:price:amount']",
    ]
    for sel in price_selectors:
        try:
            el = soup.select_one(sel)
            if not el:
                continue
            raw = el.get("content") if el.name == "meta" else el.get_text()
            p = to_int(raw)
            if p and p >= 1000:
                return p, sel
        except Exception:
            pass
    return None, None


def get_seller_and_price(url, mall_type):
    direct_url = resolve_url(url, mall_type)
    soup, final_url = fetch_soup(direct_url)
    if not soup:
        return None, None

    seller = None
    price = None

    if mall_type == "11st":
        seller, seller_sel = _find_first_text(soup, [
            "h4.c_product_seller_title",
            "h1.c_product_store_title",
            "[class*='seller']",
            "[class*='store']",
        ])
        if _DEBUG:
            print(f"      [11st] seller_sel={seller_sel} seller={seller}")

    elif mall_type == "gmarket":
        seller, seller_sel = _find_first_text(soup, [
            "span.text__seller",
            ".box__seller-name",
            ".text__seller-name",
            ".seller_name",
            "a[href*='Shop']",
            "a[href*='shop']",
            "a[href*='Seller']",
            "a[href*='seller']",
            "[class*='seller']",
            "[class*='shop']",
            "[class*='store']",
        ])
        if _DEBUG:
            print(f"      [gmarket] final_url={final_url}")
            print(f"      [gmarket] seller_sel={seller_sel} seller={seller}")

    elif mall_type == "auction":
        seller, seller_sel = _find_first_text(soup, [
            "span.text__seller",
            ".box__seller-name",
            ".text__seller-name",
            ".seller_name",
            "a[href*='Shop']",
            "a[href*='shop']",
            "a[href*='Seller']",
            "a[href*='seller']",
            "[class*='seller']",
            "[class*='shop']",
            "[class*='store']",
        ])
        if _DEBUG:
            print(f"      [auction] final_url={final_url}")
            print(f"      [auction] seller_sel={seller_sel} seller={seller}")

    price, price_sel = _find_price(soup)
    if _DEBUG:
        print(f"      [price] selector={price_sel} price={price}")

    return seller, price


OPEN_MARKETS = {"쿠팡", "11번가", "G마켓", "옥션", "Gmarket", "Auction"}


def enrich_items_with_seller(items):
    targets = [
        (i, item) for i, item in enumerate(items)
        if item.get("mallName", "") in OPEN_MARKETS
    ]
    if not targets:
        return items

    print("\n  [2단계] 오픈마켓 판매자 크롤링 (%d건)..." % len(targets))

    for idx, (item_i, item) in enumerate(targets, 1):
        link = item.get("link", "")
        mall = item.get("mallName", "")
        lprice = to_int(item.get("lprice", 0)) or 0

        print("    [%03d/%03d] %s " % (idx, len(targets), mall), end="", flush=True)

        try:
            mall_type = detect_mall(link)

            if mall_type == "coupang":
                items[item_i]["real_seller"] = mall
                items[item_i]["real_price"] = lprice
                print("-> (쿠팡 차단) %s원" % "{:,}".format(lprice))

            elif mall_type in ("gmarket", "auction", "11st"):
                seller, price = get_seller_and_price(link, mall_type)
                seller = seller or mall
                price = price or lprice
                items[item_i]["real_seller"] = seller
                items[item_i]["real_price"] = price
                print("-> %s | %s원" % (seller, "{:,}".format(price)))

            else:
                items[item_i]["real_seller"] = mall
                items[item_i]["real_price"] = lprice
                print("-> (미지원) %s원" % "{:,}".format(lprice))

        except Exception as e:
            print("-> [실패] %s" % str(e))
            items[item_i]["real_seller"] = mall
            items[item_i]["real_price"] = lprice

        time.sleep(1.5)

    print("  [2단계] 완료\n")
    return items
