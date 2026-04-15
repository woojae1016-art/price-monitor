"""
온라인 가격 모니터링 자동화 - 2단계
게이트웨이 URL에서 상품번호 추출 → 직접 상품 페이지 접속 → 판매자 추출

의존성: pip install beautifulsoup4
"""

import re
import time
import requests
import urllib3
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 전역 Session 제거 — fetch_soup에서 매번 새로 생성해 차단 회피
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://search.naver.com/",
}


def detect_mall(url):
    host = urlparse(url).netloc.lower()
    if "coupang" in host:                    return "coupang"
    if "11st" in host:                       return "11st"
    if "gmarket" in host or "link.gmarket" in host: return "gmarket"
    if "auction" in host or "link.auction" in host: return "auction"
    return "unknown"


def resolve_url(url, mall_type):
    """
    네이버 API 게이트웨이 URL → 실제 상품 페이지 URL 변환
    - 11번가: Gateway.tmall?prdNo=12345 → /products/12345
    - G마켓:  link.gmarket.co.kr/gate/pcs?item-no=12345 → item.gmarket.co.kr/Item?goodscode=12345
    - 옥션:   link.auction.co.kr/gate/pcs?item-no=B12345 → itempage3.auction.co.kr/DetailView.aspx?itemno=B12345
    """
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)

    if mall_type == "11st":
        prd_no = qs.get("prdNo", [None])[0]
        if prd_no:
            return "https://www.11st.co.kr/products/%s" % prd_no
        # 이미 직접 URL인 경우
        if "/products/" in url:
            return url

    elif mall_type == "gmarket":
        item_no = qs.get("item-no", [None])[0]
        if item_no:
            return "https://item.gmarket.co.kr/Item?goodscode=%s" % item_no
        # 이미 직접 URL인 경우
        if "goodscode" in url:
            return url

    elif mall_type == "auction":
        item_no = qs.get("item-no", [None])[0]
        if item_no:
            return "https://itempage3.auction.co.kr/DetailView.aspx?itemno=%s" % item_no
        # 이미 직접 URL인 경우
        if "itemno" in url:
            return url

    return url


def fetch_soup(url):
    try:
        s = requests.Session()
        s.verify = False
        s.headers.update(_HEADERS)
        resp = s.get(url, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            return BeautifulSoup(resp.text, "html.parser")
    except Exception:
        pass
    return None


def to_int(val):
    try:
        return int(re.sub(r"[^\d]", "", str(val)))
    except Exception:
        return None


def clean_name(name):
    if not name:
        return None
    name = re.sub(r"[♩♪♫♬★☆◆◇●○]", "", name).strip()
    if len(name) < 2 or re.match(r"^\d+$", name):
        return None
    return name


def get_seller_and_price(url, mall_type):
    """실제 상품 페이지 URL로 접속해서 판매자·가격 추출."""
    # 게이트웨이 URL → 실제 상품 페이지 URL 변환
    direct_url = resolve_url(url, mall_type)

    soup = fetch_soup(direct_url)
    if not soup:
        return None, None

    seller = None
    price  = None

    # ── 판매자 ──────────────────────────────────
    if mall_type in ("gmarket", "auction"):
        el = soup.select_one("span.text__seller")
        if el:
            seller = clean_name(el.get_text(strip=True))

    elif mall_type == "11st":
        for sel in ["h4.c_product_seller_title", "h1.c_product_store_title"]:
            el = soup.select_one(sel)
            if el:
                seller = clean_name(el.get_text(strip=True))
                if seller:
                    break

    # ── 가격 ────────────────────────────────────
    for sel in ["span.total-price strong", "em#productPrice",
                "#goods_price strong", "em.price_num",
                "strong[class*='price']", "em[class*='price']",
                "span[class*='price'] strong", ".price-real strong"]:
        el = soup.select_one(sel)
        if el:
            p = to_int(el.get_text())
            if p and p >= 1000:
                price = p
                break

    return seller, price


OPEN_MARKETS = {"쿠팡", "11번가", "G마켓", "옥션", "Gmarket", "Auction"}


def enrich_items_with_seller(items):
    """오픈마켓 항목에 real_seller, real_price 추가."""
    targets = [
        (i, item) for i, item in enumerate(items)
        if item.get("mallName", "") in OPEN_MARKETS
    ]
    if not targets:
        return items

    print("\n  [2단계] 오픈마켓 판매자 크롤링 (%d건)..." % len(targets))

    for idx, (item_i, item) in enumerate(targets, 1):
        link   = item.get("link", "")
        mall   = item.get("mallName", "")
        lprice = to_int(item.get("lprice", 0)) or 0

        print("    [%03d/%03d] %s " % (idx, len(targets), mall), end="", flush=True)

        try:
            mall_type = detect_mall(link)

            if mall_type == "coupang":
                items[item_i]["real_seller"] = mall
                items[item_i]["real_price"]  = lprice
                print("-> (쿠팡 차단) %s원" % "{:,}".format(lprice))

            elif mall_type in ("gmarket", "auction", "11st"):
                seller, price = get_seller_and_price(link, mall_type)
                seller = seller or mall
                price  = price  or lprice
                items[item_i]["real_seller"] = seller
                items[item_i]["real_price"]  = price
                print("-> %s | %s원" % (seller, "{:,}".format(price)))

            else:
                items[item_i]["real_seller"] = mall
                items[item_i]["real_price"]  = lprice
                print("-> (미지원) %s원" % "{:,}".format(lprice))

        except Exception as e:
            print("-> [실패] %s" % str(e))
            items[item_i]["real_seller"] = mall
            items[item_i]["real_price"]  = lprice

        time.sleep(1.5)  # 회사 네트워크 차단 방지

    print("  [2단계] 완료\n")
    return items