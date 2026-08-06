import re
import time
from dataclasses import dataclass, asdict
from urllib.parse import urljoin, urlparse, parse_qs

import pandas as pd
import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


BASE_URL = "https://andar.co.kr"
TARGET_CATE_NOS = ["2100", "2111"]
OUTFILE = f"andar_할인율모니터링_{time.strftime('%Y%m%d')}.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}


@dataclass
class ProductRow:
    상세상품URL: str
    품번: str
    품명: str
    소비자가: str
    할인가: str
    할인율: str
    리뷰수: str
    별점: str


def normalize_price(text: str) -> str:
    if not text:
        return ""
    t = re.sub(r"\s+", "", str(text))
    m = re.search(r"([0-9]{1,3}(?:,[0-9]{3})*)원", t)
    return (m.group(1) + "원") if m else ""


def keep_number_with_comma(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"[^\d,]", "", str(text))


def keep_float(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"([0-5](?:\.[0-9])?)", str(text))
    return m.group(1) if m else ""


def split_skus(model: str) -> list[str]:
    if not model:
        return [""]
    parts = [p.strip() for p in str(model).replace(" ", "").split("_") if p.strip()]
    return parts if parts else [str(model).strip()]


def to_int_price(text: str) -> int:
    if not text:
        return 0
    m = re.search(r"([0-9]{1,3}(?:,[0-9]{3})*)", str(text))
    if not m:
        return 0
    return int(m.group(1).replace(",", ""))


def calc_discount_rate(consumer: str, sale: str) -> str:
    c = to_int_price(consumer)
    s = to_int_price(sale)
    if c <= 0 or s <= 0:
        return ""
    if s >= c:
        return "0%"
    rate = (c - s) / c * 100.0
    return f"{int(round(rate))}%"


def get_product_no(url: str) -> str:
    try:
        q = parse_qs(urlparse(url).query)
        return q.get("product_no", [""])[0]
    except Exception:
        return ""


def build_driver(headless: bool = True) -> webdriver.Chrome:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1400,900")
    opts.add_argument("--lang=ko-KR")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_page_load_timeout(60)
    return driver


def wait_dom_ready(driver: webdriver.Chrome, timeout: int = 30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def is_valid_product_url(href: str) -> bool:
    if not href:
        return False

    href_lower = href.lower()

    if "product/detail.html" not in href_lower:
        return False
    if "product_no=" not in href_lower:
        return False

    bad_keywords = [
        "javascript:",
        "/board/",
        "/article/",
        "review",
    ]
    if any(bad in href_lower for bad in bad_keywords):
        return False

    return True


def get_category_product_links(driver: webdriver.Chrome) -> list[tuple[str, str]]:
    selectors = [
        ".xans-product-listnormal a[href*='product/detail.html']",
        "ul.prdList a[href*='product/detail.html']",
        "li[id^='anchorBoxId_'] a[href*='product/detail.html']",
        ".prdList .thumbnail a[href*='product/detail.html']",
        ".prdList .description a[href*='product/detail.html']",
    ]

    a_tags = []
    for sel in selectors:
        found = driver.find_elements(By.CSS_SELECTOR, sel)
        if found:
            a_tags = found
            break

    if not a_tags:
        a_tags = driver.find_elements(By.CSS_SELECTOR, "a[href*='product/detail.html']")

    page_links: list[tuple[str, str]] = []
    seen_on_page = set()

    for a in a_tags:
        href = (a.get_attribute("href") or "").strip()
        if not href:
            continue

        if href.startswith("/"):
            href = urljoin(BASE_URL, href)

        href = href.split("#")[0]

        if not is_valid_product_url(href):
            continue

        pn = get_product_no(href)
        if not pn:
            continue
        if pn in seen_on_page:
            continue

        seen_on_page.add(pn)
        page_links.append((pn, href))

    return page_links


def collect_links_pagination(
    driver: webdriver.Chrome,
    start_url: str,
    max_pages: int = 500,
) -> list[str]:
    dedup: dict[str, str] = {}
    no_new_streak = 0

    for page in range(1, max_pages + 1):
        url = f"{start_url}&page={page}"
        driver.get(url)
        wait_dom_ready(driver, 30)
        time.sleep(0.8)

        page_links = get_category_product_links(driver)

        if len(page_links) == 0:
            print(f"  - page={page}: 상품 링크 없음 -> 종료")
            break

        before = len(dedup)
        for pn, href in page_links:
            if pn not in dedup:
                dedup[pn] = href
        after = len(dedup)
        added = after - before

        print(f"  - page={page}: 링크 {len(page_links)}개 탐색, 누적 {after}개 (+{added})")

        if added == 0:
            no_new_streak += 1
        else:
            no_new_streak = 0

        if no_new_streak >= 3:
            print(f"  - page={page}: 새 상품 추가 3페이지 연속 없음 -> 종료")
            break

    return list(dedup.values())


def extract_text_by_label(soup: BeautifulSoup, label: str) -> str:
    text = soup.get_text("\n", strip=True)
    pattern_inline = re.compile(rf"{re.escape(label)}\s*([^\n]+)")
    m = pattern_inline.search(text)
    if m:
        return m.group(1).strip()
    return ""


def pick_first_text(soup: BeautifulSoup, selectors: list[str]) -> str:
    for sel in selectors:
        el = soup.select_one(sel)
        if el:
            t = el.get_text(" ", strip=True)
            if t:
                return t
    return ""


def extract_price_cafe24(soup: BeautifulSoup) -> tuple[str, str]:
    consumer = normalize_price(
        pick_first_text(
            soup,
            [
                "#span_product_price_custom",
                "span#span_product_price_custom",
            ],
        )
    )

    sale = normalize_price(
        pick_first_text(
            soup,
            [
                "#span_product_price_sale",
                "span#span_product_price_sale",
                "#span_product_price_text",
                "span#span_product_price_text",
            ],
        )
    )

    if not consumer:
        consumer = (
            normalize_price(extract_text_by_label(soup, "소비자가"))
            or normalize_price(extract_text_by_label(soup, "PRICE"))
        )

    if not sale:
        sale = (
            normalize_price(extract_text_by_label(soup, "할인판매가"))
            or normalize_price(extract_text_by_label(soup, "판매가"))
            or normalize_price(extract_text_by_label(soup, "PRICE"))
        )

    if not consumer and sale:
        consumer = sale

    return consumer, sale


def fast_fetch_model_candidates(url: str, timeout: int = 15) -> list[str]:
    """
    Selenium 대신 requests로 상세페이지 HTML만 빠르게 받아
    품번 후보를 뽑아낸다.
    """
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
    except Exception:
        return []

    html = resp.text
    soup = BeautifulSoup(html, "lxml")

    candidates = []

    model = extract_text_by_label(soup, "MODEL SIZE INFO")
    if model:
        candidates.extend(split_skus(model))

    model2 = extract_text_by_label(soup, "모델명")
    if model2:
        candidates.extend(split_skus(model2))

    # 페이지 전체 텍스트에서 E로 시작하는 SKU 패턴도 후보로 추가
    text = soup.get_text(" ", strip=True)
    regex_hits = re.findall(r"\b(E[A-Z0-9]+(?:_[A-Z0-9]+)*)\b", text)
    for hit in regex_hits:
        candidates.extend(split_skus(hit))

    cleaned = []
    seen = set()
    for c in candidates:
        x = c.strip().upper()
        if not x:
            continue
        if x in seen:
            continue
        seen.add(x)
        cleaned.append(x)

    return cleaned


def is_e_product_fast(url: str) -> bool:
    skus = fast_fetch_model_candidates(url)
    return any(sku.startswith("E") for sku in skus)


def extract_review_count_selenium(driver: webdriver.Chrome, timeout_sec: float = 8.0) -> str:
    end = time.time() + timeout_sec
    last_seen = ""

    xpaths = [
        "//*[self::a or self::button or self::li or self::span][contains(normalize-space(.), '리뷰')]",
        "//*[contains(@href,'review')][contains(normalize-space(.), '리뷰')]",
    ]

    while time.time() < end:
        for xp in xpaths:
            elems = driver.find_elements(By.XPATH, xp)
            for e in elems[:60]:
                t = (e.text or "").strip()
                if not t:
                    continue
                last_seen = t
                m = re.search(r"리뷰\s*\(\s*([0-9,]+)\s*개\s*\)", t)
                if m:
                    n = m.group(1)
                    if n.replace(",", "") != "0":
                        return f"{n}개"
        time.sleep(0.5)

    soup = BeautifulSoup(driver.page_source, "lxml")
    links = soup.select("a[href*='/article/review/4/']")
    uniq = {a.get("href") for a in links if a.get("href")}
    if uniq:
        return f"{len(uniq)}개"

    m = re.search(r"리뷰\s*\(\s*([0-9,]+)\s*개\s*\)", last_seen)
    if m:
        return f"{m.group(1)}개"

    return ""


def extract_rating(driver: webdriver.Chrome, soup: BeautifulSoup, timeout_sec: float = 6.0) -> str:
    candidate_selectors = [
        ".crema-product-reviews-score__score",
        ".crema-product-reviews-score__average",
        ".review_score",
        ".prd-review .score",
        "[class*='crema'][class*='score']",
        "[class*='rating']",
    ]

    for css in candidate_selectors:
        el = soup.select_one(css)
        if el:
            t = el.get_text(" ", strip=True)
            m = re.search(r"([0-5](?:\.[0-9])?)", t)
            if m:
                return m.group(1)

    end = time.time() + timeout_sec
    while time.time() < end:
        soup_live = BeautifulSoup(driver.page_source, "lxml")
        for css in candidate_selectors:
            el = soup_live.select_one(css)
            if el:
                t = el.get_text(" ", strip=True)
                m = re.search(r"([0-5](?:\.[0-9])?)", t)
                if m:
                    return m.group(1)

        text_all = soup_live.get_text(" ", strip=True)
        nums = re.findall(r"\b([0-5]\.[0-9])\b", text_all)
        preferred = [x for x in nums if 3.0 <= float(x) <= 5.0]
        if preferred:
            return preferred[0]

        time.sleep(0.7)

    return ""


def parse_detail_page(driver: webdriver.Chrome, url: str) -> ProductRow:
    driver.get(url)
    wait_dom_ready(driver, 30)
    time.sleep(0.8)

    soup = BeautifulSoup(driver.page_source, "lxml")

    name = ""
    for sel in ["div.headingArea h2", "div.headingArea h1", "h2", "h1"]:
        el = soup.select_one(sel)
        if el and el.get_text(strip=True):
            name = el.get_text(strip=True)
            break

    model = extract_text_by_label(soup, "MODEL SIZE INFO")
    if not model:
        model = extract_text_by_label(soup, "모델명")
    model = (model or "").strip()

    consumer, sale = extract_price_cafe24(soup)

    discount_rate = ""
    if consumer and sale:
        discount_rate = calc_discount_rate(consumer, sale)

    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.6);")
        time.sleep(0.6)
    except Exception:
        pass

    review_cnt = extract_review_count_selenium(driver, timeout_sec=8.0)
    rating = extract_rating(driver, soup, timeout_sec=6.0)

    return ProductRow(
        상세상품URL=url,
        품번=model,
        품명=name,
        소비자가=keep_number_with_comma(consumer),
        할인가=keep_number_with_comma(sale),
        할인율=discount_rate,
        리뷰수=keep_number_with_comma(review_cnt),
        별점=keep_float(rating),
    )


def main():
    driver = build_driver(headless=True)

    try:
        print("[1/4] 카테고리 페이지 링크 수집 중...")

        all_links_dict: dict[str, str] = {}

        for cate_no in TARGET_CATE_NOS:
            start_url = f"{BASE_URL}/product/list.html?cate_no={cate_no}"
            print(f"  - cate_no={cate_no} 수집 시작")

            links = collect_links_pagination(
                driver,
                start_url=start_url,
                max_pages=500,
            )

            for url in links:
                pn = get_product_no(url)
                if pn and pn not in all_links_dict:
                    all_links_dict[pn] = url

        all_links = list(all_links_dict.values())
        print(f"  - 전체 링크 수집 완료: {len(all_links)}개")

        print("[2/4] requests로 품번 선확인 후 E 상품만 추리는 중...")
        e_links = []

        for i, url in enumerate(all_links, start=1):
            try:
                if is_e_product_fast(url):
                    e_links.append(url)
                    print(f"  ({i}/{len(all_links)}) E 포함")
                else:
                    print(f"  ({i}/{len(all_links)}) 제외")
            except Exception as e:
                print(f"  ({i}/{len(all_links)}) 선확인 실패 - {e}")

        print(f"  - E 상품 링크 수: {len(e_links)}개")

        print("[3/4] E 상품만 Selenium 상세 파싱 중...")
        rows: list[ProductRow] = []

        for i, url in enumerate(e_links, start=1):
            try:
                row = parse_detail_page(driver, url)

                for sku in split_skus(row.품번):
                    if not sku.strip().upper().startswith("E"):
                        continue

                    rows.append(
                        ProductRow(
                            상세상품URL=row.상세상품URL,
                            품번=sku.strip().upper(),
                            품명=row.품명,
                            소비자가=row.소비자가,
                            할인가=row.할인가,
                            할인율=row.할인율,
                            리뷰수=row.리뷰수,
                            별점=row.별점,
                        )
                    )

                print(f"  ({i}/{len(e_links)}) OK - {row.품명[:30]}")
                time.sleep(0.4)

            except TimeoutException:
                print(f"  ({i}/{len(e_links)}) TIMEOUT - {url}")

            except Exception as e:
                print(f"  ({i}/{len(e_links)}) SKIP/FAIL - {url} / {e}")

        print("[4/4] 엑셀 저장 중...")
        df = pd.DataFrame([asdict(r) for r in rows])
        cols = ["상세상품URL", "품번", "품명", "소비자가", "할인가", "할인율", "리뷰수", "별점"]
        df = df.reindex(columns=cols)

        df = df.drop_duplicates(subset=["상세상품URL", "품번"], keep="first").reset_index(drop=True)

        df.to_excel(OUTFILE, index=False)
        print(f"완료ㄟ(≧◇≦)ㄏ -> {OUTFILE}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()