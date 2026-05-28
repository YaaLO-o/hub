import re

CN_NUM = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
          "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}

TABLE_HEADERS = {"品名", "编码", "备注", "件数", "每件数量", "总数量", "单价",
                 "总金额", "货款", "运费", "重量", "总重量", "重量/KG", "总重量/KG",
                 "货号", "图片"}

NUMERIC_FIELDS = {"件数", "每件数量", "总数量", "单价", "总金额", "货款",
                  "运费", "重量", "总重量"}


def _parse_int(s: str) -> int | None:
    s = s.strip()
    if s.isdigit():
        return int(s)
    if s in CN_NUM:
        return CN_NUM[s]
    return None


def _clean_text(text: str) -> str:
    text = text.replace("：", ":").replace("；", ";").replace("．", ".")
    text = text.replace("（", "(").replace("）", ")")
    lines = [re.sub(r'[ \t]+', ' ', line) for line in text.split('\n')]
    return '\n'.join(lines)


def _strip_units(val: str) -> str:
    """Strip trailing unit words like 元, 只, 副, 个, 台, 件, 盒, KG, kg."""
    return re.sub(r'\s*(元|只|副|个|台|件|盒|条|只|双|瓶|袋|箱|米|斤|KG|kg|克|千克)$', '', val.strip())


def _try_parse_tabular(text: str) -> list[dict] | None:
    """Try to parse as tabular format (columns separated by tabs or 2+ spaces)."""
    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    if len(lines) < 2:
        return None

    first = lines[0]
    has_tabs = '\t' in first

    # Detect split pattern from first line
    if has_tabs:
        headers = [h.strip() for h in first.split('\t') if h.strip()]
    else:
        parts = re.split(r'\s{2,}', first)
        headers = [p.strip() for p in parts if p.strip()]

    if not headers:
        return None

    known = sum(1 for h in headers if h in TABLE_HEADERS)
    if known < 2:
        return None

    results = []
    for line in lines[1:]:
        line = re.sub(r'^\d+[\.\)、]\s*', '', line).strip()
        if not line:
            continue

        if has_tabs:
            vals = [v.strip() for v in line.split('\t') if v.strip()]
        else:
            vals = re.split(r'\s{2,}', line)
            vals = [v.strip() for v in vals if v.strip()]

        if len(vals) < 2:
            continue

        product = {}
        for i, h in enumerate(headers):
            if i >= len(vals):
                break
            val = vals[i]
            # Map header to canonical field name
            field = h.replace("/KG", "").strip()
            if field in NUMERIC_FIELDS:
                cleaned = _strip_units(val)
                try:
                    cleaned_val = cleaned.replace(',', '').replace('，', '')
                    product[field] = float(cleaned_val) if '.' in cleaned_val else int(cleaned_val)
                except ValueError:
                    product[field] = val
            else:
                product[field] = val

        if product.get("品名"):
            results.append(product)

    return results if results else None


def parse_product_text(text: str) -> list[dict]:
    """Parse product info text. Supports two formats:

    1. Tabular (pasted from Excel):
        品名    编码    备注    件数    单价
        智能手表  SW-1001  xxx     50     899

    2. Key-value (one field per line or inline):
        货号:150# 单价:6元 一件240个
    """
    # Try tabular format FIRST (before cleaning, which collapses spaces)
    tabular = _try_parse_tabular(text)
    if tabular is not None:
        return tabular

    # Fallback: key-value format
    text = _clean_text(text)

    blocks = re.split(r'\n\s*\n', text.strip())
    if len(blocks) <= 1:
        lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
        lines = [re.sub(r'^\d+[\.\)、]\s*', '', l).strip() for l in lines]

        # Count lines that look like "label:value" pairs
        kv_lines = [l for l in lines if re.match(r'.+?[:：]\s*\S', l)]

        if len(kv_lines) == len(lines) and len(lines) > 2:
            # All lines are key-value → treat as one multi-line product
            blocks = [' '.join(lines)]
        else:
            marker_lines = [l for l in lines if re.match(r'(?:货号|品名|编号)\s*[:：]', l)]
            blocks = marker_lines if len(marker_lines) > 1 else lines
    else:
        blocks = [b.strip() for b in blocks if b.strip()]

    blocks = [re.sub(r'^\d+[\.\)、]\s*', '', b).strip() for b in blocks]

    results = []
    for block in blocks:
        product = _parse_single(block)
        if product.get("品名"):
            results.append(product)
    return results


def _parse_single(text: str) -> dict:
    """Parse a single product's key-value format text."""
    data = {}

    # 品名优先，货号次之
    m = re.search(r'品名\s*[:：]?\s*(\S+)', text)
    if m:
        data["品名"] = m.group(1).strip()
    else:
        m = re.search(r'货号\s*[:：]?\s*(\S+)', text)
        if m:
            data["品名"] = m.group(1).strip()

    m = re.search(r'单价\s*[:：]?\s*([\d.]+)', text)
    if m:
        data["单价"] = float(m.group(1))

    m = re.search(r'(?<!货)编码\s*[:：]?\s*(\S+)', text)
    if m:
        data["编码"] = m.group(1).strip()

    m = re.search(r'备注\s*[:：]?\s*(.+?)(?=\s+\S+[:：]?\s*|\s*$)', text)
    if m:
        data["备注"] = m.group(1).strip()

    # === 件数 + 每件数量 ===
    pkg, qty_per = None, None

    m = re.search(r'([零一二三四五六七八九十两\d]+)\s*件\s*([\d]+)\s*个', text)
    if m:
        pkg = _parse_int(m.group(1))
        qty_per = int(m.group(2))
    else:
        m = re.search(r'每[件箱]\s*([\d]+)\s*个', text)
        if m:
            qty_per = int(m.group(1))

        m = re.search(r'件数\s*[:：]?\s*([零一二三四五六七八九十两\d]+)', text)
        if m:
            pkg = _parse_int(m.group(1))

        m = re.search(r'每[件箱]数量\s*[:：]?\s*([\d]+)', text)
        if m:
            qty_per = int(m.group(1))

        if qty_per is None:
            m = re.search(r'([\d]+)\s*个\s*[/／]\s*[件箱]', text)
            if m:
                qty_per = int(m.group(1))

        if qty_per is None:
            m = re.search(r'([\d]+)\s*个', text)
            if m:
                qty_per = int(m.group(1))

        if pkg is None and qty_per is not None:
            m = re.search(r'([零一二三四五六七八九十两\d])\s*件', text)
            if m:
                pkg = _parse_int(m.group(1))

        if pkg is None and qty_per is not None:
            if '件' in text:
                pkg = 1

    if pkg is not None:
        data["件数"] = pkg
    if qty_per is not None:
        data["每件数量"] = qty_per

    m = re.search(r'重量(?:/KG)?\s*[:：]?\s*([\d.]+)', text)
    if m:
        data["重量"] = float(m.group(1))

    m = re.search(r'总重量(?:/KG)?\s*[:：]?\s*([\d.]+)', text)
    if m:
        data["总重量"] = float(m.group(1))

    m = re.search(r'货款\s*[:：]?\s*([\d.]+)', text)
    if m:
        data["货款"] = float(m.group(1))

    m = re.search(r'运费\s*[:：]?\s*([\d.]+)', text)
    if m:
        data["运费"] = float(m.group(1))

    return data
