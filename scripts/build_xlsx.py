#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_xlsx.py —— 把生成好的评论写成「对齐蒸汽电波参考表」的 .xlsx。
零第三方依赖，只用标准库（json/sys/zipfile/re）；XML 全部手工拼接字符串，
不 import xml、更不需要 openpyxl，避免环境里没装第三方库时翻车。

用法：
    python3 build_xlsx.py <输入.json> <输出.xlsx> [--stats] [--check]
    --stats  额外打印整池鸟瞰 + 逐批的「长度 × 收尾 × 句式」三轴分布 + 机器痕迹自查
    --check  任一批参差不达标时以退出码 2 退出并报失败批次（供生成循环「不达标自动重写该批」或 CI 硬门槛）

输入 JSON 结构：
{
  "account":      "蒸汽电波",          # 账号名，建议填；缺省则留空、照样出表
  "publish_date": "",                  # 发布日期，纯文本，可留空
  "order_date":   "",                  # 下单日期，纯文本，可留空
  "video_link":   "",                  # 默认始终留空，由用户发布后自己补；脚本正文即使带链接也别塞进来
  "batch_size":   20,                  # 每批次多少条，默认 20
  "comments":     ["评论1", "评论2", ...]    # 扁平列表，脚本自动切成 batch_size 条一批
}

加 --stats 会额外打印三轴分布 + 标点混用情况（含"空格分句"机器痕迹检测）；
加 --check 则在不达标时以退出码 2 退出，方便脚本化门槛。

表格布局（完全照抄参考表）：
  A1 账号        B1 {account}
  A2 发布日期    B2 {publish_date}
  A3 下单日期    B3 {order_date}
  A4 视频链接    B4 {video_link}      （A4:A6、B4:B6 合并）
  A7 总数        B7 {条数}
  批次评论：从 D 列起，每批次占一列、每列 batch_size 条（行 2..1+batch_size）。
           每 5 个批次为一组，组间空一列（复刻参考表 D–H · 空I · J–N 的观感）。
           批次表头「第一批次…第五批次」在每组内循环。
"""
import json
import sys
import zipfile

CN_NUM = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
GROUP = 5  # 每 5 个批次一组，组后留一空列


def col_letter(idx0):
    """0-based 列号 -> 字母（A=0）。支持到 ZZ。"""
    s = ""
    n = idx0
    while True:
        s = chr(ord("A") + n % 26) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def esc(t):
    return (
        str(t)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def batch_col_index(b):
    """第 b 个批次(0-based)落在哪一列(0-based)。起点 D=3，每组 5 列后跳一列。"""
    group, within = divmod(b, GROUP)
    return 3 + group * (GROUP + 1) + within


def build_cells(data):
    """返回 {(col0, row1): (value, is_number, style)} 与合并区、最大列。"""
    cells = {}
    S_WRAP = 1   # 自动换行样式
    S_HEAD = 2   # 表头/字段名样式（加粗底纹）

    def put(c0, r1, val, num=False, style=0):
        cells[(c0, r1)] = (val, num, style)

    # 左侧元数据块
    put(0, 1, "账号", style=S_HEAD);      put(1, 1, data.get("account", ""))
    put(0, 2, "发布日期", style=S_HEAD);  put(1, 2, data.get("publish_date", ""))
    put(0, 3, "下单日期", style=S_HEAD);  put(1, 3, data.get("order_date", ""))
    put(0, 4, "视频链接", style=S_HEAD);  put(1, 4, data.get("video_link", ""), style=S_WRAP)
    total = len(data.get("comments", []))
    put(0, 7, "总数", style=S_HEAD);      put(1, 7, total, num=True)

    merges = ["A4:A6", "B4:B6"]

    # 批次评论块
    comments = data.get("comments", [])
    bs = int(data.get("batch_size", 20) or 20)
    max_col = 1
    b = 0
    i = 0
    while i < len(comments):
        chunk = comments[i:i + bs]
        c0 = batch_col_index(b)
        max_col = max(max_col, c0)
        put(c0, 1, "第%s批次" % CN_NUM[b % GROUP], style=S_HEAD)
        for r, txt in enumerate(chunk):
            put(c0, 2 + r, txt, style=S_WRAP)
        b += 1
        i += bs

    return cells, merges, max_col, b


def sheet_xml(data):
    cells, merges, max_col, num_batches = build_cells(data)
    bs = int(data.get("batch_size", 20) or 20)
    max_row = max(1 + bs, 7)

    # 列宽
    cols_xml = ['<col min="1" max="1" width="11" customWidth="1"/>',
                '<col min="2" max="2" width="42" customWidth="1"/>']
    if max_col >= 3:
        cols_xml.append('<col min="4" max="%d" width="30" customWidth="1"/>' % (max_col + 1))
    cols_block = "<cols>%s</cols>" % "".join(cols_xml)

    # 行
    rows_xml = []
    for r in range(1, max_row + 1):
        row_cells = []
        for c0 in range(0, max_col + 1):
            if (c0, r) not in cells:
                continue
            val, num, style = cells[(c0, r)]
            ref = "%s%d" % (col_letter(c0), r)
            s_attr = ' s="%d"' % style if style else ""
            if num:
                row_cells.append('<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, esc(val)))
            else:
                if val == "":
                    continue
                row_cells.append(
                    '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (ref, s_attr, esc(val))
                )
        if row_cells:
            rows_xml.append('<row r="%d">%s</row>' % (r, "".join(row_cells)))

    merge_block = ""
    if merges:
        merge_block = '<mergeCells count="%d">%s</mergeCells>' % (
            len(merges), "".join('<mergeCell ref="%s"/>' % m for m in merges)
        )

    dim = "A1:%s%d" % (col_letter(max_col), max_row)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="%s"/>%s<sheetData>%s</sheetData>%s</worksheet>'
        % (dim, cols_block, "".join(rows_xml), merge_block)
    )


STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="2"><font><sz val="11"/><name val="等线"/></font>'
    '<font><b/><sz val="11"/><name val="等线"/></font></fonts>'
    '<fills count="3"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    '<alignment vertical="top" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1">'
    '<alignment vertical="center" wrapText="1"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="评论" sheetId="1" r:id="rId1"/></sheets></workbook>'
)

WB_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)


def write_xlsx(data, out_path):
    sheet = sheet_xml(data)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
        z.writestr("xl/styles.xml", STYLES_XML)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def analyze(comments):
    """按「长度 × 收尾 × 句式」三轴 + 机器痕迹算参差自查。
    返回 (report_lines, problems)：report_lines 是分布行；problems 是判为「不达标」
    的告警列表，供 --check 决定退出码。阈值判断一律用精确百分比，不做向下取整
    （旧版用 v*100//n 会把 40.9% 显示成 40% 且漏报，见修复记录）。"""
    import re
    n = len(comments)
    if n == 0:
        return ["  （无评论，空表）"], []   # 空表由 main 的完整性检查负责报，避免重复

    def pct(c):
        return c * 100.0 / n
    def show(c):
        return "%d(%d%%)" % (c, round(pct(c)))

    lines = []
    problems = []

    # 轴一 · 长度（别都挤中间）
    buckets = {"极短≤6": 0, "短7-14": 0, "中15-25": 0, "长26-40": 0, "超长40+": 0}
    for c in comments:
        L = len(c)
        if L <= 6:
            buckets["极短≤6"] += 1
        elif L <= 14:
            buckets["短7-14"] += 1
        elif L <= 25:
            buckets["中15-25"] += 1
        elif L <= 40:
            buckets["长26-40"] += 1
        else:
            buckets["超长40+"] += 1
    lines.append("  长短分布: " + "  ".join("%s=%s" % (k, show(v)) for k, v in buckets.items()))

    # 轴二 · 收尾（既不能全不加，也不能每条都工整收口）
    def ending(c):
        c = (c or "").rstrip()
        if not c:
            return "空"
        last = c[-1]
        if last in "。.":
            return "句号"
        if last in "！!":
            return "感叹"
        if last in "？?":
            return "问号"
        if last == "…":
            return "省略"
        if last in "~～":
            return "波浪"
        if last in "，,、":
            return "逗号"    # 逗号断句但结尾不收，也是一种「不收尾」
        if ("一" <= last <= "鿿") or last.isalnum():
            return "无标点"
        return "其他"        # 多为 emoji 收尾
    ends = {}
    for c in comments:
        e = ending(c)
        ends[e] = ends.get(e, 0) + 1
    lines.append("  收尾方式: " + "  ".join("%s=%s" % (k, show(v))
                 for k, v in sorted(ends.items(), key=lambda kv: -kv[1])))

    # 轴三 · 句式（启发式）：run-on＝长句一口气不断句；碎片＝极短无标点
    def is_runon(c):
        return len(c) >= 16 and not re.search(r"[，,。.！!？?…；;、~～]", c)
    runon = sum(1 for c in comments if is_runon(c))
    frag = sum(1 for c in comments if len(c) <= 6 and ending(c) == "无标点")
    longish = sum(1 for c in comments if len(c) >= 16)
    lines.append("  句式(启发): run-on一句到底=%d  碎片/情绪弹=%d" % (runon, frag))

    # 机器痕迹：中文字之间用空格代替标点（如「百度 发声」），健康值应为 0
    cjk_space = sum(1 for c in comments if re.search(r"[一-鿿] +[一-鿿]", c))
    lines.append("  空格分句: %d 条 %s" % (cjk_space, "✓" if cjk_space == 0 else "✗"))

    # —— 判不达标（进 problems；--check 据此返回非零退出码）——
    if cjk_space > 0:
        problems.append("有 %d 条「空格分句」（中文字之间夹空格），最假的机器痕迹，改成逗号/句号" % cjk_space)
    no_punctuation_pct = pct(ends.get("无标点", 0))
    if no_punctuation_pct < 35:
        problems.append("收尾太工整：无标点仅占 %.1f%%（<35%%）——手机评论不会每条都补收尾标点，至少约三分之一直接以文字收尾" % no_punctuation_pct)
    if no_punctuation_pct > 60:
        problems.append("收尾太统一：无标点占 %.1f%%（>60%%）——掺些句号/问号/感叹，但别回到每条都工整收口" % no_punctuation_pct)
    for style, cnt in ends.items():
        if style != "无标点" and pct(cnt) > 40:
            problems.append("收尾太统一：「%s」占 %.1f%%（>40%%）——真实评论区收尾应混着来，别都一样" % (style, pct(cnt)))
    if pct(buckets["中15-25"]) > 45:
        problems.append("长度太挤中间：中15-25字占 %.1f%%（>45%%）——多铺两字爆点和长评拉开落差" % pct(buckets["中15-25"]))
    if pct(buckets["极短≤6"]) < 15:
        problems.append("两字爆点太少：极短≤6字仅 %.1f%%（<15%%）——补些「离谱/绝了/典」" % pct(buckets["极短≤6"]))
    if buckets["超长40+"] > 0:
        problems.append("超长评论 %d 条，抖音评论区几乎没有，建议压到 0" % buckets["超长40+"])
    if longish >= 3 and runon == 0:
        problems.append("句式偏统一：%d 条长评全在规矩断句、无一句到底的 run-on，掺 1–2 条不断句的" % longish)

    return lines, problems


def report(comments, bs):
    """打印「整池鸟瞰 + 逐批自检」，返回 problems（已按批标注，供 --check 用）。

    三轴配额是按「每 bs 条一批」定义的，所以**判达标以逐批为准**：整池分布只作鸟瞰、
    不参与判定——否则一个全统一的烂批会被整池平均稀释掉，看着达标其实没达标（默认
    200 条=10 批时尤其危险）。逐批报失败批次序号，正好支撑「不达标自动重写该批」。"""
    problems = []
    n = len(comments)

    print("  【整池鸟瞰】(仅信息，不判定)")
    for ln in analyze(comments)[0]:
        print(ln)

    batches = [comments[i:i + bs] for i in range(0, n, bs)]
    single = len(batches) <= 1
    print("  【逐批自检】(判达标以此为准，配额按每批 %d 条定义)" % bs)
    for bi, batch in enumerate(batches, 1):
        if len(batch) < bs and not single:
            print("    第%d批：%d 条（末批不满，条数完整性另报，跳过三轴判定）" % (bi, len(batch)))
            continue
        bprobs = analyze(batch)[1]
        if bprobs:
            print("    第%d批 ✗ %d 项：" % (bi, len(bprobs)))
            for p in bprobs:
                print("        ⚠ " + p)
                problems.append("第%d批 · %s" % (bi, p))
        else:
            print("    第%d批 ✓" % bi)
    return problems


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if len(args) != 2:
        print("用法: python3 build_xlsx.py <输入.json> <输出.xlsx> [--stats] [--check]", file=sys.stderr)
        print("  --stats  打印整池鸟瞰 + 逐批的长度/收尾/句式三轴分布 + 机器痕迹自查", file=sys.stderr)
        print("  --check  任一批参差不达标（无标点<35%或>60%/单一标点>40%/长度挤中间/末列不满…）时以退出码 2 退出，", file=sys.stderr)
        print("           并报出失败批次序号，供生成循环「不达标自动重写该批」或 CI 做硬门槛", file=sys.stderr)
        sys.exit(1)
    in_path, out_path = args
    with open(in_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 输入清洗（一次到位，让 write_xlsx / 完整性检查 / 三轴自检看到同一份数据）：
    # 跳过 None 与空白条（否则会写出字面 "None" 单元格、或留空洞却仍计入总数），非字符串转成字符串。
    raw = data.get("comments", [])
    comments, dropped = [], 0
    for c in raw:
        s = "" if c is None else (c if isinstance(c, str) else str(c))
        if s.strip() == "":
            dropped += 1
            continue
        comments.append(s)
    data["comments"] = comments
    n = len(comments)

    write_xlsx(data, out_path)
    print("已写出 %s（%d 条评论）" % (out_path, n))
    if dropped:
        print("  ⚠ 输入有 %d 条为空/None，已跳过（未写入表、未计入总数）" % dropped)

    # 条数完整性自查：空表 / 末列不满 batch_size（避免"看似对齐实则缺角"的表）
    bs = int(data.get("batch_size", 20) or 20)
    integrity = []
    if n == 0:
        integrity.append("comments 为空，生成的是一张没有评论的空表")
    elif n % bs != 0:
        integrity.append("共 %d 条，非批次大小 %d 的整数倍，最后一批只有 %d 条（末列不满）"
                         % (n, bs, n % bs))
    for m in integrity:
        print("  ⚠ " + m)

    problems = []
    if "--stats" in flags or "--check" in flags:
        problems = report(comments, bs)

    if "--check" in flags:
        problems = integrity + problems
        if problems:
            print("参差自检不达标：%d 项 → 退出码 2（含失败批次，用于自动重写该批 / CI 门槛）"
                  % len(problems), file=sys.stderr)
            sys.exit(2)
        print("参差自检通过 ✓（逐批均达标）")


if __name__ == "__main__":
    main()
