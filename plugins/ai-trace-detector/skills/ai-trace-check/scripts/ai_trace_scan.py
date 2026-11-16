#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ai_trace_scan.py - 文档 AI 生成痕迹扫描器（只读，不修改文件）

支持: .docx .pptx .xlsx .pdf（以及 .txt/.md/.html 的正文痕迹检查）
一次传入多个文件时，会自动比对它们是否同源、批注是否被改过时间。
只依赖 Python 标准库；装有 pypdf 时 PDF 检查更完整（批注、正文）。

用法:
  python ai_trace_scan.py 文件或文件夹 [更多路径...]
      [--mode review|self|both] [--format md|json] [--out 报告路径]
      [--expect-author 姓名] [--expect-org 机构]

检查 9 类: 作者 / 路径 / 时间 / 批注与修订 / 文件结构 / 正文残留 / 官方标识 / 嵌入对象 / 文件外痕迹
（文件外痕迹指下载来源、隔离标记、文件系统时间，只有在文件原本所在的电脑上运行才查得到）
证据分级: strong(直接痕迹) / medium(间接疑点) / weak(弱提示) / info(仅供参考)
注意: 本工具判断的是"文件是否由程序或 AI 工具生成、粘贴、加工"，
      不能判断文字是不是人想出来的；任何一条都不是定论。
"""
import argparse
import datetime as dt
import hashlib
import html
import io
import itertools
import json
import logging
import os
import plistlib
import re
import statistics
import subprocess
import sys
import zipfile

logging.getLogger("pypdf").setLevel(logging.ERROR)

VERSION = "0.2.0"
UTC = dt.timezone.utc
TOL = dt.timedelta(hours=14)  # 时区误差容忍：很多程序把本地时间当 UTC 写入

CATS = {
    "author": "作者", "path": "路径", "time": "时间", "review": "批注与修订",
    "struct": "文件结构", "text": "正文残留", "label": "官方标识", "embed": "嵌入对象", "fs": "文件外痕迹",
}
LEVELS = ["strong", "medium", "weak", "info"]
LEVEL_ZH = {"strong": "强", "medium": "中", "weak": "弱", "info": "参考"}

# ---------------------------------------------------------------- 指纹库
# (正则, 说明, 级别)
GENERATOR_NAMES = [
    (r"python-docx", "python-docx 库的默认作者名（ChatGPT 等生成 Word 时常用）", "strong"),
    (r"^openpyxl$", "openpyxl 库的默认作者名（代码生成 Excel）", "strong"),
    (r"^un-named$", "docx (JavaScript) 库的默认作者名（Claude 等生成 Word 时常用）", "strong"),
    (r"pptxgenjs", "PptxGenJS 库的默认作者名（代码生成 PPT）", "strong"),
    (r"^steve canny$", "python-pptx 自带模板的作者（库作者本人）", "strong"),
    (r"apache poi|docx4j|aspose|xlsxwriter|phpword|phpspreadsheet|phpoffice|officegen|docxtemplater|open ?xml sdk",
     "文档生成库的名称", "medium"),
    (r"chatgpt|openai|\bgpt-?\d|claude|anthropic|copilot|gemini|deepseek|豆包|doubao|kimi|moonshot|文心|通义|tongyi|qwen|讯飞星火|智谱|perplexity|manus|gamma\.app|\bai assistant\b|ai助手|ai 助手",
     "AI 产品或厂商的名称", "strong"),
]
GENERIC_NAMES = re.compile(
    r"^(administrator|admin|user|用户|author|作者|owner|microsoft office user|microsoft office 用户|"
    r"windows user|windows 用户|root|test|pc|home|guest|anonymous|unknown|reviewer|审阅者)$", re.I)

APP_NAMES = [
    (r"python|openpyxl|apache poi|docx4j|aspose|openxml|phpword|phpoffice|docxtemplater|officegen|pptxgenjs|sheetjs|exceljs|xlsxwriter",
     "由文档生成库写出，不是办公软件保存的", "medium"),
    (r"chatgpt|openai|claude|anthropic|copilot|gemini|deepseek|豆包|kimi|文心|通义", "程序名直接是 AI 产品", "strong"),
]
NORMAL_APPS = re.compile(r"microsoft|wps|kingsoft|libreoffice|openoffice|pages|keynote|numbers|google|onlyoffice|永中|金山", re.I)

LIB_DATES = {
    "2013-12-23T23:15:00Z": "python-docx 自带模板的固定创建时间",
    "2013-01-27T09:14:16Z": "python-pptx 自带模板的固定创建时间",
}

PDF_PRODUCERS = [
    (r"reportlab", "ReportLab（Python 库，ChatGPT/Claude 生成 PDF 时最常用）", "medium"),
    (r"weasyprint|wkhtmltopdf|pdfkit|jspdf|fpdf|pyfpdf|pdf-lib|pdfmake|tcpdf|dompdf|mpdf|itext|pdfsharp|puppeteer|playwright|headlesschrome|headless chrome|matplotlib|pandoc|typst|prince",
     "程序化生成 PDF 的库或工具", "medium"),
    (r"skia/pdf", "Skia/PDF（Chrome 打印或无头浏览器导出，Google 文档导出也是它）", "weak"),
    (r"chatgpt|openai|claude|anthropic|copilot|gemini|deepseek|豆包|kimi|文心|通义", "生成程序直接是 AI 产品", "strong"),
]

TEXT_STRONG = [
    (r"utm_source=(chatgpt\.com|openai[\w.]*|claude[\w.]*|copilot[\w.]*|perplexity[\w.]*|gemini[\w.]*|deepseek[\w.]*|doubao[\w.]*|kimi[\w.]*|grok[\w.]*|you\.com)",
     "链接带 AI 产品的来源尾巴（从 AI 对话里原样复制的链接）"),
    (r":?contentReference\[oaicite:\d+\]|oaicite[:_]\d+|\{index=\d+\}", "ChatGPT 引用占位符残留"),
    (r"\bturn\d+(?:search|view|news|file|image|product|finance)\d+\b|\bfilecite\b|\uE200cite|\uE200|\uE201|\uE202", "ChatGPT 搜索引用标记残留"),
    (r"【\d+(?::\d+)?†[^】]{0,80}】", "ChatGPT 旧版引用标记【n†来源】残留"),
    (r"\[\^?\d+\^?\]\(https?://[^)]*\)\s*\[\^?\d+\^?\]\(", "连续的 Markdown 编号引用链接（AI 搜索回答的格式）"),
]
TEXT_MEDIUM = [
    (r"as an ai (language )?model|as a large language model|as an ai assistant", "AI 自述语句"),
    (r"as of my last (knowledge )?update|my knowledge cutoff|as of my knowledge cutoff|i (do not|don't) have (real-time|access to real-time)", "AI 知识截止声明"),
    (r"regenerate response|stop generating", "ChatGPT 界面按钮文字被一并复制"),
    (r"i hope this (helps|email finds)|let me know if you('d| would) like (me to|any)|would you like me to|feel free to (ask|let me know)|here('s| is) (a|an|the) (revised|rewritten|polished|updated|concise|shorter) (version|draft)|certainly[!,.] here|sure[!,.] here('s| is)|i cannot fulfill|i can't assist with",
     "AI 对话的开场白或收尾语"),
    (r"作为(一个|一名)?\s*(AI|人工智能)\s*(语言模型|助手|模型)|作为一个大语言模型", "AI 自述语句"),
    (r"截至我(的)?(知识|最后|训练)|我的知识截止|我无法(访问|获取)实时", "AI 知识截止声明"),
    (r"希望(这|以上|这些)[^。\n]{0,8}(对你|对您)有(所)?帮助|如果(你|您)(还)?(有|需要)[^。\n]{0,16}(请随时|随时告诉我|可以告诉我|我可以)|需要我[^。？\n]{0,20}吗[？?]|(当然|好的)(可以)?[！!，,]\s*(以下|下面|这是)|以下是(为您|为你|我)?[^。\n]{0,14}(版本|改写|润色|总结|草稿|优化)[：:]",
     "AI 对话的开场白或收尾语"),
    (r"\[(your name|your company|insert [^\]]{1,40}|company name|date|recipient'?s? name|您的姓名|你的名字|公司名称|插入[^\]]{1,20}|此处填写[^\]]{0,20}|日期|姓名)\]", "未填写的模板占位符"),
]
INVISIBLE = {
    "\u200b": "零宽空格 U+200B", "\u2060": "词连接符 U+2060", "\u202f": "窄不换行空格 U+202F",
    "\u00ad": "软连字符 U+00AD", "\ufeff": "零宽不换行空格 U+FEFF", "\u200c": "零宽不连字 U+200C",
}
AI_IMAGE_NAMES = re.compile(
    r"ChatGPT Image|DALL[·\-. ]?E|Gemini_Generated_Image|Midjourney|Stable ?Diffusion|Adobe Firefly|"
    r"Doubao|豆包|即梦|Flux[_ ]|Imagen|grok[-_ ]image|img-[A-Za-z0-9]{20,}", re.I)
SANDBOX_PATH = re.compile(
    r"(sandbox:/[^\s\"'<>)]*|/mnt/(?:data|user-data|outputs|skills|uploads)/[^\s\"'<>)]*|"
    r"/home/(?:claude|oai|sandbox|user|ubuntu|runner|appuser|jovyan|node|coder)/[^\s\"'<>)]*|"
    r"/(?:workspace|workdir|sessions)/[^\s\"'<>)]{3,})")
TMP_PATH = re.compile(r"(?<![\w.:/])/tmp/[\w\-./]{3,}")
USER_PATH = re.compile(
    r"([A-Za-z]:\\(?:Users|Documents and Settings)\\([^\\\"<>|\r\n]+)\\[^\"<>|\r\n]*)|"
    r"(/Users/([^/\"<>\s]+)/[^\"<>\s]*)|(/home/([^/\"<>\s]+)/[^\"<>\s]*)")
PDF_HEADERS = [
    (b"\xc4\xe5\xf2\xe5\xeb\xa7\xf3\xa0\xd0\xc4\xc6", "macOS Quartz"),
    (b"\xb5\xb5\xb5\xb5", "Microsoft Office"),
    (b"\x93\x8c\x8b\x9e", "ReportLab"),
    (b"\xbf\xf7\xa2\xfe", "qpdf"),
    (b"\xd0\xd4\xc5\xd8", "pdfTeX"),
    (b"\xc7\xec\x8f\xa2", "Ghostscript"),
    (b"\xe2\xe3\xcf\xd3", "Acrobat / iText / pypdf 等通用写法"),
]
AI_DOMAINS = re.compile(
    r"chatgpt\.com|chat\.openai\.com|oaiusercontent\.com|openai\.com|claude\.ai|claudeusercontent\.com|anthropic\.com|"
    r"gemini\.google\.com|aistudio\.google\.com|notebooklm\.google|copilot\.microsoft\.com|perplexity\.ai|doubao\.com|"
    r"deepseek\.com|kimi\.(?:com|ai)|moonshot\.cn|yiyan\.baidu\.com|tongyi\.(?:aliyun\.com|com)|chatglm\.cn|grok\.com|"
    r"poe\.com|manus\.im|gamma\.app|genspark\.ai", re.I)
AI_AGENTS = re.compile(r"chatgpt|claude|copilot|gemini|perplexity|doubao|豆包|deepseek|kimi|codex|cursor|comet|atlas", re.I)
C2PA_AI = re.compile(rb"trainedAlgorithmicMedia|compositeWithTrainedAlgorithmicMedia|OpenAI|ChatGPT|DALL|Anthropic|Claude|"
                     rb"Midjourney|Firefly|Stability|Google AI|Gemini|Imagen|Copilot|Designer|Bing Image", re.I)


# ---------------------------------------------------------------- 工具函数
def now():
    return dt.datetime.now(UTC)


def parse_iso(s):
    if not s:
        return None
    s = s.strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(Z|[+-]\d{2}:?\d{2})?", s)
    if not m:
        m2 = re.match(r"(\d{4})-(\d{2})-(\d{2})$", s)
        if not m2:
            return None
        return dt.datetime(int(m2[1]), int(m2[2]), int(m2[3]), tzinfo=UTC)
    try:
        d = dt.datetime(int(m[1]), int(m[2]), int(m[3]), int(m[4]), int(m[5]), int(m[6] or 0))
    except ValueError:
        return None
    tz = m[7]
    if tz and tz != "Z":
        sign = 1 if tz[0] == "+" else -1
        hh, mm = int(tz[1:3]), int(tz[-2:])
        d -= sign * dt.timedelta(hours=hh, minutes=mm)
    return d.replace(tzinfo=UTC)


def parse_pdf_date(s):
    if not s:
        return None
    m = re.match(r"(?:D:)?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?\s*([Zz+\-])?(\d{2})?'?(\d{2})?", s.strip())
    if not m:
        return parse_iso(s)
    try:
        d = dt.datetime(int(m[1]), int(m[2] or 1), int(m[3] or 1), int(m[4] or 0), int(m[5] or 0), int(m[6] or 0))
    except ValueError:
        return None
    if m[7] in ("+", "-"):
        sign = 1 if m[7] == "+" else -1
        d -= sign * dt.timedelta(hours=int(m[8] or 0), minutes=int(m[9] or 0))
    return d.replace(tzinfo=UTC)


def fmt(d):
    return d.strftime("%Y-%m-%d %H:%M:%S UTC") if d else "（无）"


def tag(xml, name):
    m = re.search(r"<(?:\w+:)?%s\b[^>]*?(?:/>|>(.*?)</(?:\w+:)?%s>)" % (name, name), xml, re.S)
    if not m or m.group(1) is None:
        return ""
    return html.unescape(re.sub(r"<[^>]+>", "", m.group(1))).strip()


def xml_text(xml):
    xml = re.sub(r"</(?:w:p|a:p|si|text:p)>|<(?:w:br|a:br|w:cr)\b[^>]*/>", "\n", xml)
    xml = re.sub(r"<(?:w:tab)\b[^>]*/>", "\t", xml)
    # 丢弃删除线修订、字段指令等非可见文字
    xml = re.sub(r"<w:instrText\b.*?</w:instrText>|<w:delText\b.*?</w:delText>", "", xml, flags=re.S)
    return html.unescape(re.sub(r"<[^>]+>", "", xml))


def count_words(text):
    cjk = len(re.findall(r"[\u3400-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]", text))
    latin = len(re.findall(r"[A-Za-z0-9][A-Za-z0-9'’\-]*", text))
    return cjk + latin


def clip(s, n=160):
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s if len(s) <= n else s[: n - 1] + "…"


class Report:
    def __init__(self, path):
        self.path = path
        self.kind = ""
        self.findings = []
        self.facts = {}        # 展示用的基本属性
        self.exposure = []     # 自查：对外会暴露的信息
        self.notes = []        # 未能检查的内容
        self.fp = {}           # 指纹：用于多文件比对和文件外痕迹比对
        self._seen = set()

    def add(self, cat, level, title, evidence="", innocent="", where=""):
        key = (cat, title, clip(evidence, 80))
        if key in self._seen:
            return
        self._seen.add(key)
        self.findings.append({"category": cat, "category_zh": CATS[cat], "level": level,
                              "title": title, "evidence": clip(evidence, 400),
                              "innocent": innocent, "where": where})

    def expose(self, what, value, where=""):
        item = (what, clip(value, 200), where)
        if item not in self.exposure:
            self.exposure.append(item)

    def counts(self):
        c = {k: 0 for k in LEVELS}
        for f in self.findings:
            c[f["level"]] += 1
        return c

    def verdict(self):
        c = self.counts()
        if c["strong"]:
            return "A", "发现直接痕迹：文件由程序或 AI 工具生成或改写，或含有从 AI 对话带出的内容"
        if c["medium"] >= 3:
            return "B", "多项间接疑点叠加，建议向提交人核实原始文件和写作过程"
        if c["medium"]:
            return "C", "有少量疑点，各自都有正常解释，不足以下结论"
        return "D", "未发现明显痕迹（不代表文字不是 AI 写的：把 AI 文字手动敲进或粘进正常文档不会留下这类痕迹）"

    def to_dict(self):
        code, text = self.verdict()
        return {"file": self.path, "kind": self.kind, "verdict_code": code, "verdict": text,
                "counts": self.counts(), "facts": self.facts,
                "findings": sorted(self.findings, key=lambda f: LEVELS.index(f["level"])),
                "exposure": [{"what": a, "value": b, "where": c} for a, b, c in self.exposure],
                "not_checked": self.notes}


# ---------------------------------------------------------------- 通用检查
def check_name(R, value, field, where):
    """作者类字段：库名、AI 产品名、通用名。"""
    if not value:
        return
    v = value.strip()
    for rx, desc, lvl in GENERATOR_NAMES:
        if re.search(rx, v, re.I):
            R.add("author", lvl, f"{field}是「{v}」：{desc}", f"{field} = {v}",
                  "有人把自己的用户名设成这类名字的可能性很小" if lvl == "strong" else "企业报表系统、模板平台也用这些库批量出文件", where)
            return
    if GENERIC_NAMES.match(v):
        R.add("author", "weak", f"{field}是通用名「{v}」，无法对应到具体的人", f"{field} = {v}",
              "未登录或未激活的 Office、公用电脑、重装系统后的默认用户名都会这样", where)


def check_expected(R, names, orgs, opts):
    ea, eo = opts.get("expect_author"), opts.get("expect_org")

    def norm(s):
        return re.sub(r"[\s·.\-_,]", "", s or "").lower()
    if ea:
        got = [n for n in names if n]
        if got and not any(norm(ea) in norm(n) or norm(n) in norm(ea) for n in got):
            R.add("author", "medium", f"文件里的作者信息与预期作者「{ea}」对不上", "文件中出现的名字: " + "、".join(sorted(set(got))),
                  "用了别人的电脑或模板、公司统一用户名、代为排版；也可能是 AI 编造的作者名", "docProps")
        elif not got:
            R.add("author", "weak", f"文件里没有任何作者信息，无法与预期作者「{ea}」核对", "",
                  "用文档检查器清过元数据、由在线文档导出", "docProps")
    if eo:
        got = [o for o in orgs if o]
        if got and not any(norm(eo) in norm(o) or norm(o) in norm(eo) for o in got):
            R.add("author", "medium", f"文件里的机构信息与预期机构「{eo}」对不上", "文件中出现的机构: " + "、".join(sorted(set(got))),
                  "沿用了前雇主或外部模板；也可能是 AI 编造的机构名", "docProps")


def check_core_times(R, created, modified, raw_created, raw_modified, where):
    n = now()
    for raw in (raw_created, raw_modified):
        if raw in LIB_DATES:
            R.add("time", "strong", f"时间戳是库模板的固定值：{LIB_DATES[raw]}", f"{raw}",
                  "这个时间是库自带模板里写死的，人工保存不会碰巧得到", where)
    if created and modified and created - modified > dt.timedelta(minutes=2):
        R.add("time", "medium", "创建时间晚于最后修改时间", f"创建 {fmt(created)}；修改 {fmt(modified)}",
              "从旧文件复制内容到新文件、跨时区或系统时钟错误偶尔会造成", where)
    for label, d in (("创建", created), ("修改", modified)):
        if d and d - n > TOL:
            R.add("time", "medium", f"{label}时间在未来", fmt(d), "电脑时钟设置错误；也可能是 AI 编造的日期", where)
        if d and d.year < 1995:
            R.add("time", "medium", f"{label}时间早得不合理", fmt(d), "程序写入了默认值或空值", where)
    for label, d, raw in (("创建", created, raw_created), ("修改", modified, raw_modified)):
        if d and raw and d.hour == 0 and d.minute == 0 and d.second == 0 and "T" in raw:
            R.add("time", "weak", f"{label}时间恰好是零点整", raw, "程序只写了日期没写时间；人工保存很少恰好落在零点整", where)


def _far(a, b):
    """两处批注的位置是否相距很远（不同页，或同页纵向相差超过约 3 行）。"""
    if a is None or b is None:
        return True
    return a[0] != b[0] or abs(a[1] - b[1]) > 40


def check_events(R, events, label, created, modified, where, minute_precision_expected=False, mass_ok=False, positions=None):
    """events: [(作者, datetime 或 None, 原始时间字符串)]"""
    if not events:
        return
    n = now()
    dated = [(a, d, r) for a, d, r in events if d]
    undated = len(events) - len(dated)
    if undated:
        R.add("review", "weak", f"{undated} 条{label}没有时间", f"共 {len(events)} 条，其中 {undated} 条缺时间",
              "部分软件（旧版 WPS、在线文档导出）不写时间；程序插入的批注也常缺", where)
    early = [e for e in dated if created and created.year > 2013 and created - e[1] > TOL]
    if early:
        e = min(early, key=lambda x: x[1])
        R.add("review", "medium", f"{len(early)} 条{label}早于文档创建时间", f"最早一条 {fmt(e[1])}（{e[0] or '无名'}）；文档创建于 {fmt(created)}",
              "整段内容连同批注从旧文件复制过来、文件被另存为新文件", where)
    late = [e for e in dated if modified and e[1] - modified > TOL]
    if late:
        e = max(late, key=lambda x: x[1])
        R.add("review", "medium", f"{len(late)} 条{label}晚于文档最后保存时间", f"最晚一条 {fmt(e[1])}（{e[0] or '无名'}）；文档最后保存于 {fmt(modified)}",
              "保存它的程序没有更新修改时间；也可能批注时间是编造的", where)
    fut = [e for e in dated if e[1] - n > TOL]
    if fut:
        R.add("review", "medium", f"{len(fut)} 条{label}的时间在未来", fmt(max(x[1] for x in fut)), "电脑时钟错误；也可能是编造的时间", where)
    if len(dated) >= 3:
        times = sorted(d for _, d, _ in dated)
        if times[0] == times[-1] and not mass_ok:
            R.add("review", "medium", f"全部 {len(dated)} 条{label}的时间完全相同", fmt(times[0]),
                  "Word 的「比较/合并文档」功能会这样；否则说明是程序一次性写入的，不是人逐条审阅", where)
        elif mass_ok and len(dated) >= 20 and times[0] == times[-1]:
            R.add("review", "medium", f"全部 {len(dated)} 处{label}的时间完全相同", fmt(times[0]),
                  "Word 的「比较文档」功能会这样；否则说明修订是程序一次性生成的", where)
        subsec = any(d.second != 0 for d in times)
        if subsec and times[0] != times[-1]:
            pts = sorted(((e[1], positions[i] if positions else None) for i, e in enumerate(events) if e[1]), key=lambda x: x[0])
            clusters = clustered = 0
            i = 0
            while i < len(pts):
                j = i
                while j + 1 < len(pts) and (pts[j + 1][0] - pts[j][0]).total_seconds() <= 1:
                    j += 1
                if j > i and any(_far(pts[k][1], pts[k + 1][1]) for k in range(i, j)):
                    clusters += 1
                    clustered += j - i + 1
                i = j + 1
            span = int((times[-1] - times[0]).total_seconds())
            if clusters >= 3 and clustered >= len(pts) * 0.4:
                R.add("review", "medium", f"多处{label}在同一秒内成批出现，不像人工逐条操作",
                      f"{len(pts)} 条中有 {clustered} 条分成 {clusters} 批，每批落在同一两秒内"
                      + ("，同批之间位置相距很远" if positions else "") + f"；全部{label}的时间跨度 {span} 秒",
                      "脚本、宏或自动化工具批量标注；人手逐条点选做不到同一秒标出相隔很远的几处", where)
        if all(d.minute == 0 and d.second == 0 for d in times) and len(set(times)) >= 2:
            R.add("review", "medium", f"{label}的时间全是整点", "、".join(fmt(t) for t in times[:4]),
                  "人工审阅很难每条都落在整点，像是编造的时间", where)
    if minute_precision_expected and dated and any(d.second != 0 for _, d, _ in dated):
        R.add("review", "weak", f"{label}时间精确到秒，而 Word 自己只写到分钟", clip(next(r for _, d, r in dated if d.second != 0)),
              "WPS、在线文档导出、第三方审阅工具会写到秒；程序插入的批注也是", where)
    for a in sorted({a for a, _, _ in events if a}):
        hit = False
        for rx, desc, lvl in GENERATOR_NAMES:
            if re.search(rx, a, re.I):
                R.add("review", lvl, f"{label}的署名是「{a}」：{desc}", "", "审阅人的名字被设成这类名字的可能性很小", where)
                hit = True
                break
        if not hit and GENERIC_NAMES.match(a.strip()):
            R.add("review", "weak", f"{label}的署名是通用名「{a}」", "", "未登录的 Office、公用电脑；程序插入批注时也常用这类占位名", where)


def scan_text(R, text, where, strong_only=False):
    """正文里的 AI 对话残留。"""
    if not text:
        return
    low = text
    for rx, desc in TEXT_STRONG:
        ms = list(re.finditer(rx, low, re.I))
        if ms:
            R.add("text", "strong", desc, f"{len(ms)} 处，例：{clip(low[max(0, ms[0].start()-30): ms[0].end()+20], 140)}",
                  "只能说明这段内容或链接经过了 AI 工具，不能说明全文由 AI 撰写", where)
    if strong_only:
        return
    for rx, desc in TEXT_MEDIUM:
        ms = list(re.finditer(rx, low, re.I))
        if ms:
            R.add("text", "medium", desc, f"{len(ms)} 处，例：{clip(low[max(0, ms[0].start()-20): ms[0].end()+30], 140)}",
                  "讨论 AI 的文章会引用这类句子；正式信函里也有人这样写，需看上下文", where)
    inv = {name: low.count(ch) for ch, name in INVISIBLE.items() if low.count(ch)}
    # 文件开头的 BOM 不算
    if inv:
        total = sum(inv.values())
        R.add("text", "weak", "正文含隐形或特殊空白字符", "；".join(f"{k} ×{v}" for k, v in inv.items()),
              "法语排版、网页复制、部分输入法和排版软件都会带入；多家分析认为它不是水印" + ("" if total >= 5 else "，数量很少"), where)
    md = len(re.findall(r"(?m)^\s{0,3}#{1,6}\s+\S", low)) + len(re.findall(r"\*\*[^*\n]{1,80}\*\*", low)) \
        + len(re.findall(r"(?m)^\s*\|?\s*:?-{3,}:?\s*\|", low)) + low.count("```")
    if md >= 3:
        ex = re.search(r"(?m)^\s{0,3}#{1,6}\s+\S.{0,40}|\*\*[^*\n]{1,60}\*\*", low)
        R.add("text", "medium" if md >= 8 else "weak", "正文残留 Markdown 符号（#、**、表格线）",
              f"{md} 处，例：{clip(ex.group(0)) if ex else ''}",
              "技术文档作者本来就写 Markdown；否则多半是从 AI 对话框直接粘贴的", where)


def scan_paths(R, raw, where):
    raw = html.unescape(raw)
    for m in SANDBOX_PATH.finditer(raw):
        R.add("path", "strong", "文件里残留 AI 沙盒或代码运行环境的路径", m.group(0),
              "这类路径只存在于 AI 的代码执行环境或云端容器，不会出现在个人电脑上", where)
        R.expose("沙盒路径", m.group(0), where)
    for m in TMP_PATH.finditer(raw):
        R.add("path", "medium", "文件里残留临时目录路径 /tmp/…", m.group(0), "Mac/Linux 上的脚本或转换工具也会用临时目录", where)
        R.expose("临时路径", m.group(0), where)
    for m in USER_PATH.finditer(raw):
        full = m.group(1) or m.group(3) or m.group(5)
        user = m.group(2) or m.group(4) or m.group(6)
        if SANDBOX_PATH.match(full):
            continue
        R.expose("本机路径（含用户名「%s」）" % user, full, where)
        R.add("path", "info", f"文件里记录了本机路径，用户名是「{user}」", full, "请与作者栏、提交人对照：不一致说明文件出自另一台电脑", where)
        R.facts.setdefault("路径中的用户名", set()).add(user)


def scan_binary_labels(R, data, where, is_image=False):
    """C2PA 内容凭证、AIGC 隐式标识、AI 绘图工具写入的元数据。"""
    if b"c2pa" in data or b"jumbf" in data.lower()[:200000] or b"contentauth" in data:
        ai = C2PA_AI.search(data)
        if ai:
            R.add("label", "strong", "含 C2PA 内容凭证，且凭证指向 AI 生成", "凭证内出现: " + ai.group(0).decode("latin1"),
                  "凭证带签名，是厂商主动打的标；它说明这个文件或图片经过 AI 工具，不说明全文都是", where)
        else:
            R.add("label", "medium", "含 C2PA 内容凭证（来源签名）", "检测到 c2pa / JUMBF 标记，未读出生成方",
                  "相机、Photoshop 也会写 C2PA；需用 contentcredentials.org/verify 查看具体来源", where)
    m = re.search(rb"(?:AIGC[\"']?\s*[:=>]\s*[\"']?\{|[\"']?(?:ContentProducer|ProduceID|ContentPropagator|PropagateID)[\"']?\s*[:=])[^\x00]{0,300}", data)
    if m:
        R.add("label", "strong", "含 AIGC 隐式标识（国标 GB 45438-2025 要求 AI 生成内容写入的元数据）",
              m.group(0).decode("utf-8", "replace"), "国内 AI 服务按规定写入，属于厂商主动声明", where)
    if is_image:
        m = re.search(rb"(DALL.?E|Midjourney|Stable Diffusion|ComfyUI|NovelAI|Adobe Firefly|Made with Google AI|"
                      rb"sd-metadata|Negative prompt:|trainedAlgorithmicMedia)", data, re.I)
        if m:
            R.add("label", "strong", "内嵌图片的元数据标明由 AI 绘图工具生成", m.group(0).decode("latin1"),
                  "只说明这张图是 AI 生成的", where)


# ---------------------------------------------------------------- OOXML (docx / pptx / xlsx)
def read(z, name, limit=30_000_000):
    try:
        info = z.getinfo(name)
        if info.file_size > limit:
            return ""
        return z.read(name).decode("utf-8", "replace")
    except KeyError:
        return ""


def scan_ooxml(path, kind, R, opts, depth=0):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        R.notes.append("文件不是有效的压缩包结构（可能已加密、损坏，或其实是旧版 .doc/.xls/.ppt 改了扩展名）")
        return
    names = z.namelist()
    infos = z.infolist()
    R.kind = {"docx": "Word 文档", "pptx": "PowerPoint 演示文稿", "xlsx": "Excel 工作簿"}[kind]

    core = read(z, "docProps/core.xml")
    app = read(z, "docProps/app.xml")
    custom = read(z, "docProps/custom.xml")

    creator, last_by = tag(core, "creator"), tag(core, "lastModifiedBy")
    raw_c, raw_m = tag(core, "created"), tag(core, "modified")
    created, modified = parse_iso(raw_c), parse_iso(raw_m)
    revision, descr, title = tag(core, "revision"), tag(core, "description"), tag(core, "title")
    last_printed = parse_iso(tag(core, "lastPrinted"))
    application, app_ver = tag(app, "Application"), tag(app, "AppVersion")
    company, template, manager = tag(app, "Company"), tag(app, "Template"), tag(app, "Manager")
    total_time, app_words = tag(app, "TotalTime"), tag(app, "Words")

    R.fp.update(created=created, modified=modified)
    try:
        with open(path, "rb") as fh:
            R.fp["sha"] = hashlib.sha256(fh.read()).hexdigest()
    except OSError:
        pass
    F = R.facts
    F["作者"] = creator or "（空）"
    F["最后保存者"] = last_by or "（空）"
    F["创建时间"] = fmt(created) if created else (raw_c or "（无）")
    F["最后修改时间"] = fmt(modified) if modified else (raw_m or "（无）")
    if revision:
        F["修订号"] = revision
    if application:
        F["保存程序"] = application + (f" {app_ver}" if app_ver else "")
    if company:
        F["公司"] = company
    if template:
        F["模板"] = template
    if total_time:
        F["累计编辑时间(分钟)"] = total_time
    if last_printed:
        F["上次打印"] = fmt(last_printed)

    for label, v in (("作者", creator), ("最后保存者", last_by), ("公司", company), ("经理", manager), ("模板", template), ("标题", title)):
        if v:
            R.expose(label, v, "文档属性")

    # ---- 1 作者
    check_name(R, creator, "作者", "docProps/core.xml")
    if last_by != creator:
        check_name(R, last_by, "最后保存者", "docProps/core.xml")
    if not core:
        R.add("struct", "medium", "缺少核心属性文件（作者、时间全无）", "docProps/core.xml 不存在",
              "部分在线文档导出、精简的生成库不写这个文件；Office 保存的文件一定有", "docProps")
    elif not creator and not last_by:
        R.add("author", "weak", "作者和最后保存者都是空的", "", "用「检查文档」清除过个人信息、在线文档导出、部分生成库都会这样", "docProps/core.xml")
    if descr and re.search(r"generated by|created (with|by)|python|由.{0,6}生成", descr, re.I):
        R.add("author", "strong" if re.search(r"python-(docx|pptx)|openpyxl|chatgpt|openai|claude|\bAI\b", descr, re.I) else "medium",
              "文档「备注」栏写着生成来源", descr, "生成库自动填写的说明", "docProps/core.xml")
    check_expected(R, [creator, last_by], [company], opts)

    # ---- 3 时间
    check_core_times(R, created, modified, raw_c, raw_m, "docProps/core.xml")
    if raw_m in LIB_DATES or raw_c in LIB_DATES:
        # 属性时间是库模板的旧值，不能再拿来和批注时间比较
        ev_created, ev_modified = None, None
    else:
        ev_created, ev_modified = created, modified
    if core and not raw_c:
        R.add("time", "weak", "没有创建时间", "", "在线文档导出、部分生成库不写创建时间", "docProps/core.xml")
    if last_printed and created and created - last_printed > TOL:
        R.add("time", "info", "上次打印时间早于创建时间", f"打印 {fmt(last_printed)}；创建 {fmt(created)}",
              "说明文件是从更早的文件另存或套用而来，很常见", "docProps/core.xml")

    # ---- 正文提取
    if kind == "docx":
        body_parts = [n for n in names if re.match(r"word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$", n)]
    elif kind == "pptx":
        body_parts = [n for n in names if re.match(r"ppt/(slides|notesSlides)/[^/]+\.xml$", n)]
    else:
        body_parts = [n for n in names if n == "xl/sharedStrings.xml"]
    body_xml = {n: read(z, n) for n in body_parts}
    body_text = "\n".join(xml_text(x) for x in body_xml.values())
    words = count_words(body_text)
    F["实际字数（估算）"] = words

    # ---- 5 结构：保存程序与统计
    if not app:
        R.add("struct", "medium", "缺少应用属性文件（保存程序、编辑时间、字数统计全无）", "docProps/app.xml 不存在",
              "Office 保存的文件一定有；docx(JS) 等生成库、部分在线导出不写", "docProps")
    else:
        hit = False
        for rx, desc, lvl in APP_NAMES:
            if application and re.search(rx, application, re.I):
                R.add("struct", lvl, f"保存程序是「{application}」：{desc}", "", "企业系统批量生成的报表也会这样", "docProps/app.xml")
                hit = True
        if application and not hit and not NORMAL_APPS.search(application):
            R.add("struct", "weak", f"保存程序「{application}」不是常见办公软件", "", "小众办公软件或转换工具", "docProps/app.xml")
        if not application:
            R.add("struct", "weak", "应用属性里没有写保存程序", "", "部分生成库和在线导出不写程序名", "docProps/app.xml")
        if total_time.isdigit():
            tt = int(total_time)
            if kind == "docx" and tt == 0 and words >= 300:
                R.add("time", "medium", f"累计编辑时间为 0 分钟，但正文约 {words} 字", f"TotalTime = 0",
                      "「另存为」新文件会清零；从别处整体粘贴后立即保存也是 0；它只统计文件打开的时长", "docProps/app.xml")
            elif kind == "docx" and tt > 0 and words >= 800 and words / tt > 120:
                R.add("time", "weak", f"编辑时间与篇幅不匹配：{tt} 分钟写了约 {words} 字", f"约每分钟 {words // tt} 字",
                      "内容是从别的文件粘贴来的（不一定是 AI）、另存后继续编辑", "docProps/app.xml")
        if kind == "docx" and app_words.isdigit():
            aw = int(app_words)
            if words >= 200 and (aw == 0 or aw < words / 4 or aw > words * 4):
                R.add("struct", "medium", "文件里缓存的字数统计与实际正文对不上", f"属性里记的是 {aw} 字，实际约 {words} 字",
                      "Word 每次保存都会更新这个数；对不上说明最后一次写入正文的不是 Word（python-docx 会原样保留模板里的旧统计）", "docProps/app.xml")
        if creator and re.match(r"^administrator$", creator.strip(), re.I) and total_time == "0" and "TitlesOfParts" not in app:
            R.add("struct", "medium", "作者 Administrator + 零编辑时间 + 缺少部件标题", "这个组合被邮件安全厂商用作程序生成文件的检测规则",
                  "未激活 Office 的电脑另存文件", "docProps")

    # ---- 5 结构：压缩包层
    real = [i for i in infos if i.date_time[0] > 1980]
    systems = {i.create_system for i in infos}
    if real:
        zt = sorted(dt.datetime(*i.date_time) for i in real)
        F["压缩包内部时间戳"] = f"{zt[0]:%Y-%m-%d %H:%M:%S} ~ {zt[-1]:%Y-%m-%d %H:%M:%S}（打包机器的本地时间）"
        R.add("struct", "medium" if 3 in systems else "weak",
              "压缩包内部带真实时间戳" + ("，且由 Unix/Linux/macOS 上的非 Office 程序打包" if 3 in systems else ""),
              f"{zt[-1]:%Y-%m-%d %H:%M:%S}；Microsoft Office 保存的文件这里固定是 1980-01-01",
              "WPS、LibreOffice、在线文档导出、解压后重新压缩都会留下真实时间；代码生成的文件也是", "ZIP 目录")
        R.expose("真实打包时间", f"{zt[-1]:%Y-%m-%d %H:%M:%S}（本地时间，无法通过改文档属性抹掉）", "ZIP 目录")
        zlast = zt[-1].replace(tzinfo=UTC)
        if modified and abs(zlast - modified) > dt.timedelta(hours=26):
            R.add("time", "medium", "文档属性里的修改时间与实际打包时间不一致",
                  f"属性写的是 {fmt(modified)}；压缩包实际生成于 {zt[-1]:%Y-%m-%d %H:%M:%S}",
                  "生成库不更新属性时间（沿用模板的旧日期），或属性时间是编造的；压缩包时间更可信", "ZIP 目录")
        if not modified and not created:
            R.add("time", "info", "文档属性没有时间，但压缩包记录了实际生成时间", f"{zt[-1]:%Y-%m-%d %H:%M:%S}", "", "ZIP 目录")
    else:
        F["压缩包内部时间戳"] = "1980-01-01（Microsoft Office 的默认写法）"
    if names and names[0] != "[Content_Types].xml":
        R.add("struct", "weak", "压缩包内文件顺序不是 Office 的写法", f"第一个条目是 {names[0]}（Office 固定先写 [Content_Types].xml）",
              "WPS、LibreOffice、重新压缩过的文件", "ZIP 目录")

    # ---- 7 官方标识：自定义属性
    if custom:
        props = re.findall(r'<property\b[^>]*\bname="([^"]*)"[^>]*>(.*?)</property>', custom, re.S)
        for pname, pval in props:
            val = html.unescape(re.sub(r"<[^>]+>", "", pval)).strip()
            if re.search(r"AIGC", pname) or re.search(r"AIGC", val):
                R.add("label", "strong", "自定义属性里有 AIGC 隐式标识（国标要求 AI 生成内容写入）", f"{pname} = {val}",
                      "国内 AI 服务按规定写入，属于厂商主动声明", "docProps/custom.xml")
            elif re.search(r"chatgpt|openai|claude|anthropic|copilot|gemini|deepseek|doubao|kimi|ai[-_ ]?generated|generated[-_ ]?by", pname + " " + val, re.I):
                R.add("label", "strong", "自定义属性提到 AI 工具或生成来源", f"{pname} = {val}", "", "docProps/custom.xml")
            elif pname.startswith("MSIP_Label") and pname.endswith("_SiteId"):
                R.add("label", "info", "带企业敏感度标签，记录了所属组织的租户 ID", f"{pname} = {val}",
                      "可与提交人自称的机构核对：标签来自别的组织说明文件出自那里", "docProps/custom.xml")
                R.expose("敏感度标签的租户 ID", val, "docProps/custom.xml")
            elif pname.startswith("MSIP_Label") and pname.endswith("_Name"):
                R.expose("敏感度标签名称", val, "docProps/custom.xml")
            elif not pname.startswith("MSIP_Label"):
                R.expose(f"自定义属性 {pname}", val, "docProps/custom.xml")
    for part in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
        raw = read(z, part)
        m = re.search(r"[^<>]{0,60}(intelligence data|ai[- ]assisted|copilot)[^<>]{0,60}", raw, re.I)
        if m:
            R.add("label", "medium", "文档属性里出现 AI 辅助功能相关的字样", m.group(0),
                  "Office 的智能功能（Copilot、编辑器）会写入；说明用过 AI 辅助，不说明全文由 AI 写", part)

    # ---- 2 路径 + 6 正文残留 + 7 二进制标识：遍历所有部件
    emb = []
    for info in infos:
        n = info.filename
        low = n.lower()
        if info.file_size > 40_000_000:
            R.notes.append(f"{n} 过大，未扫描")
            continue
        if low.endswith((".xml", ".rels", ".vml")):
            raw = read(z, n)
            scan_paths(R, raw, n)
            if low.endswith(".rels"):
                scan_text(R, html.unescape(raw), n, strong_only=True)
                for t in re.findall(r'Type="[^"]*attachedTemplate"[^>]*Target="([^"]*)"|Target="([^"]*)"[^>]*Type="[^"]*attachedTemplate"', raw):
                    tgt = html.unescape(t[0] or t[1])
                    R.expose("挂接的模板路径", tgt, n)
                    F["挂接模板"] = tgt
            m = AI_IMAGE_NAMES.search(raw)
            if m and (n in body_xml or "drawing" in low or "slide" in low):
                ctx = raw[max(0, m.start() - 40): m.end() + 60]
                R.add("embed", "strong", "插入的图片保留着 AI 绘图工具的原始文件名", clip(html.unescape(re.sub(r"<[^>]+>", " ", ctx)), 140),
                      "只说明这张图来自 AI 工具", n)
            if n.startswith("docProps/") and "AIGC" in raw and "custom" not in n:
                R.add("label", "strong", "文档属性里有 AIGC 标识", clip(re.search(r".{0,40}AIGC.{0,80}", raw, re.S).group(0)), "", n)
        elif "/media/" in low or low.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".tif", ".tiff", ".bmp", ".emf", ".wmf")):
            data = z.read(n)
            scan_binary_labels(R, data, n, is_image=True)
            m = re.search(rb"(?:Software|CreatorTool)[\x00-\x20:=\"'>]{1,12}([\x20-\x7e]{3,60})", data[:300000])
            if m:
                R.expose("内嵌图片的制作软件", m.group(1).decode("latin1"), n)
        elif "/embeddings/" in low or low.endswith((".bin", ".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt", ".pdf")) and "/" in n:
            if "/embeddings/" in low:
                emb.append(n)

    scan_text(R, body_text, "正文")

    # ---- 8 嵌入对象
    for n in emb:
        data = z.read(n)
        R.expose("嵌入对象", n, "embeddings")
        if data[:2] == b"PK" and depth < 2:
            try:
                sub = zipfile.ZipFile(io.BytesIO(data))
                sc = sub.read("docProps/core.xml").decode("utf-8", "replace") if "docProps/core.xml" in sub.namelist() else ""
                s_creator, s_last = tag(sc, "creator"), tag(sc, "lastModifiedBy")
                s_created = parse_iso(tag(sc, "created"))
                R.add("embed", "info", f"嵌入了一个 Office 文件，它自带独立的元数据", f"{n}：作者 {s_creator or '空'}，最后保存者 {s_last or '空'}，创建 {fmt(s_created)}",
                      "嵌入对象的元数据不会被「检查文档」清除，可与外层作者对照", n)
                check_name(R, s_creator, "嵌入文件的作者", n)
                if s_last != s_creator:
                    check_name(R, s_last, "嵌入文件的最后保存者", n)
                for v in (s_creator, s_last):
                    if v:
                        R.expose("嵌入文件里的名字", v, n)
            except Exception:
                pass
        else:
            R.add("embed", "info", "含嵌入对象（旧格式或其他类型），未深入解析", n, "嵌入对象可能带自己的作者和路径信息", n)
            scan_paths(R, data.decode("utf-16-le", "ignore") + data.decode("latin1", "ignore"), n)

    # ---- 各格式专项
    if kind == "docx":
        docx_specific(z, names, R, body_xml, words, ev_created, ev_modified, creator, last_by)
    elif kind == "pptx":
        pptx_specific(z, names, R, ev_created, ev_modified)
    else:
        xlsx_specific(z, names, R, ev_created, ev_modified)
    z.close()


def docx_specific(z, names, R, body_xml, words, created, modified, creator, last_by):
    settings = read(z, "word/settings.xml")
    doc = body_xml.get("word/document.xml", "")
    rsids = set(re.findall(r'<w:rsid(?:Root)?\s+w:val="([0-9A-Fa-f]+)"', settings))
    in_doc = set(re.findall(r'w:rsid(?:R|RPr|P|Del|RDefault|Sect|Tr)="([0-9A-Fa-f]+)"', doc))
    R.facts["编辑会话标识(rsid)数"] = len(rsids | in_doc)
    if words >= 200 and not rsids and not in_doc:
        R.add("struct", "medium", "全文没有任何编辑会话标识（rsid）", f"约 {words} 字，rsid 数为 0",
              "Word 每次编辑保存都会写入 rsid；WPS、Google 文档导出、LibreOffice 不写，关闭了「保存时存储随机数」的 Word 也不写；代码生成的文件一定没有", "word/settings.xml")
    elif words >= 1500 and len(rsids | in_doc) <= 2:
        R.add("struct", "weak", f"约 {words} 字的文档只有 {len(rsids | in_doc)} 个编辑会话标识", "",
              "说明内容基本是一次性写入或粘贴进来的（来源不一定是 AI）", "word/settings.xml")

    # 批注
    comments = read(z, "word/comments.xml")
    ev = []
    for m in re.finditer(r"<w:comment\b([^>]*)>", comments):
        a = re.search(r'w:author="([^"]*)"', m.group(1))
        d = re.search(r'w:date="([^"]*)"', m.group(1))
        ev.append((html.unescape(a.group(1)) if a else "", parse_iso(d.group(1)) if d else None, d.group(1) if d else ""))
    if len(ev) >= 2:
        bodies = [hashlib.md5(xml_text(b).strip().encode("utf-8")).hexdigest()[:10]
                  for b in re.findall(r"<w:comment\b[^>]*>(.*?)</w:comment>", comments, re.S)]
        if len(bodies) == len(ev):
            R.fp["geom"] = tuple(sorted((0, "comment", (), ev[i][0] + ":" + bodies[i]) for i in range(len(ev))))
            R.fp["times"] = tuple(sorted(r for _, _, r in ev))
    if ev:
        authors = sorted({a for a, _, _ in ev if a})
        R.facts["批注"] = f"{len(ev)} 条，署名：{'、'.join(authors) or '无'}"
        ds = sorted(d for _, d, _ in ev if d)
        if ds:
            R.facts["批注时间范围"] = f"{fmt(ds[0])} ~ {fmt(ds[-1])}"
        for a in authors:
            R.expose("批注人", a, "word/comments.xml")
        R.expose("批注数量", f"{len(ev)} 条（对外发送前确认是否要保留）", "word/comments.xml")
        check_events(R, ev, "批注", created, modified, "word/comments.xml", minute_precision_expected=True)
        if "word/people.xml" not in names and authors:
            R.add("review", "weak", "有批注但没有审阅人名单文件", "word/people.xml 不存在",
                  "Word 2013 及以后添加批注时会生成这个文件；旧版 Word、WPS 不生成，程序插入的批注通常也没有", "word/")
        ct = xml_text(comments)
        scan_text(R, ct, "批注内容")
    # 修订
    tv = []
    for m in re.finditer(r"<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b([^>]*)>", doc):
        a = re.search(r'w:author="([^"]*)"', m.group(2))
        d = re.search(r'w:date="([^"]*)"', m.group(2))
        if a or d:
            tv.append((html.unescape(a.group(1)) if a else "", parse_iso(d.group(1)) if d else None, d.group(1) if d else ""))
    if tv:
        authors = sorted({a for a, _, _ in tv if a})
        R.facts["修订记录"] = f"{len(tv)} 处，署名：{'、'.join(authors) or '无'}"
        for a in authors:
            R.expose("修订人", a, "word/document.xml")
        R.expose("未接受的修订", f"{len(tv)} 处（收件人可以看到删改前的原文）", "word/document.xml")
        check_events(R, tv, "修订", created, modified, "word/document.xml", minute_precision_expected=True, mass_ok=True)
    hidden = len(re.findall(r"<w:vanish\b(?![^>]*w:val=\"(?:0|false)\")", doc))
    if hidden:
        R.expose("隐藏文字", f"{hidden} 处", "word/document.xml")
        R.add("embed", "info", f"正文含 {hidden} 处隐藏文字", "", "隐藏文字对外发送后仍可被显示出来", "word/document.xml")


def pptx_specific(z, names, R, created, modified):
    pres = read(z, "ppt/presentation.xml")
    slides = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    notes = [n for n in names if re.match(r"ppt/notesSlides/", n) and n.endswith(".xml")]
    R.facts["幻灯片数"] = len(slides)
    hidden = [n for n in slides if re.search(r'<p:sld\b[^>]*\bshow="0"', read(z, n)[:2000])]
    if hidden:
        R.expose("隐藏的幻灯片", f"{len(hidden)} 页", "ppt/slides")
    note_text = "\n".join(xml_text(read(z, n)) for n in notes)
    note_words = count_words(re.sub(r"\b\d+\b", "", note_text))
    if note_words > 5:
        R.expose("演讲者备注", f"约 {note_words} 字（直接发 pptx 时会一并带出）", "ppt/notesSlides")
    authors = {}
    ca = read(z, "ppt/commentAuthors.xml")
    for m in re.finditer(r"<p:cmAuthor\b([^>]*)/?>", ca):
        i = re.search(r'\bid="([^"]*)"', m.group(1))
        nm = re.search(r'\bname="([^"]*)"', m.group(1))
        if i and nm:
            authors[i.group(1)] = html.unescape(nm.group(1))
    for n in names:
        if n.startswith("ppt/authors"):
            for m in re.finditer(r"<p\d*:author\b([^>]*)/?>", read(z, n)):
                i = re.search(r'\bid="([^"]*)"', m.group(1))
                nm = re.search(r'\bname="([^"]*)"', m.group(1))
                if i and nm:
                    authors[i.group(1)] = html.unescape(nm.group(1))
    ev = []
    for n in names:
        if re.match(r"ppt/comments/.*\.xml$", n):
            raw = read(z, n)
            for m in re.finditer(r"<p\d*:cm\b([^>]*)>", raw):
                aid = re.search(r'\bauthorId="([^"]*)"', m.group(1))
                d = re.search(r'\b(?:dt|created)="([^"]*)"', m.group(1))
                ev.append((authors.get(aid.group(1), "") if aid else "", parse_iso(d.group(1)) if d else None, d.group(1) if d else ""))
            scan_text(R, xml_text(raw), "批注内容")
    if ev:
        an = sorted({a for a, _, _ in ev if a})
        R.facts["批注"] = f"{len(ev)} 条，署名：{'、'.join(an) or '无'}"
        for a in an:
            R.expose("批注人", a, "ppt/comments")
        check_events(R, ev, "批注", created, modified, "ppt/comments")
    masters = [n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]
    layouts = [n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]
    R.facts["母版/版式数"] = f"{len(masters)} / {len(layouts)}"
    if len(slides) >= 5 and len(layouts) <= 1:
        R.add("struct", "weak", f"{len(slides)} 页幻灯片只用了 {len(layouts)} 个版式", "",
              "PptxGenJS 等生成库只写一个空白版式，所有内容靠绝对定位；人工做的 PPT 模板通常有多个版式", "ppt/slideLayouts")
    if pres and "ppt/presProps.xml" not in names:
        R.add("struct", "weak", "缺少演示文稿属性部件", "ppt/presProps.xml 不存在", "PowerPoint 保存的文件一定有；部分生成库和转换工具不写", "ppt/")


def xlsx_specific(z, names, R, created, modified):
    wb = read(z, "xl/workbook.xml")
    m = re.search(r'<(?:\w+:)?absPath\b[^>]*\burl="([^"]*)"', wb)
    if m:
        p = html.unescape(m.group(1))
        R.facts["上次保存位置"] = p
        scan_paths(R, p + "x", "xl/workbook.xml")
        R.expose("工作簿上次保存的文件夹", p, "xl/workbook.xml")
    fv = re.search(r"<fileVersion\b([^>]*)/?>", wb)
    if fv:
        an = re.search(r'appName="([^"]*)"', fv.group(1))
        R.facts["工作簿写入程序"] = an.group(1) if an else "（未注明）"
    elif wb:
        R.add("struct", "weak", "工作簿没有文件版本信息", "workbook.xml 缺 fileVersion", "Excel 保存的文件有；openpyxl 等生成库不写", "xl/workbook.xml")
    hidden = re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\bstate="(hidden|veryHidden)"', wb)
    for nm, st in hidden:
        R.expose("隐藏工作表", f"{html.unescape(nm)}（{st}）", "xl/workbook.xml")
    f_total = f_nocache = 0
    for n in names:
        if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            raw = read(z, n)
            for c in re.finditer(r"<c\b[^>]*?(?:/>|>(.*?)</c>)", raw, re.S):
                inner = c.group(1) or ""
                if "<f" in inner:
                    f_total += 1
                    if not re.search(r"<v\b[^>]*>[^<]+</v>", inner):
                        f_nocache += 1
            scan_text(R, xml_text(re.sub(r"<f\b.*?</f>|<f\b[^>]*/>", "", raw, flags=re.S)), n, strong_only=True)
    if f_total:
        R.facts["公式单元格"] = f"{f_total} 个，其中 {f_nocache} 个没有缓存结果"
        if f_nocache == f_total and f_total >= 3:
            R.add("struct", "medium", f"全部 {f_total} 个公式都没有计算结果缓存", "说明这个文件写出后从未在 Excel 里打开并保存过",
                  "openpyxl 等库写公式时不计算结果；人工做的表保存时一定带结果。收到后只在预览里看过也会保持这样", "xl/worksheets")
    if f_total and "xl/calcChain.xml" not in names and f_nocache < f_total:
        pass
    ev = []
    for n in names:
        if re.match(r"xl/comments\d*\.xml$", n):
            raw = read(z, n)
            for a in re.findall(r"<author>(.*?)</author>", raw, re.S):
                a = html.unescape(a).strip()
                if a:
                    ev.append((a, None, ""))
                    R.expose("批注人", a, n)
            scan_text(R, xml_text(raw), "批注内容")
        if re.match(r"xl/threadedComments/.*\.xml$", n):
            raw = read(z, n)
            for m2 in re.finditer(r"<threadedComment\b([^>]*)>", raw):
                d = re.search(r'\bdT="([^"]*)"', m2.group(1))
                ev.append(("", parse_iso(d.group(1)) if d else None, d.group(1) if d else ""))
        if re.match(r"xl/persons/.*\.xml$", n):
            for nm in re.findall(r'displayName="([^"]*)"', read(z, n)):
                R.expose("批注人", html.unescape(nm), n)
        if re.match(r"xl/externalLinks/_rels/.*\.rels$", n):
            for t in re.findall(r'Target="([^"]*)"', read(z, n)):
                R.expose("外部链接的工作簿路径", html.unescape(t), n)
    dated = [e for e in ev if e[1]]
    if dated:
        check_events(R, dated, "批注", created, modified, "xl/threadedComments")
    for a in sorted({a for a, _, _ in ev if a}):
        check_name(R, a, "批注署名", "xl/comments")
    R.facts["工作表数"] = len(re.findall(r"<sheet\b", wb))


# ---------------------------------------------------------------- PDF
def pdf_str(b):
    b = b.strip()
    if b.startswith(b"<") and not b.startswith(b"<<"):
        try:
            raw = bytes.fromhex(re.sub(rb"[^0-9A-Fa-f]", b"", b).decode())
        except Exception:
            return ""
    else:
        raw = b[1:-1] if b.startswith(b"(") else b
        raw = re.sub(rb"\\([nrtbf()\\])", lambda m: {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"", b"f": b""}.get(m.group(1), m.group(1)), raw)
        raw = re.sub(rb"\\(\d{1,3})", lambda m: bytes([int(m.group(1), 8) & 255]), raw)
    if raw[:2] == b"\xfe\xff":
        return raw[2:].decode("utf-16-be", "replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin1")


def scan_pdf(path, R, opts):
    R.kind = "PDF"
    data = open(path, "rb").read()
    info, xmp_raw, annots, text = {}, "", [], ""
    apos, ageom, id_pair = [], [], None
    used_pypdf = False
    R.fp["sha"] = hashlib.sha256(data).hexdigest()
    try:
        import pypdf  # type: ignore
        rd = pypdf.PdfReader(io.BytesIO(data))
        if rd.is_encrypted:
            try:
                rd.decrypt("")
            except Exception:
                pass
        md = rd.metadata or {}
        for k, v in md.items():
            info[str(k).lstrip("/")] = str(v)
        used_pypdf = True
        R.facts["页数"] = len(rd.pages)
        try:
            ids = rd.trailer.get("/ID")
            if ids and len(ids) >= 2:
                def _hx(v):
                    v = v.get_object() if hasattr(v, "get_object") else v
                    b = v.original_bytes if hasattr(v, "original_bytes") else (v if isinstance(v, bytes) else str(v).encode("latin1"))
                    return b.hex()
                id_pair = (_hx(ids[0]), _hx(ids[1]))
        except Exception:
            pass
        chunks = []
        for i, page in enumerate(rd.pages):
            if i < 150:
                try:
                    chunks.append(page.extract_text() or "")
                except Exception:
                    pass
            try:
                for a in page.get("/Annots") or []:
                    a = a.get_object()
                    st = str(a.get("/Subtype", ""))
                    if st in ("/Link", "/Widget", "/Popup"):
                        if st == "/Link":
                            act = a.get("/A")
                            uri = str(act.get_object().get("/URI", "")) if act is not None else ""
                            if uri:
                                chunks.append(uri)
                        continue
                    annots.append((str(a.get("/T", "")), parse_pdf_date(str(a.get("/M", "") or a.get("/CreationDate", ""))),
                                   str(a.get("/M", "") or a.get("/CreationDate", ""))))
                    c = a.get("/Contents")
                    try:
                        rect = [round(float(x), 1) for x in a.get("/Rect", [])]
                    except Exception:
                        rect = []
                    apos.append((i, rect[1]) if len(rect) == 4 else None)
                    ageom.append((i, st, tuple(rect), str(c or "")))
                    if c:
                        chunks.append(str(c))
                        R.expose("批注或文本框里填写的内容", str(c), f"第 {i + 1} 页")
            except Exception:
                pass
        text = "\n".join(chunks)
    except ImportError:
        R.notes.append("未安装 pypdf：PDF 的批注和正文痕迹未检查（pip install pypdf 后重跑可补上）")
    except Exception as e:
        R.notes.append(f"pypdf 解析失败（{type(e).__name__}），已退回基础检查：批注和正文痕迹未检查")
    if not info:
        for key in (b"Producer", b"Creator", b"Author", b"Title", b"CreationDate", b"ModDate", b"Subject", b"Keywords", b"Company"):
            ms = re.findall(rb"/" + key + rb"\s*(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]+>)", data)
            if ms:
                info[key.decode()] = pdf_str(ms[-1])
    m = re.search(rb"<x:xmpmeta.*?</x:xmpmeta>", data, re.S)
    if m:
        xmp_raw = m.group(0).decode("utf-8", "replace")
    xmp = {k: tag(xmp_raw, k) for k in ("CreatorTool", "Producer", "CreateDate", "ModifyDate", "MetadataDate", "creator")} if xmp_raw else {}
    xmp_author = html.unescape(re.sub(r"<[^>]+>", " ", xmp.get("creator", ""))).strip() if xmp else ""

    producer = info.get("Producer", "") or xmp.get("Producer", "")
    creator_tool = info.get("Creator", "") or xmp.get("CreatorTool", "")
    author = info.get("Author", "") or xmp_author
    created = parse_pdf_date(info.get("CreationDate", "")) or parse_iso(xmp.get("CreateDate", ""))
    modified = parse_pdf_date(info.get("ModDate", "")) or parse_iso(xmp.get("ModifyDate", ""))
    F = R.facts
    F["作者"] = author or "（空）"
    F["制作程序(Creator)"] = creator_tool or "（空）"
    F["PDF 生成器(Producer)"] = producer or "（空）"
    F["创建时间"] = fmt(created)
    F["修改时间"] = fmt(modified)
    if info.get("Title"):
        F["标题"] = info["Title"]
    for label, v in (("作者", author), ("标题", info.get("Title")), ("制作程序", creator_tool), ("生成器", producer),
                     ("主题", info.get("Subject")), ("关键词", info.get("Keywords")), ("公司", info.get("Company"))):
        if v:
            R.expose(label, v, "PDF 属性")

    check_name(R, author, "作者", "PDF Info")
    check_expected(R, [author], [info.get("Company", "")], opts)
    both = f"{producer} | {creator_tool}"
    for rx, desc, lvl in PDF_PRODUCERS:
        mm = re.search(rx, both, re.I)
        if mm:
            R.add("struct", lvl, f"PDF 由「{mm.group(0)}」生成：{desc}", f"Producer = {producer or '空'}；Creator = {creator_tool or '空'}",
                  "发票系统、报表平台、网页打印也走这些工具；关键是它与文件自称的来源是否相符（如自称 Word 写的报告却出自 ReportLab）", "PDF Info")
    if re.search(r"reportlab", producer, re.I) and re.match(r"^(anonymous)?$", author.strip(), re.I) and re.match(r"^(untitled)?$", (info.get("Title") or "").strip(), re.I):
        R.add("author", "medium", "作者和标题是 ReportLab 的默认值（anonymous / untitled）", f"作者 = {author or '空'}；标题 = {info.get('Title') or '空'}",
              "脚本生成时没有设置属性，AI 代码生成的 PDF 通常如此", "PDF Info")
    if not producer and not creator_tool:
        R.add("struct", "weak", "PDF 没有记录任何生成程序", "", "元数据被清除过，或由精简的工具生成", "PDF Info")
    check_core_times(R, created, modified, "", "", "PDF Info")
    if not created:
        R.add("time", "weak", "PDF 没有创建时间", "", "元数据被清除、部分生成库不写", "PDF Info")
    xc, xm = parse_iso(xmp.get("CreateDate", "")), parse_iso(xmp.get("ModifyDate", ""))
    ic, im = parse_pdf_date(info.get("CreationDate", "")), parse_pdf_date(info.get("ModDate", ""))
    for label, a, b in (("创建", ic, xc), ("修改", im, xm)):
        if a and b and abs(a - b) > TOL:
            R.add("time", "medium", f"PDF 两处元数据记录的{label}时间不一致", f"Info: {fmt(a)}；XMP: {fmt(b)}",
                  "后期用工具改过其中一处属性（常见于编辑、合并、改日期）", "PDF Info / XMP")
    ip, xp = info.get("Producer", ""), xmp.get("Producer", "")
    if ip and xp and ip.strip() != xp.strip():
        R.add("struct", "weak", "PDF 两处元数据记录的生成器不一致", f"Info: {ip}；XMP: {xp}", "文件被第二个工具处理过（合并、压缩、签名）", "PDF Info / XMP")
    eofs = data.count(b"%%EOF")
    if eofs > 1:
        F["增量保存次数"] = eofs - 1
        R.add("struct", "info", f"PDF 在生成后被追加保存过 {eofs - 1} 次", "", "加批注、签名、填表都会追加；被改动的旧内容可能仍留在文件里", "文件结构")
        R.expose("增量保存的历史版本", f"{eofs - 1} 次追加，旧内容可能可恢复", "文件结构")
    # ---- 文件头、日期写法是否与声称的生成器一致
    hm = re.match(rb"%PDF-\d\.\d\s*%([\x80-\xff][^\r\n]{2,14})", data[:48])
    hdr_name = ""
    if hm:
        for sig, nm in PDF_HEADERS:
            if hm.group(1).startswith(sig):
                hdr_name = nm
                break
        F["文件头写法"] = hdr_name or "未识别（" + hm.group(1)[:6].hex() + "）"
    is_quartz = bool(re.search(r"quartz pdfcontext", producer, re.I))
    if is_quartz and hm and hdr_name != "macOS Quartz":
        R.add("struct", "medium", "生成器栏写着 macOS，但文件头不是 macOS 的写法",
              f"文件头属于：{hdr_name or '未识别的程序'}；说明文件被别的程序整体重写过，只是保留了原来的生成器名称",
              "用 PDF 工具做过合并、压缩、解密；改写元数据的脚本也会留下这个特征", "文件头")
    if hdr_name == "ReportLab" and not re.search(r"reportlab", producer, re.I):
        R.add("struct", "medium", "文件头是 ReportLab 的写法，生成器栏却写着别的程序", f"Producer = {producer or '空'}",
              "生成器名称被改过，或文件被二次处理", "文件头")
    raw_dates = [v for k, v in info.items() if k in ("CreationDate", "ModDate") and v] + [r for _, _, r in annots if r]
    if is_quartz and raw_dates:
        odd = [d for d in raw_dates if not re.search(r"Z(00'00'?)?$", d.strip())]
        if odd:
            R.add("time", "medium", "日期的写法与声称的生成器不符", f"{len(raw_dates)} 个日期中有 {len(odd)} 个不是 macOS 的 UTC 写法，例：{odd[0]}",
                  "macOS 写日期用 Z00'00' 结尾的 UTC 格式；带时区偏移的写法说明这些日期是别的程序写入的", "PDF Info / 批注")
    if ic and im and ic.second == 0 and im.second == 0 and ic != im and any(d and d.second for _, d, _ in annots):
        R.add("time", "weak", "创建和修改时间都恰好落在整分钟", f"{info.get('CreationDate')}；{info.get('ModDate')}",
              "两个时间同时秒数为 0 的概率很低，像是手工或脚本填写的", "PDF Info")
    if id_pair:
        F["文件标识(ID)"] = f"{id_pair[0][:12]}… / {id_pair[1][:12]}…"
        R.fp.update(id0=id_pair[0].lower(), id1=id_pair[1].lower())
    R.fp.update(created=created, modified=modified)
    if ageom:
        R.fp["geom"] = tuple(sorted(ageom))
        R.fp["times"] = tuple(sorted(r for _, _, r in annots))
    raw_txt = data.decode("latin1")
    scan_paths(R, raw_txt[:20_000_000], "PDF 原始数据")
    scan_binary_labels(R, data, "PDF 原始数据", is_image=True)
    if xmp_raw:
        scan_text(R, xmp_raw, "XMP 元数据", strong_only=True)
    scan_text(R, raw_txt, "PDF 原始数据（链接）", strong_only=True)
    if text:
        F["实际字数（估算，前 150 页）"] = count_words(text)
        scan_text(R, text, "正文")
    if annots:
        an = sorted({a for a, _, _ in annots if a})
        F["批注"] = f"{len(annots)} 条，署名：{'、'.join(an) or '无'}"
        for a in an:
            R.expose("批注人", a, "PDF 批注")
        ds = sorted(d for _, d, _ in annots if d)
        if len(ds) >= 2:
            gaps = [(b - a).total_seconds() for a, b in zip(ds, ds[1:])]
            F["批注时间线"] = f"{fmt(ds[0])} ~ {fmt(ds[-1])}，跨度 {int((ds[-1] - ds[0]).total_seconds())} 秒，间隔中位数 {int(statistics.median(gaps))} 秒"
        check_events(R, annots, "批注", created, modified, "PDF 批注", positions=apos if len(apos) == len(annots) else None)
    if not used_pypdf and b"/ObjStm" in data and not info:
        R.notes.append("PDF 属性存放在压缩对象流里，基础检查读不到（安装 pypdf 可解决）")


# ---------------------------------------------------------------- 纯文本
def scan_plain(path, R, opts):
    R.kind = "文本文件"
    data = open(path, "rb").read(20_000_000)
    text = data.decode("utf-8-sig", "replace")
    R.facts["实际字数（估算）"] = count_words(text)
    if path.lower().endswith((".md", ".markdown")):
        t2 = re.sub(r"(?m)^\s{0,3}#{1,6}\s+|\*\*|```|^\s*\|?\s*:?-{3,}:?\s*\|", "", text)
        scan_text(R, t2, "正文")
    else:
        scan_text(R, text, "正文")
    scan_paths(R, text, "正文")
    scan_binary_labels(R, data, "文件内容")
    R.notes.append("纯文本文件没有内嵌元数据，只做了正文残留检查")


# ---------------------------------------------------------------- 文件外痕迹（下载来源、隔离标记、文件系统时间）
def _xattr(path, name):
    try:
        if hasattr(os, "getxattr"):
            return os.getxattr(path, name)
    except OSError:
        pass
    if sys.platform == "darwin":
        try:
            out = subprocess.run(["xattr", "-px", name, path], capture_output=True, timeout=10)
            if out.returncode == 0:
                return bytes.fromhex(re.sub(rb"[^0-9A-Fa-f]", b"", out.stdout).decode())
        except Exception:
            pass
    return None


def parse_wherefroms(raw):
    try:
        v = plistlib.loads(raw)
        return [str(x) for x in (v if isinstance(v, list) else [v]) if x]
    except Exception:
        return [u.decode("latin1") for u in re.findall(rb"https?://[\x21-\x7e]+", raw or b"")]


def parse_quarantine(raw):
    """com.apple.quarantine 的格式: 标志;十六进制时间戳;程序名;UUID"""
    parts = (raw or b"").decode("utf-8", "replace").strip().split(";")
    when = None
    if len(parts) >= 2:
        try:
            when = dt.datetime.fromtimestamp(int(parts[1], 16), UTC)
        except Exception:
            pass
    return (parts[2] if len(parts) >= 3 else ""), when


def scan_fs(path, R):
    urls, agent, qtime, found = [], "", None, False
    raw = _xattr(path, "com.apple.metadata:kMDItemWhereFroms")
    if raw:
        urls += parse_wherefroms(raw)
        found = True
    raw = _xattr(path, "com.apple.quarantine")
    if raw:
        agent, qtime = parse_quarantine(raw)
        found = True
    for nm in ("user.xdg.origin.url", "user.xdg.referrer.url"):
        raw = _xattr(path, nm)
        if raw:
            urls.append(raw.decode("utf-8", "replace"))
            found = True
    if os.name == "nt":
        try:
            with open(path + ":Zone.Identifier", "r", encoding="utf-8", errors="replace") as fh:
                z = fh.read()
            urls += re.findall(r"(?m)^(?:HostUrl|ReferrerUrl)=(.+)$", z)
            found = True
        except OSError:
            pass
    urls = [u.strip() for u in urls if u and u.strip()]
    native = sys.platform == "darwin" or os.name == "nt"
    if not found and not native:
        R.notes.append("文件外痕迹（下载来源、隔离标记、文件系统时间）未检查：这些记录只保存在文件原本所在的电脑上，"
                       "上传或通过聊天发送后就没有了。要查这一类，请在那台电脑上直接运行本脚本")
        return
    if urls:
        R.facts["下载来源"] = "；".join(clip(u, 120) for u in urls[:4])
        for u in urls:
            R.expose("下载来源网址", u, "文件扩展属性")
        ai = [u for u in urls if AI_DOMAINS.search(u)]
        if ai:
            R.add("fs", "strong", "电脑记录显示这个文件是从 AI 产品的网站下载的", ai[0],
                  "说明文件本身出自该产品的下载链接；内容是否经过人工修改需另外核实", "文件扩展属性")
        else:
            R.add("fs", "info", "电脑记录了这个文件的下载来源", clip(urls[0], 200), "可与提交人说的文件来历对照", "文件扩展属性")
    if agent:
        R.facts["下载或创建它的程序"] = agent + (f"（{fmt(qtime)}）" if qtime else "")
        R.expose("下载或创建它的程序", agent, "隔离标记")
        if AI_AGENTS.search(agent):
            R.add("fs", "strong", f"电脑记录显示这个文件是由「{agent}」放到本机的", fmt(qtime) if qtime else "",
                  "说明文件经过了这个 AI 客户端", "隔离标记")
    try:
        st = os.stat(path)
    except OSError:
        return
    mtime = dt.datetime.fromtimestamp(st.st_mtime, UTC)
    birth = getattr(st, "st_birthtime", None) or (st.st_ctime if os.name == "nt" else None)
    birth = dt.datetime.fromtimestamp(birth, UTC) if birth else None
    R.facts["电脑记录的最后写入时间"] = fmt(mtime)
    if birth:
        R.facts["电脑记录的创建时间"] = fmt(birth)
    im = R.fp.get("modified")
    if im and im - mtime > TOL:
        R.add("fs", "medium", "文件内部声称的修改时间晚于电脑记录的最后写入时间", f"内部 {fmt(im)}；电脑 {fmt(mtime)}",
              "电脑时钟错误；否则说明内部时间是写进去的，不是保存时自然产生的", "文件系统")
    if im and birth and birth - im > dt.timedelta(hours=1) and not urls and not agent:
        R.add("fs", "info", "文件在这台电脑上出现的时间晚于它声称的最后修改时间，且没有下载记录",
              f"内部最后修改 {fmt(im)}；本机创建 {fmt(birth)}",
              "从别处复制、解压、云盘同步、隔空投送都会这样；如果文件声称是在本机编辑的，这就讲不通", "文件系统")
    if qtime and im and qtime - im > TOL and R.fp.get("geom"):
        R.add("fs", "info", "文件下载到本机之后没有再保存过，但里面带着批注", f"下载于 {fmt(qtime)}；内部最后修改 {fmt(im)}",
              "批注是下载之前就有的（别人加的或在别的设备上加的）", "隔离标记")


# ---------------------------------------------------------------- 多文件比对
def cross_compare(reports):
    for A, B in itertools.combinations(reports, 2):
        a, b = A.fp, B.fp
        if not a or not b or (a.get("sha") and a.get("sha") == b.get("sha")):
            continue
        na, nb = os.path.basename(A.path), os.path.basename(B.path)
        same_final = bool(a.get("id1")) and a.get("id1") == b.get("id1")
        same_doc = bool(a.get("id0")) and a.get("id0") == b.get("id0")
        same_geom = bool(a.get("geom")) and a.get("geom") == b.get("geom")
        diff_times = a.get("times") != b.get("times")
        if same_final:
            for X, other in ((A, nb), (B, na)):
                X.add("struct", "strong", f"与「{other}」的最终保存标识完全相同，但文件内容不同", f"ID = {a['id1'][:16]}…",
                      "正常保存每次都会换一个新标识。标识相同而内容不同，说明其中一份是用工具从另一份改写出来的，没有经过正常保存", "PDF /ID")
        elif same_doc and not same_geom:
            for X, other in ((A, nb), (B, na)):
                X.add("struct", "info", f"与「{other}」出自同一份原始 PDF", f"原始 ID = {a['id0'][:16]}…", "同一份题目或模板的不同副本，很常见", "PDF /ID")
        if same_geom and diff_times:
            n = len(a["geom"])
            for X, other in ((A, nb), (B, na)):
                X.add("review", "strong" if (same_final or same_doc) else "medium",
                      f"与「{other}」的 {n} 条批注位置和内容完全相同，但时间不同", "两份文件的批注逐条重合，只有时间戳不一样",
                      "同一批批注不可能产生于两个时间：至少有一份的批注时间不是原始记录。手动放置的批注在两次独立操作中落到完全相同的坐标，可能性极小", "批注")
        if same_final or (same_geom and diff_times):
            am, bm = a.get("modified"), b.get("modified")
            if am and bm and am != bm:
                X, Y, other = (A, B, nb) if am < bm else (B, A, na)
                rewritten = any("文件头" in f["title"] or "日期的写法" in f["title"] for f in X.findings)
                X.add("time", "strong" if rewritten else "medium", f"声称的最后修改时间早于同源文件「{other}」",
                      f"本文件 {fmt(X.fp['modified'])}；对方 {fmt(Y.fp['modified'])}" + ("；且本文件带有被其他程序重写的特征" if rewritten else ""),
                      "如果本文件是从对方改写而来，它的时间就不可能更早：说明时间被改写过。也可能反过来，需结合文件头和日期写法判断哪份是原件", "多文件比对")


# ---------------------------------------------------------------- 调度与输出
EXT = {".docx": "docx", ".docm": "docx", ".dotx": "docx", ".pptx": "pptx", ".pptm": "pptx", ".ppsx": "pptx",
       ".xlsx": "xlsx", ".xlsm": "xlsx", ".pdf": "pdf", ".txt": "txt", ".md": "txt", ".markdown": "txt", ".html": "txt", ".htm": "txt"}
LEGACY = {".doc", ".xls", ".ppt", ".wps", ".et", ".dps", ".rtf", ".odt", ".ods", ".odp"}


def scan_file(path, opts):
    R = Report(path)
    ext = os.path.splitext(path)[1].lower()
    try:
        st = os.stat(path)
        R.facts["文件大小"] = f"{st.st_size:,} 字节"
    except OSError as e:
        R.notes.append(f"无法读取文件: {e}")
        return R
    base = os.path.basename(path)
    m = AI_IMAGE_NAMES.search(base) or re.search(r"chatgpt|claude|gemini|copilot|deepseek|豆包|kimi", base, re.I)
    if m:
        R.add("path", "medium", "文件名本身带 AI 产品名", base, "文件内容可能就是关于这个产品的", "文件名")
    try:
        if ext in LEGACY:
            R.kind = "旧格式或其他格式"
            R.notes.append(f"{ext} 格式暂不支持深入解析：请用原软件另存为 docx/xlsx/pptx 后再查（另存会改变部分元数据，报告里要注明）")
            data = open(path, "rb").read(30_000_000)
            scan_paths(R, data.decode("utf-16-le", "ignore") + data.decode("latin1", "ignore"), "原始数据")
            scan_text(R, data.decode("utf-16-le", "ignore") + data.decode("utf-8", "ignore"), "原始数据", strong_only=True)
        elif EXT.get(ext) in ("docx", "pptx", "xlsx"):
            scan_ooxml(path, EXT[ext], R, opts)
        elif EXT.get(ext) == "pdf":
            scan_pdf(path, R, opts)
        elif EXT.get(ext) == "txt":
            scan_plain(path, R, opts)
        else:
            R.kind = "不支持的类型"
            R.notes.append(f"不支持的文件类型 {ext}")
    except Exception as e:  # 单个文件出错不影响批量
        R.notes.append(f"扫描中途出错（{type(e).__name__}: {e}），以上结果不完整")
    try:
        scan_fs(path, R)
    except Exception as e:
        R.notes.append(f"文件外痕迹检查出错（{type(e).__name__}）")
    for k, v in list(R.facts.items()):
        if isinstance(v, set):
            R.facts[k] = "、".join(sorted(v))
    return R


def collect(paths):
    out = []
    for p in paths:
        if os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                dirs[:] = [d for d in dirs if not d.startswith(".")]
                for f in sorted(files):
                    e = os.path.splitext(f)[1].lower()
                    if (e in EXT or e in LEGACY) and not f.startswith("~$") and not f.startswith("."):
                        out.append(os.path.join(root, f))
        else:
            out.append(p)
    return out


def render_md(reports, mode):
    L = []
    if len(reports) > 1:
        L += ["# AI 痕迹扫描汇总", "", "| 文件 | 结论 | 强 | 中 | 弱 |", "|---|---|---|---|---|"]
        for R in reports:
            c = R.counts()
            code, text = R.verdict()
            L.append(f"| {os.path.basename(R.path)} | {code}：{text.split('：')[0].split('（')[0]} | {c['strong']} | {c['medium']} | {c['weak']} |")
        L.append("")
    for R in reports:
        code, text = R.verdict()
        c = R.counts()
        L += [f"# {os.path.basename(R.path)}", "", f"- 类型：{R.kind}",
              f"- 结论（{code} 级）：{text}", f"- 命中：强 {c['strong']}，中 {c['medium']}，弱 {c['weak']}，参考 {c['info']}", ""]
        L += ["## 基本属性", "", "| 项目 | 值 |", "|---|---|"]
        for k, v in R.facts.items():
            L.append(f"| {k} | {str(v).replace('|', '/')} |")
        L.append("")
        if mode in ("review", "both"):
            L += ["## 疑点（审核用）", ""]
            fs = sorted(R.findings, key=lambda f: LEVELS.index(f["level"]))
            if not fs:
                L += ["没有命中任何检查项。", ""]
            for f in fs:
                L.append(f"- **[{LEVEL_ZH[f['level']]}｜{f['category_zh']}] {f['title']}**")
                if f["evidence"]:
                    L.append(f"  - 证据：{f['evidence']}")
                if f["where"]:
                    L.append(f"  - 位置：{f['where']}")
                if f["innocent"]:
                    L.append(f"  - 正常解释：{f['innocent']}")
            L.append("")
        if mode in ("self", "both"):
            L += ["## 对外会暴露的信息（自查用）", ""]
            if not R.exposure:
                L += ["没有发现会随文件带出去的个人或路径信息。", ""]
            else:
                L += ["| 信息 | 内容 | 位置 |", "|---|---|---|"]
                for a, b, w in R.exposure:
                    L.append(f"| {a} | {b.replace('|', '/')} | {w} |")
                L.append("")
            if mode == "self":
                hot = [f for f in R.findings if f["level"] in ("strong", "medium")]
                if hot:
                    L += ["收件人用同类工具检查时会看到的疑点：", ""]
                    for f in sorted(hot, key=lambda f: LEVELS.index(f["level"])):
                        L.append(f"- [{LEVEL_ZH[f['level']]}｜{f['category_zh']}] {f['title']}" + (f"（{f['evidence']}）" if f["evidence"] else ""))
                    L.append("")
        if R.notes:
            L += ["## 没有检查到的内容", ""] + [f"- {n}" for n in R.notes] + [""]
    L += ["---", "说明：本报告检查的是文件留下的技术痕迹，不是文风。强证据说明文件或其中内容经过了程序或 AI 工具，"
          "不等于全文由 AI 撰写；没有痕迹也不等于没用 AI。Claude 等模型的统计型文本水印目前无法用本工具检测。"
          "任何处理决定都应先向当事人核实。"]
    return "\n".join(L)


def main(argv=None):
    ap = argparse.ArgumentParser(description="文档 AI 生成痕迹扫描器（只读）")
    ap.add_argument("paths", nargs="+", help="文件或文件夹")
    ap.add_argument("--mode", choices=["review", "self", "both"], default="both", help="review=审别人的文件，self=自查，both=都输出")
    ap.add_argument("--format", choices=["md", "json"], default="md")
    ap.add_argument("--out", help="把报告写到这个文件（默认打印到屏幕）")
    ap.add_argument("--expect-author", help="预期的作者姓名，用于核对")
    ap.add_argument("--expect-org", help="预期的机构名称，用于核对")
    a = ap.parse_args(argv)
    opts = {"expect_author": a.expect_author, "expect_org": a.expect_org}
    files = collect(a.paths)
    if not files:
        print("没有找到可检查的文件", file=sys.stderr)
        return 2
    reports = [scan_file(f, opts) for f in files]
    if len(reports) > 1:
        cross_compare(reports)
    if a.format == "json":
        out = json.dumps({"tool": "ai_trace_scan", "version": VERSION, "scanned_at": fmt(now()),
                          "reports": [r.to_dict() for r in reports]}, ensure_ascii=False, indent=2)
    else:
        out = render_md(reports, a.mode)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as fh:
            fh.write(out)
        print(f"报告已写入 {a.out}（{len(reports)} 个文件）")
    else:
        print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
