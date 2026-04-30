from __future__ import annotations

import datetime as _dt
import os
import posixpath
import zipfile
from pathlib import Path
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape


BASE_DIR = Path(r"D:\Lunaredu\USER\TODO")
OUT = BASE_DIR / "ICC第一次课课件_可画导入版.pptx"
ASSET_DIR = BASE_DIR / "icc_ppt_assets"

EMU = 914400
W_IN, H_IN = 13.333333, 7.5
SLIDE_W, SLIDE_H = int(W_IN * EMU), int(H_IN * EMU)

COL = {
    "ink": "172033",
    "muted": "5D6A7C",
    "line": "D9E2EC",
    "blue": "123F72",
    "blue2": "1F6DB1",
    "green": "2F7D64",
    "amber": "B56A14",
    "red": "9D3F46",
    "bg": "F5F7FB",
    "panel": "FFFFFF",
    "soft_blue": "E8F2FB",
    "soft_green": "E8F5EF",
    "soft_amber": "FFF3DF",
    "soft_red": "FAE9EA",
    "dark": "0E243F",
}

IMAGES = {
    "fridge": "https://images.pexels.com/photos/4061622/pexels-photo-4061622.jpeg?auto=compress&cs=tinysrgb&w=1200",
    "apple": "https://commons.wikimedia.org/wiki/Special:FilePath/Apple_fruits_scab.jpg?width=960",
    "office": "https://images.pexels.com/photos/37241150/pexels-photo-37241150.jpeg?auto=compress&cs=tinysrgb&w=1200",
    "car": "https://images.unsplash.com/photo-1503376780353-7e6692767b70?auto=format&fit=crop&w=1200&q=85",
}


def emu(v: float) -> int:
    return int(round(v * EMU))


def pt(v: float) -> str:
    return str(int(round(v * 100)))


def xesc(s: str) -> str:
    return escape(str(s), {'"': "&quot;"})


def jpeg_size(data: bytes) -> tuple[int, int] | None:
    if len(data) < 4 or data[:2] != b"\xff\xd8":
        return None
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        i += 2
        while marker == 0xFF and i < len(data):
            marker = data[i]
            i += 1
        if marker in (0xD8, 0xD9):
            continue
        if i + 2 > len(data):
            return None
        length = int.from_bytes(data[i:i + 2], "big")
        if length < 2 or i + length > len(data):
            return None
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            if length >= 7:
                h = int.from_bytes(data[i + 3:i + 5], "big")
                w = int.from_bytes(data[i + 5:i + 7], "big")
                return w, h
        i += length
    return None


def crop_for(img_size: tuple[int, int] | None, target_w: float, target_h: float) -> str:
    if not img_size:
        return ""
    iw, ih = img_size
    if iw <= 0 or ih <= 0:
        return ""
    img_ratio = iw / ih
    target_ratio = target_w / target_h
    attrs = {}
    if img_ratio > target_ratio:
        keep_w = target_ratio * ih
        crop = max(0.0, min(0.45, (iw - keep_w) / iw / 2))
        val = str(int(round(crop * 100000)))
        attrs["l"] = val
        attrs["r"] = val
    elif img_ratio < target_ratio:
        keep_h = iw / target_ratio
        crop = max(0.0, min(0.45, (ih - keep_h) / ih / 2))
        val = str(int(round(crop * 100000)))
        attrs["t"] = val
        attrs["b"] = val
    if not attrs:
        return ""
    return "<a:srcRect " + " ".join(f'{k}="{v}"' for k, v in attrs.items()) + "/>"


def download_images() -> dict[str, dict]:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    out: dict[str, dict] = {}
    for key, url in IMAGES.items():
        path = ASSET_DIR / f"{key}.jpg"
        data = None
        if path.exists() and path.stat().st_size > 0:
            data = path.read_bytes()
        else:
            req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urlopen(req, timeout=25) as resp:
                data = resp.read()
            path.write_bytes(data)
        out[key] = {"path": path, "size": jpeg_size(data or b"")}
    return out


class Slide:
    def __init__(self, idx: int):
        self.idx = idx
        self.shapes: list[str] = []
        self.rels: list[tuple[str, str, str]] = [(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "../slideLayouts/slideLayout1.xml",
        )]
        self.next_id = 2
        self.next_rel = 2
        self.images: list[tuple[str, Path]] = []
        self.bg()

    def sid(self) -> int:
        sid = self.next_id
        self.next_id += 1
        return sid

    def rid(self) -> str:
        rid = f"rId{self.next_rel}"
        self.next_rel += 1
        return rid

    def bg(self, color: str = COL["bg"]) -> None:
        self.rect(0, 0, W_IN, H_IN, color, color, name="Background")

    def rect(self, x, y, w, h, fill, line=None, radius=False, name="Shape") -> None:
        sid = self.sid()
        geom = "roundRect" if radius else "rect"
        line_xml = '<a:ln><a:noFill/></a:ln>' if not line else f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
        self.shapes.append(f"""
        <p:sp>
          <p:nvSpPr><p:cNvPr id="{sid}" name="{xesc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>
            <a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
            {line_xml}
          </p:spPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
        </p:sp>""")

    def line(self, x1, y1, x2, y2, color=COL["line"], width=2, name="Line") -> None:
        sid = self.sid()
        self.shapes.append(f"""
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="{sid}" name="{xesc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr>
            <a:xfrm><a:off x="{emu(min(x1, x2))}" y="{emu(min(y1, y2))}"/><a:ext cx="{emu(abs(x2-x1))}" cy="{emu(abs(y2-y1))}"/></a:xfrm>
            <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
            <a:ln w="{int(width*12700)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
          </p:spPr>
        </p:cxnSp>""")

    def text(self, x, y, w, h, text, size=24, color=COL["ink"], bold=False, align="l", valign="top", name="Text"):
        sid = self.sid()
        anchor = {"top": "t", "mid": "mid", "bottom": "b"}.get(valign, "t")
        algn = {"l": "l", "c": "ctr", "r": "r"}.get(align, "l")
        paras = []
        lines = str(text).split("\n") if str(text) else [""]
        for line in lines:
            paras.append(f"""
            <a:p>
              <a:pPr algn="{algn}"/>
              <a:r>
                <a:rPr lang="zh-CN" sz="{pt(size)}" b="{1 if bold else 0}">
                  <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                  <a:latin typeface="Microsoft YaHei"/>
                  <a:ea typeface="Microsoft YaHei"/>
                </a:rPr>
                <a:t>{xesc(line)}</a:t>
              </a:r>
              <a:endParaRPr lang="zh-CN" sz="{pt(size)}"/>
            </a:p>""")
        self.shapes.append(f"""
        <p:sp>
          <p:nvSpPr><p:cNvPr id="{sid}" name="{xesc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/><a:ln><a:noFill/></a:ln>
          </p:spPr>
          <p:txBody>
            <a:bodyPr wrap="square" anchor="{anchor}" lIns="0" tIns="0" rIns="0" bIns="0"/>
            <a:lstStyle/>
            {''.join(paras)}
          </p:txBody>
        </p:sp>""")

    def image(self, key: str, img_info: dict[str, dict], x, y, w, h, name="Picture"):
        info = img_info.get(key)
        if not info:
            self.rect(x, y, w, h, "EDF2F7", "D9E2EC", True, name="Missing image")
            self.text(x + 0.15, y + 0.15, w - 0.3, h - 0.3, "图片加载失败", 18, COL["muted"], True, "c", "mid")
            return
        rid = self.rid()
        media_name = f"image_{self.idx}_{len(self.images)+1}.jpg"
        self.images.append((media_name, info["path"]))
        self.rels.append((rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", f"../media/{media_name}"))
        sid = self.sid()
        src_rect = crop_for(info.get("size"), w, h)
        self.shapes.append(f"""
        <p:pic>
          <p:nvPicPr><p:cNvPr id="{sid}" name="{xesc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
          <p:blipFill>
            <a:blip r:embed="{rid}"/>
            {src_rect}
            <a:stretch><a:fillRect/></a:stretch>
          </p:blipFill>
          <p:spPr>
            <a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
        </p:pic>""")

    def card(self, x, y, w, h, title, body="", fill=COL["panel"], line=COL["line"], accent=None):
        self.rect(x, y, w, h, fill, line, True, name=f"Card {title}")
        if accent:
            self.rect(x, y, 0.08, h, accent, accent, False, name="Accent")
        self.text(x + 0.18, y + 0.15, w - 0.36, 0.32, title, 17, COL["blue"], True, name=f"{title} title")
        if body:
            self.text(x + 0.18, y + 0.58, w - 0.36, h - 0.70, body, 13.5, COL["ink"], False, name=f"{title} body")


def slide_xml(slide: Slide) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/><a:chOff x="0" y="0"/><a:chExt cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr>
      {''.join(slide.shapes)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''


def rels_xml(rels: list[tuple[str, str, str]]) -> str:
    body = "".join(f'<Relationship Id="{rid}" Type="{typ}" Target="{xesc(target)}"/>' for rid, typ, target in rels)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>'''


def text_header(s: Slide, eyebrow: str, title: str, time: str | None = None) -> None:
    s.text(0.62, 0.35, 8.7, 0.28, eyebrow, 13, COL["blue2"], True, name="Eyebrow")
    s.text(0.62, 0.72, 10.2, 0.62, title, 29, COL["blue"], True, name="Title")
    if time:
        s.rect(11.35, 0.48, 1.35, 0.42, COL["blue"], COL["blue"], True, "Time badge")
        s.text(11.35, 0.56, 1.35, 0.24, time, 15, "FFFFFF", True, "c", "mid", "Time")
    s.line(0.62, 1.45, 12.72, 1.45, "E8EEF6", 2, "Header rule")


def append(slides: list[Slide], slide: Slide) -> None:
    slide.idx = len(slides) + 1
    slides.append(slide)


def make_slides(img_info: dict[str, dict]) -> list[Slide]:
    slides: list[Slide] = []

    s = Slide(1)
    s.rect(0.0, 0.0, W_IN, H_IN, "F6F9FC", "F6F9FC")
    s.text(0.72, 0.70, 5.7, 0.28, "ICC全球发明大会项目｜第一次课", 14, COL["blue2"], True)
    s.text(0.72, 1.30, 6.2, 1.45, "AI时代，\n先学会发现问题", 43, COL["blue"], True)
    s.text(0.76, 3.10, 5.7, 0.72, "今天我们不急着做东西，先练一个更重要的能力：发现身边真正的麻烦，并把它说清楚。", 20, "26364C")
    for i, pill in enumerate(["5.1", "2课时", "每课时1小时", "项目题目第一版"]):
        s.rect(0.75 + i * 1.38, 4.18, 1.18, 0.36, COL["soft_blue"], "C7DCEF", True)
        s.text(0.75 + i * 1.38, 4.27, 1.18, 0.18, pill, 11.5, COL["blue"], True, "c", "mid")
    s.rect(7.25, 0.85, 5.20, 5.55, "FFFFFF", "CDD9E8", True, "Visual panel")
    nodes = [
        (7.70, 1.22, "真实问题", COL["blue"]),
        (10.50, 1.48, "用户场景", COL["green"]),
        (8.05, 5.05, "AI追问", COL["amber"]),
        (10.10, 4.88, "项目作品", COL["red"]),
    ]
    s.line(8.62, 2.00, 9.82, 3.65, "ADC1D8", 2)
    s.line(11.12, 2.25, 9.82, 3.65, "ADC1D8", 2)
    s.line(8.68, 5.52, 9.82, 3.65, "ADC1D8", 2)
    s.line(10.72, 5.35, 9.82, 3.65, "ADC1D8", 2)
    for x, y, txt, col in nodes:
        s.rect(x, y, 1.45, 0.68, col, col, True)
        s.text(x, y + 0.20, 1.45, 0.25, txt, 15, "FFFFFF", True, "c", "mid")
    s.rect(9.18, 3.05, 1.45, 1.05, "FFFFFF", "AFC6DA", True)
    s.text(9.36, 3.28, 1.08, 0.52, "定义\n问题", 18, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(2)
    text_header(s, "今日路线", "从生活麻烦到项目题目第一版")
    route = [
        ("1", "猜一猜", "AI很会回答问题，但它不知道你身边最真实的麻烦。"),
        ("2", "找一找", "把“我想做个酷东西”改成“谁遇到了什么麻烦”。"),
        ("3", "问一问", "用四个问题说清楚：谁、在哪里、怎么做、为什么不够好。"),
        ("4", "写一写", "写出项目题目第一版、可能办法、课后小任务和小组分工。"),
    ]
    for i, (n, t, b) in enumerate(route):
        x = 0.68 + i * 3.08
        s.card(x, 1.92, 2.68, 1.55, t, b, "FFFFFF", COL["line"], COL["blue"])
        s.rect(x + 0.18, 2.12, 0.38, 0.38, COL["blue"], COL["blue"], True)
        s.text(x + 0.18, 2.20, 0.38, 0.15, n, 12, "FFFFFF", True, "c", "mid")
    outcomes = [
        ("会发现", "从身边场景出发，看见真实用户遇到的具体困难。", COL["soft_blue"]),
        ("会表达", "用小组讨论和 AI 追问，把想法说成一句清楚的项目题目。", COL["soft_green"]),
        ("会行动", "带着明确任务回到真实场景观察，为后续制作作品做准备。", COL["soft_amber"]),
    ]
    for i, (t, b, fill) in enumerate(outcomes):
        s.card(1.05 + i * 4.05, 4.45, 3.45, 1.25, t, b, fill, "D8E4EF")
    append(slides, s)

    s = Slide(3)
    text_header(s, "今日产出", "两小时结束时，每组要带走什么")
    deliverables = [
        ("1. 场景痛点表", "围绕抽到的场景，写出至少 3 个具体麻烦。"),
        ("2. 主问题一句话", "能说明谁、在哪里、遇到什么麻烦、希望变成什么样。"),
        ("3. 可能办法草图", "先列 2-3 个基础版办法，不急着定最终方案。"),
        ("4. 证据任务清单", "知道课后要访谈谁、拍什么、查什么已有方案。"),
        ("5. 小组分工表", "推进、记录、查询、制作准备、汇报表达都有人负责。"),
        ("6. 材料衔接", "今天内容进入发明日志，课后证据进入查新报告线索。"),
    ]
    for i, (t, b) in enumerate(deliverables):
        x = 0.78 + (i % 3) * 4.18
        y = 1.78 + (i // 3) * 1.82
        s.card(x, y, 3.65, 1.26, t, b, ["FFFFFF", COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"], "FFFFFF"][i], "D8E4EF", [COL["blue"], COL["green"], COL["amber"], COL["red"], COL["blue2"], COL["green"]][i])
    s.text(1.20, 6.22, 10.90, 0.25, "判断标准：不是谁想法最酷，而是谁把问题说得最清楚、证据任务最具体。", 16, COL["blue"], True, "c")
    append(slides, s)

    s = Slide(3)
    text_header(s, "0-10分钟｜开场讨论", "AI这么厉害，为什么还要我们找问题？", "0-10")
    s.card(0.72, 1.85, 5.60, 1.95, "猜一猜", "如果你问 AI：“帮我做一个发明”，它会很快给你答案。可是它知道你家冰箱里什么快坏了吗？知道学校哪里最浪费电吗？", "FFFFFF", COL["line"], COL["blue"])
    s.card(6.72, 1.85, 5.60, 1.95, "AI擅长什么", "• 帮我们查资料\n• 帮我们想很多办法\n• 帮我们整理文字\n• 但不能替我们观察真实生活", COL["soft_blue"], "C7DCEF")
    s.card(0.72, 4.20, 5.60, 1.45, "一个容易踩的坑", "“我要做一个智能设备”听起来很厉害，但如果说不清谁会用、在哪里用、解决什么麻烦，就还不是一个好题目。", COL["soft_red"], "EFD1D3")
    s.card(6.72, 4.20, 5.60, 1.45, "现在请你们抢答", "AI 可以帮我们做很多事，那人最需要练什么能力？每组说出 2 个词。", COL["soft_green"], "CBE5DA")
    for i, tag in enumerate(["提问", "判断", "验证", "选择"]):
        s.rect(7.02 + i * 1.2, 5.25, 0.86, 0.30, "FFFFFF", "C7DCEF", True)
        s.text(7.02 + i * 1.2, 5.31, 0.86, 0.14, tag, 10.5, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(4)
    text_header(s, "10-25分钟｜发明与创新", "发明不是炫酷，是解决真实麻烦", "10-25")
    pillars = [
        ("有真实麻烦", "能说清楚谁遇到了麻烦，这个麻烦发生在哪里。", COL["soft_blue"], COL["blue"]),
        ("有一点改进", "不一定要做世界上从没出现过的东西，也可以让原来的办法更好用。", COL["soft_green"], COL["green"]),
        ("能被证明", "不能只说“我觉得有用”，要能用照片、访谈、试用和比较来证明。", COL["soft_amber"], COL["amber"]),
    ]
    for i, (t, b, fill, accent) in enumerate(pillars):
        s.card(0.82 + i * 4.13, 1.85, 3.55, 1.70, t, b, fill, "D8E4EF", accent)
    s.card(1.25, 4.35, 4.65, 1.35, "三个判断", "• 酷，不一定有用。\n• 智能，不一定解决问题。\n• 能帮到真实的人，才是项目的开始。", "FFFFFF", COL["line"], COL["blue"])
    s.card(7.10, 4.35, 4.65, 1.35, "请用一句话判断", "你们抽到的场景里，真正让人烦恼的地方是什么？\n留下结果：分清“酷想法”和“真实题目”。", "FFFFFF", COL["line"], COL["green"])
    append(slides, s)

    s = Slide(5)
    text_header(s, "发明与创新", "什么题目不适合今天继续做？")
    bad_cases = [
        ("只有酷，没有人", "例：我要做一个会飞的智能盒子。\n问题：不知道谁会用，也不知道解决什么麻烦。"),
        ("只换颜色或外观", "例：把普通杯子做成更好看的杯子。\n问题：没有说明改进带来的真实帮助。"),
        ("太大、太危险、太难验证", "例：我要做无人驾驶汽车。\n问题：小组无法在课程里做出基础版并测试。"),
        ("已经有成熟产品，改进点不清楚", "例：直接复制智能冰箱或盲区雷达。\n问题：需要说明我们的不同和改进。"),
    ]
    fills = [COL["soft_red"], COL["soft_amber"], COL["soft_red"], COL["soft_blue"]]
    for i, (t, b) in enumerate(bad_cases):
        x = 0.80 + (i % 2) * 6.05
        y = 1.85 + (i // 2) * 1.78
        s.card(x, y, 5.55, 1.30, t, b, fills[i], "D8E4EF", COL["red"] if i != 3 else COL["blue"])
    s.rect(1.45, 6.05, 10.45, 0.52, "F8FBFF", "C7DCEF", True)
    s.text(1.45, 6.18, 10.45, 0.18, "今天只推进：能观察、能比较、能做出基础版、能讲清楚改进的题目。", 16, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(5)
    text_header(s, "比较页", "把“想做一个东西”改成“解决一个问题”")
    rows = [
        ("我要做一个智能冰箱。", "帮助忙碌家庭更早发现临期或遗忘的食材，减少浪费。"),
        ("我要做一个农业识别机器人。", "帮助果农更早记录果实病斑，减少人工巡检遗漏。"),
        ("我要做一个自动关灯系统。", "发现无人或光线充足区域仍开灯的情况，减少无效用电。"),
        ("我要做一个车辆安全提醒设备。", "在车辆转弯或倒车时提醒盲区风险，降低剐蹭和行人风险。"),
    ]
    s.text(1.08, 1.80, 4.0, 0.30, "粗糙想法", 17, COL["red"], True, "c")
    s.text(7.05, 1.80, 4.0, 0.30, "真实题目", 17, COL["green"], True, "c")
    for i, (bad, good) in enumerate(rows):
        y = 2.25 + i * 0.82
        s.rect(0.85, y, 4.65, 0.55, COL["soft_red"], "EFD1D3", True)
        s.text(1.05, y + 0.13, 4.25, 0.18, bad, 13.5, COL["ink"], False, "c", "mid")
        s.text(5.72, y + 0.13, 0.55, 0.18, "→", 18, COL["blue"], True, "c", "mid")
        s.rect(6.45, y, 5.45, 0.55, COL["soft_green"], "CBE5DA", True)
        s.text(6.65, y + 0.10, 5.05, 0.24, good, 12.5, COL["ink"], False, "c", "mid")
    s.rect(1.55, 6.10, 10.20, 0.55, "F8FBFF", "C7DCEF", True)
    s.text(1.55, 6.24, 10.20, 0.18, "真实题目 = 具体用户 + 具体场景 + 真实困难 + 目标结果", 18, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(6)
    text_header(s, "内容页", "同一个方向里，用户不同，问题也不同")
    user_cards = [
        ("冰箱食材检测", "• 忙碌家长：买了菜但忘记吃，最后变质\n• 老人家庭：靠经验判断新鲜度，容易不放心\n• 独居学生：食材少但管理乱，经常重复买", COL["soft_blue"], COL["blue"]),
        ("果园果实病状检查", "• 果农：巡园时间有限，早期病斑容易漏看\n• 采摘园：开放前要快速判断果实状态\n• 管理人员：需要记录不同区域的果实问题", COL["soft_green"], COL["green"]),
        ("办公区无效照明监控", "• 员工：离开时忘记关灯\n• 物业：很难逐个房间检查\n• 行政：想知道哪些区域经常浪费电", COL["soft_amber"], COL["amber"]),
        ("车辆盲区视觉提醒", "• 司机：倒车或右转时看不到关键区域\n• 骑行者：靠近车辆侧后方时不容易被看到\n• 行人：停车场和小区道路里风险更高", COL["soft_red"], COL["red"]),
    ]
    for i, c in enumerate(user_cards):
        x = 0.75 + (i % 2) * 6.05
        y = 1.95 + (i // 2) * 2.05
        s.card(x, y, 5.55, 1.65, c[0], c[1], c[2], "D8E4EF", c[3])
    s.rect(2.20, 6.25, 8.95, 0.45, "F8FBFF", "C7DCEF", True)
    s.text(2.20, 6.36, 8.95, 0.16, "先选清楚用户，再讨论做什么方案。", 17, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(7)
    text_header(s, "抽到哪个，就研究哪个", "四个项目方向，各 4 个生活场景")
    scene_cards = [
        ("冰箱食材检测", "fridge", ["忙碌家庭的临期食材提醒", "老人家庭的食材新鲜度判断", "独居学生的冰箱食材遗忘问题", "家庭重复购买导致的食材浪费"], COL["blue"]),
        ("果园果实病状检查", "apple", ["果农日常巡园时的病果发现", "采摘园开放前的果实状态检查", "果实病斑早期识别与记录", "果园管理人员的批量巡检记录"], COL["green"]),
        ("办公区无效照明监控", "office", ["无人办公室持续开灯", "自然光充足但室内灯未关闭", "会议室使用结束后照明未关闭", "教室/办公区分区照明浪费"], COL["amber"]),
        ("车辆盲区视觉提醒", "car", ["小区道路转弯时的行人盲区", "停车场倒车时的后方盲区", "车辆右转时的骑行者盲区", "变道时侧后方车辆提醒"], COL["red"]),
    ]
    for i, (title, key, items, accent) in enumerate(scene_cards):
        x = 0.48 + i * 3.18
        s.rect(x, 1.82, 2.85, 4.75, "FFFFFF", COL["line"], True)
        s.image(key, img_info, x, 1.82, 2.85, 1.45, title)
        s.rect(x, 3.08, 2.85, 0.22, accent, accent)
        s.text(x + 0.18, 3.48, 2.49, 0.30, title, 14, COL["blue"], True, "c")
        body = "\n".join(f"{n}. {item}" for n, item in enumerate(items, 1))
        s.text(x + 0.22, 3.95, 2.42, 1.55, body, 11.3, COL["ink"])
    s.text(1.25, 6.84, 10.80, 0.20, "抽到一个场景后，不急着想作品，先找出里面的 3 个麻烦。", 14, COL["blue"], True, "c")
    append(slides, s)

    s = Slide(8)
    text_header(s, "25-45分钟｜找麻烦工具", "痛点词库：把“很麻烦”说具体")
    pain_words = [
        ("时间", ["太慢", "等待久", "来不及发现", "重复检查"]),
        ("准确", ["看不清", "判断不准", "容易漏掉", "记录混乱"]),
        ("浪费", ["多买", "坏掉", "空开灯", "重复劳动"]),
        ("安全", ["看不到", "反应慢", "容易碰撞", "提醒不及时"]),
        ("成本", ["人力多", "材料贵", "维护难", "不方便长期用"]),
        ("体验", ["记不住", "步骤多", "不愿意用", "不会判断"]),
    ]
    fills = [COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"], "FFFFFF", COL["soft_blue"]]
    for i, (title, words) in enumerate(pain_words):
        x = 0.72 + (i % 3) * 4.12
        y = 1.82 + (i // 3) * 1.75
        s.card(x, y, 3.58, 1.20, title, " / ".join(words), fills[i], "D8E4EF", [COL["blue"], COL["green"], COL["amber"], COL["red"], COL["blue2"], COL["green"]][i])
    s.text(1.10, 6.12, 11.10, 0.30, "写痛点时尽量使用具体词：谁在什么时候，因为哪个环节，产生了什么损失或风险。", 16, COL["blue"], True, "c")
    append(slides, s)

    s = Slide(8)
    text_header(s, "25-45分钟｜找麻烦", "围绕抽到的场景，写出 3 个具体麻烦", "25-45")
    s.card(0.80, 1.90, 5.50, 1.35, "先不要急着想办法", "现在只找麻烦，不讨论设备、材料和外观。", COL["soft_blue"], "C7DCEF", COL["blue"])
    s.card(6.80, 1.90, 5.50, 1.35, "这些不算好答案", "• 太麻烦了\n• 不方便\n• 想做一个 AI 设备\n\n要继续追问到具体人和具体场景。", COL["soft_red"], "EFD1D3", COL["red"])
    s.card(0.80, 3.75, 5.50, 1.55, "拆开来看", "谁在这个场景里最容易遇到麻烦？\n这个麻烦什么时候发生？\n现在靠什么办法处理？", "FFFFFF", COL["line"], COL["green"])
    s.card(6.80, 3.75, 5.50, 1.55, "现在请你们写", "每组至少写出 3 个具体痛点。\n每个痛点都要包含：谁 + 在哪里 + 遇到什么麻烦。", COL["soft_green"], "CBE5DA", COL["green"])
    append(slides, s)

    s = Slide(9)
    text_header(s, "25-45分钟｜四组示范", "同一个方向，痛点可以这样写")
    examples = [
        ("冰箱食材检测", "忙碌家庭晚饭前才发现蔬菜已经变质，既浪费钱，也影响做饭安排。"),
        ("果园果实病状检查", "果农巡园时只能凭肉眼快速看，早期小病斑容易被漏掉，后面扩散更难处理。"),
        ("办公区无效照明监控", "会议结束后房间没人，但灯继续开到下班后，物业很难及时发现。"),
        ("车辆盲区视觉提醒", "小区转弯时司机看不到侧前方行人，等发现时距离已经很近。"),
    ]
    fills = [COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"]]
    for i, (t, b) in enumerate(examples):
        x = 0.85 + (i % 2) * 6.05
        y = 1.90 + (i // 2) * 1.75
        s.card(x, y, 5.50, 1.22, t, b, fills[i], "D8E4EF")
    s.rect(1.25, 6.05, 10.80, 0.55, "F8FBFF", "C7DCEF", True)
    s.text(1.25, 6.20, 10.80, 0.16, "好痛点句子 = 用户 + 场景 + 现在的问题 + 造成的后果", 17, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(9)
    text_header(s, "45-60分钟｜四问追问", "从 3 个痛点里筛出一个主问题", "45-60")
    qs = [
        ("谁？", "具体是哪一类人遇到问题？"),
        ("在哪里？", "问题发生在家庭、果园、办公室、停车场，还是更具体的位置？"),
        ("现在怎么办？", "他们现在用什么办法解决？"),
        ("为什么不够好？", "这个办法慢、不准、容易忘、成本高，还是不安全？"),
    ]
    for i, (t, b) in enumerate(qs):
        x = 0.90 + (i % 2) * 6.00
        y = 1.95 + (i // 2) * 1.35
        s.card(x, y, 5.25, 1.02, t, b, "FFFFFF", COL["line"], COL["blue"])
    s.card(0.90, 4.95, 5.25, 1.20, "筛选标准", "• 最常见\n• 最影响人\n• 最容易观察\n• 后面能做出基础版作品", COL["soft_amber"], "F2DCB8", COL["amber"])
    s.card(6.90, 4.95, 5.25, 1.20, "留下结果", "每组筛出 1 个主问题，并能用 30 秒说清楚为什么选择它。", COL["soft_green"], "CBE5DA", COL["green"])
    append(slides, s)

    s = Slide(10)
    text_header(s, "45-60分钟｜追问示范", "把一个模糊痛点追问到主问题")
    s.rect(0.95, 1.85, 11.40, 0.60, COL["soft_red"], "EFD1D3", True)
    s.text(1.20, 2.02, 10.90, 0.18, "模糊说法：办公室经常浪费电。", 17, COL["red"], True, "c", "mid")
    steps = [
        ("谁？", "行政或物业负责发现浪费。"),
        ("在哪里？", "会议室、教室、公共办公区。"),
        ("现在怎么办？", "靠人巡查，或下班后统一检查。"),
        ("为什么不够好？", "发现晚、没有记录、不知道哪个区域最常浪费。"),
    ]
    for i, (t, b) in enumerate(steps):
        x = 0.80 + i * 3.08
        s.card(x, 3.02, 2.65, 1.18, t, b, "FFFFFF", COL["line"], COL["blue"])
    s.rect(0.95, 5.25, 11.40, 0.82, COL["soft_green"], "CBE5DA", True)
    s.text(1.20, 5.48, 10.90, 0.20, "主问题：怎样帮助行政更早发现无人会议室持续开灯，并记录哪些区域经常浪费？", 18, COL["green"], True, "c", "mid")
    append(slides, s)

    s = Slide(10)
    text_header(s, "方法一｜先观察人", "不要先想设备，先看谁遇到了麻烦")
    s.text(0.95, 1.85, 5.2, 0.80, "做发明不是一上来就画外观、想功能，而是先去看真实的人：他们什么时候最麻烦？现在怎么解决？哪里最不方便？", 20, COL["ink"])
    steps = ["看一看", "问一问", "记一记", "比一比", "再决定"]
    desc = ["现场发生了什么", "用户怎么描述麻烦", "把原话和现象留下来", "不同人的情况有什么不同", "选一个最值得做的问题"]
    for i, st in enumerate(steps):
        x = 0.95 + i * 2.35
        s.rect(x, 3.65, 1.95, 0.62, COL["blue"], COL["blue"], True)
        s.text(x, 3.82, 1.95, 0.16, st, 15, "FFFFFF", True, "c", "mid")
        s.text(x - 0.05, 4.48, 2.05, 0.42, desc[i], 12.5, COL["muted"], False, "c")
    s.rect(1.55, 6.05, 10.20, 0.52, "F8FBFF", "C7DCEF", True)
    s.text(1.55, 6.18, 10.20, 0.18, "先看人，再想办法。", 18, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(11)
    text_header(s, "方法二｜回到根本问题", "用户真正想要的结果是什么？")
    s.card(0.82, 1.85, 5.65, 1.45, "这是什么意思", "不要只问“要不要摄像头、传感器、提醒灯”，先问用户真正想要什么结果。", COL["soft_blue"], "C7DCEF", COL["blue"])
    s.card(6.85, 1.85, 5.65, 1.45, "例子", "冰箱项目不是先问“用什么识别”，而是先问：家里人想更早知道什么？想少浪费什么？", COL["soft_green"], "CBE5DA", COL["green"])
    s.card(2.08, 4.05, 9.15, 1.25, "三个追问", "• 如果没有高科技，最低限度怎样帮助用户？\n• 用户最想减少的是时间、错误、浪费，还是危险？\n• 哪个结果可以被测试出来？", "FFFFFF", COL["line"], COL["amber"])
    s.rect(1.55, 6.05, 10.20, 0.52, "F8FBFF", "C7DCEF", True)
    s.text(1.55, 6.18, 10.20, 0.18, "先问结果，再想方案；先做基础版，再逐步升级。", 17, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(12)
    text_header(s, "知识内容｜套用到四个项目", "同一个方法，怎么用在你们的项目里")
    examples = [
        ("冰箱食材检测", "根本结果：少浪费、少忘记、吃得更放心。"),
        ("果园果实病状检查", "根本结果：更早发现异常，巡检记录更清楚。"),
        ("办公区无效照明监控", "根本结果：少无效用电，知道哪些区域常浪费。"),
        ("车辆盲区视觉提醒", "根本结果：在危险发生前提醒司机注意盲区。"),
    ]
    fills = [COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"]]
    accents = [COL["blue"], COL["green"], COL["amber"], COL["red"]]
    for i, (t, b) in enumerate(examples):
        x = 0.9 + (i % 2) * 6.0
        y = 2.0 + (i // 2) * 1.65
        s.card(x, y, 5.35, 1.12, t, b, fills[i], "D8E4EF", accents[i])
    s.text(1.15, 5.90, 11.0, 0.36, "每个组都要把“想做什么设备”翻译成“想帮助谁达到什么结果”。", 20, COL["blue"], True, "c")
    append(slides, s)

    s = Slide(13)
    text_header(s, "60-80分钟｜项目题目评分", "第一版题目要过四个小检查")
    checks = [
        ("清楚的人", "能说出主要用户，不写“大家”“所有人”。"),
        ("具体的场景", "能说出发生在哪里、什么时候发生。"),
        ("真实的麻烦", "不是一句“很不方便”，而是能看见后果。"),
        ("可做的基础版", "后面能用简单硬件、模型、材料做出可测试版本。"),
    ]
    for i, (t, b) in enumerate(checks):
        x = 0.85 + i * 3.05
        s.card(x, 1.95, 2.65, 1.38, t, b, ["FFFFFF", COL["soft_blue"], COL["soft_green"], COL["soft_amber"]][i], "D8E4EF", [COL["blue"], COL["green"], COL["amber"], COL["red"]][i])
    s.text(0.95, 4.30, 11.20, 0.30, "课堂互评方式：每组读题目，其他组只提一个问题，帮助它变得更清楚。", 18, COL["blue"], True, "c")
    s.card(1.55, 5.08, 4.90, 1.00, "可以追问", "“谁最需要？” “怎么证明？” “现在办法哪里不好？”", COL["soft_green"], "CBE5DA")
    s.card(6.95, 5.08, 4.90, 1.00, "先不评价", "不急着说“好不好做”，先帮它把问题说完整。", COL["soft_red"], "EFD1D3")
    append(slides, s)

    s = Slide(13)
    text_header(s, "60-80分钟｜写题目第一版", "用一句话，把你们的项目说清楚", "60-80")
    s.card(0.95, 1.95, 5.30, 1.25, "先检查“人和场景”", "• 谁遇到这个问题？\n• 问题发生在哪里？\n• 这个场景经常出现吗？", COL["soft_blue"], "C7DCEF", COL["blue"])
    s.card(7.05, 1.95, 5.30, 1.25, "再检查“最想要的结果”", "• 想减少什么？\n• 想提高什么？\n• 怎么证明它真的变好了？", COL["soft_green"], "CBE5DA", COL["green"])
    s.rect(1.05, 4.25, 11.25, 1.20, "F8FBFF", "C7DCEF", True)
    s.text(1.35, 4.50, 10.65, 0.42, "为了解决【谁】在【什么地方】遇到的【什么麻烦】，我们希望做一个能【带来什么帮助】的办法。", 22, COL["blue"], True, "c", "mid")
    s.text(2.00, 6.12, 9.35, 0.30, "留下结果：每组写出项目题目第一版。", 19, COL["green"], True, "c")
    append(slides, s)

    s = Slide(14)
    text_header(s, "题目示范", "像这样把题目说清楚")
    drafts = [
        ("冰箱食材检测", "帮助忙碌家庭发现临期或被遗忘食材，减少食物浪费。"),
        ("果园果实病状检查", "帮助果农在巡园时更早发现病斑，并留下可比较记录。"),
        ("办公区无效照明监控", "帮助办公区发现无人或光线充足时仍开灯的浪费情况。"),
        ("车辆盲区视觉提醒", "帮助司机在转弯、倒车或变道时注意盲区风险。"),
    ]
    fills = [COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"]]
    for i, (t, b) in enumerate(drafts):
        x = 0.95 + (i % 2) * 6.05
        y = 1.95 + (i // 2) * 1.80
        s.card(x, y, 5.45, 1.30, t, b, fills[i], "D8E4EF")
    s.text(1.45, 6.10, 10.45, 0.26, "注意：这是第一版题目，后面会根据访谈、查新和制作过程继续修改。", 16, COL["muted"], False, "c")
    append(slides, s)

    s = Slide(15)
    text_header(s, "80-100分钟｜让 AI 帮我们多问几句", "找到可能办法和课后要看的东西", "80-100")
    s.card(0.72, 1.85, 5.60, 1.35, "AI可以帮什么", "• 补充可能的用户和场景\n• 列出现有解决办法\n• 提示可能的解决办法\n• 提醒还需要观察什么、询问什么", COL["soft_blue"], "C7DCEF", COL["blue"])
    s.card(6.72, 1.85, 5.60, 1.35, "AI不能替我们做什么", "• 不能替你证明问题真实存在\n• 不能替你决定最终题目\n• 不能替你完成访谈、观察和测试", COL["soft_red"], "EFD1D3", COL["red"])
    prompt = "我正在做一个青少年发明项目，方向是【冰箱食材检测 / 果园果实病状检查 / 办公区无效照明监控 / 车辆盲区视觉提醒】。\n请你不要直接替我决定方案，而是帮我追问：\n1. 具体用户是谁？ 2. 问题发生在什么场景？\n3. 现有办法有哪些？ 4. 这些办法有什么不足？\n5. 可以有哪些基础版方案？ 6. 我需要补充哪些证据？"
    s.rect(0.90, 3.85, 11.65, 2.35, "F8FBFF", "C7DCEF", True)
    s.text(1.15, 4.10, 11.10, 1.85, prompt, 13.2, COL["ink"])
    append(slides, s)

    s = Slide(16)
    text_header(s, "80-100分钟｜AI回答怎么用", "不要照抄，要筛选")
    process = [
        ("先删掉", "太贵、太危险、超出课程时间、没有办法测试的方案。"),
        ("再留下", "能做基础版、能买到材料、能在课堂里演示的方案。"),
        ("继续追问", "需要什么传感器？怎样判断成功？和已有产品有什么不同？"),
        ("变成任务", "把 AI 的想法改写成访谈、拍照、查资料和制作准备。"),
    ]
    for i, (t, b) in enumerate(process):
        x = 0.80 + (i % 2) * 6.05
        y = 1.85 + (i // 2) * 1.75
        s.card(x, y, 5.50, 1.20, t, b, ["FFFFFF", COL["soft_green"], COL["soft_blue"], COL["soft_amber"]][i], "D8E4EF", [COL["red"], COL["green"], COL["blue"], COL["amber"]][i])
    s.rect(1.25, 6.05, 10.80, 0.55, "F8FBFF", "C7DCEF", True)
    s.text(1.25, 6.20, 10.80, 0.16, "AI 是追问助手，不是最终裁判。最终选择要回到用户、证据和可制作性。", 16, COL["blue"], True, "c", "mid")
    append(slides, s)

    s = Slide(16)
    text_header(s, "整理页", "把 AI 的回答整理成两类内容")
    s.card(0.78, 1.85, 5.75, 1.55, "2-3 个可能办法", "• 提醒类：让用户及时知道风险\n• 识别类：辅助判断状态或异常\n• 记录类：把过程和变化留下来\n• 联动类：把传感器、图像、提示结合起来", COL["soft_green"], "CBE5DA", COL["green"])
    s.card(6.80, 1.85, 5.75, 1.55, "课后要看的东西", "• 要访谈谁？\n• 要拍什么照片？\n• 要查哪些已有产品？\n• 要比较哪些现有办法的不足？", COL["soft_blue"], "C7DCEF", COL["blue"])
    case_names = ["冰箱识别案例", "植物病害识别案例", "办公占用感知案例", "盲区监测案例"]
    case_desc = ["AI Vision Inside", "Plantix 拍照识别", "办公空间传感器", "Blind Spot Monitor"]
    for i, (a, b) in enumerate(zip(case_names, case_desc)):
        x = 0.90 + i * 3.05
        s.card(x, 4.58, 2.58, 1.05, a, b, "FFFFFF", COL["line"])
    s.text(1.25, 6.22, 10.90, 0.25, "留下结果：每组 2-3 个可能办法 + 一份课后证据清单。", 17, COL["green"], True, "c")
    append(slides, s)

    s = Slide(17)
    text_header(s, "查新报告入门", "今天不用写完整报告，但要知道查什么")
    items = [
        ("查什么", "已有产品、已有做法、相关案例、相似技术。"),
        ("去哪查", "淘宝 / 百度 / 小红书 / 产品官网 / 视频平台 / 生活观察。"),
        ("怎么截图", "保留名称、图片、功能说明、价格或使用场景。"),
        ("怎么比较", "写一句：它解决了什么？还有哪里不够好？"),
    ]
    for i, (t, b) in enumerate(items):
        x = 0.85 + (i % 2) * 6.05
        y = 1.85 + (i // 2) * 1.72
        s.card(x, y, 5.50, 1.18, t, b, ["FFFFFF", COL["soft_blue"], COL["soft_green"], COL["soft_amber"]][i], "D8E4EF", [COL["blue"], COL["green"], COL["amber"], COL["red"]][i])
    s.card(1.55, 5.42, 10.20, 0.88, "例子", "查到“智能冰箱能识别食材”，还要写：它适合整机冰箱，但我们的课堂基础版可以先解决“提醒临期和遗忘”这一小问题。", COL["soft_green"], "CBE5DA")
    append(slides, s)

    s = Slide(18)
    text_header(s, "发明日志入门", "今天先写骨架，课后再补证据")
    log_items = [
        ("我想解决什么问题？", "课堂写：主问题一句话。"),
        ("我是怎么想到这个问题的？", "课堂写：抽到的场景 + 小组追问。"),
        ("我想达到什么结果？", "课堂写：减少浪费、提高发现速度、降低风险等。"),
        ("可能的解决方案有哪些？", "课堂写：2-3 个可能办法，先不定最终版。"),
        ("在哪里查看是否有人做过？", "课后写：淘宝、官网、搜索截图和已有方案。"),
        ("谁可以帮助完成？", "课堂写：小组分工和可请教的人。"),
    ]
    for i, (t, b) in enumerate(log_items):
        x = 0.70 + (i % 3) * 4.18
        y = 1.80 + (i // 3) * 1.58
        s.card(x, y, 3.65, 1.05, t, b, ["FFFFFF", COL["soft_blue"], COL["soft_green"], COL["soft_amber"], COL["soft_red"], "FFFFFF"][i], "D8E4EF")
    s.text(1.10, 6.18, 11.10, 0.24, "原则：课堂先写“能确定的”，课后用真实证据把它补完整。", 16, COL["blue"], True, "c")
    append(slides, s)

    s = Slide(17)
    text_header(s, "100-115分钟｜团队分工", "让项目从今天开始有人推进", "100-115")
    roles = [
        ("推进与汇报", "确认今天要完成什么，并负责 1 分钟说明。"),
        ("记录与整理", "记录小组讨论、项目题目、修改想法和课后观察。"),
        ("查询与截图", "查找已有产品、案例、图片、资料，并保留截图。"),
        ("制作准备", "后续关注可能用到的零件、工具、结构和制作方式。"),
    ]
    for i, (t, b) in enumerate(roles):
        x = 0.70 + i * 3.10
        s.card(x, 2.05, 2.72, 1.78, t, b, "FFFFFF", COL["line"], [COL["blue"], COL["green"], COL["amber"], COL["red"]][i])
    s.card(1.25, 4.95, 4.85, 1.15, "小组任务", "每组 3-4 人，可以一人多职；关键是记录、查询、制作准备和汇报表达都有人负责。", COL["soft_blue"], "C7DCEF")
    s.card(7.25, 4.95, 4.85, 1.15, "留下结果", "小组分工表：姓名、负责事项、今天后要完成的第一件事。", COL["soft_green"], "CBE5DA")
    append(slides, s)

    s = Slide(18)
    text_header(s, "100-115分钟｜1分钟小组展示", "每组按这个顺序说")
    speech = [
        ("第1句", "我们抽到的方向是……场景是……"),
        ("第2句", "我们发现最主要的问题是……"),
        ("第3句", "这个问题影响的人是……发生在……"),
        ("第4句", "我们想到的可能办法有……"),
        ("第5句", "课后我们要去补充的证据是……"),
    ]
    for i, (t, b) in enumerate(speech):
        x = 0.90 + i * 2.45
        s.rect(x, 2.05, 1.92, 0.50, COL["blue"], COL["blue"], True)
        s.text(x, 2.18, 1.92, 0.14, t, 13, "FFFFFF", True, "c", "mid")
        s.rect(x, 2.78, 1.92, 1.55, "FFFFFF", COL["line"], True)
        s.text(x + 0.14, 3.02, 1.64, 0.85, b, 12.2, COL["ink"], False, "c", "mid")
    s.card(1.45, 5.35, 10.40, 0.88, "听的同学要做什么", "每组展示后，其他组只问一个帮助它变清楚的问题，不抢着给方案。", COL["soft_green"], "CBE5DA")
    append(slides, s)

    s = Slide(18)
    text_header(s, "115-120分钟｜收束", "今天之后，项目怎么继续往下做", "115-120")
    s.card(0.78, 1.85, 5.75, 2.00, "今天写进发明日志骨架", "• 我们想解决什么问题？\n• 这个问题影响谁？\n• 问题发生在什么场景？\n• 我们希望达到什么结果？\n• 目前想到哪些可能办法？", COL["soft_green"], "CBE5DA", COL["green"])
    s.card(6.80, 1.85, 5.75, 2.00, "课后补证据，为查新报告做准备", "• 访谈记录：至少 2 条\n• 场景照片或截图：至少 3 张\n• 已有方案截图：至少 3 个\n• 现有办法不足：每个写 1 句", COL["soft_blue"], "C7DCEF", COL["blue"])
    outputs = ["发明日志骨架：题目、用户、场景、目标结果", "真实观察证据：访谈、照片和截图", "查新报告线索：已有产品和已有办法截图", "小组表达：1 分钟讲清楚项目问题"]
    for i, out in enumerate(outputs):
        x = 0.90 + (i % 2) * 6.00
        y = 4.60 + (i // 2) * 0.78
        s.rect(x, y, 5.30, 0.48, "FFFFFF", "CBE5DA", True)
        s.text(x + 0.18, y + 0.13, 4.94, 0.16, out, 12.8, COL["ink"], False, "c", "mid")
    append(slides, s)

    s = Slide(19)
    text_header(s, "115-120分钟｜离开教室前检查", "每组把这 6 项对齐")
    checklist = [
        ("主问题", "已经能用一句话说清。"),
        ("用户", "不是“所有人”，而是具体一类人。"),
        ("场景", "知道问题在哪里、什么时候发生。"),
        ("证据", "知道课后要访谈、拍照、截图什么。"),
        ("查新", "知道要找哪些已有产品或已有做法。"),
        ("分工", "每个人知道今天后要做的第一件事。"),
    ]
    for i, (t, b) in enumerate(checklist):
        x = 0.75 + (i % 3) * 4.15
        y = 1.85 + (i // 3) * 1.70
        s.rect(x, y, 3.55, 1.12, "FFFFFF", "CBE5DA", True)
        s.text(x + 0.18, y + 0.16, 0.44, 0.28, "✓", 18, COL["green"], True, "c", "mid")
        s.text(x + 0.70, y + 0.16, 2.65, 0.24, t, 16, COL["blue"], True)
        s.text(x + 0.70, y + 0.55, 2.65, 0.28, b, 12.5, COL["ink"])
    s.text(1.25, 6.18, 10.80, 0.25, "如果有一项说不清，先不要急着做作品，回到前面的四问追问。", 16, COL["red"], True, "c")
    append(slides, s)

    s = Slide(19)
    text_header(s, "课后任务", "下节课带回四样东西")
    homework = [
        ("1. 更清楚的问题定义", "用“谁、在哪里、遇到什么问题、希望达到什么结果”写成一句话。", COL["soft_blue"], COL["blue"]),
        ("2. 一组真实证据", "访谈记录、场景照片、使用过程照片、现有产品截图。", COL["soft_green"], COL["green"]),
        ("3. 查新报告线索", "淘宝、百度、小红书、产品官网、生活里的已有做法都可以作为初查来源，截图并写一句不足。", COL["soft_amber"], COL["amber"]),
        ("4. 发明日志第一版", "把课堂讨论和课后观察放到同一份发明日志里，方便下一次继续推进。", COL["soft_red"], COL["red"]),
    ]
    for i, (t, b, fill, accent) in enumerate(homework):
        x = 0.95 + (i % 2) * 6.05
        y = 1.95 + (i // 2) * 1.75
        s.card(x, y, 5.50, 1.22, t, b, fill, "D8E4EF", accent)
    s.text(1.50, 6.25, 10.35, 0.24, "下一次课：AI时代如何学习技术，硬件基础、简单嵌入式和建模基础。", 15, COL["muted"], False, "c")
    append(slides, s)

    s = Slide(20)
    text_header(s, "素材来源", "本课件使用和推荐的图片 / 视频案例")
    sources = [
        "冰箱配图 / 视频：Pexels fridge photo & refrigerator video",
        "冰箱案例：Samsung AI Vision Inside",
        "果实病状配图：Wikimedia Commons Apple fruits scab.jpg",
        "果园视频：Pexels apple orchard video",
        "果实识别案例：FAO e-Agriculture Plantix app case",
        "办公照明配图 / 视频：Pexels office photo & empty office video",
        "办公感知案例：XY Sense workplace occupancy sensor demo",
        "车辆配图 / 视频：Unsplash car side mirror photo & Pexels side mirror video",
        "盲区监测案例：MyCarDoesWhat blind spot monitor guide",
    ]
    for i, item in enumerate(sources):
        y = 1.78 + i * 0.47
        s.rect(0.95, y, 11.40, 0.34, "FFFFFF", COL["line"], True)
        s.text(1.18, y + 0.08, 10.95, 0.12, item, 10.8, COL["ink"])
    s.text(1.25, 6.42, 10.85, 0.22, "后续制作精美版 PPT 时，可替换为高清下载图、视频封面或课堂播放二维码。", 13.5, COL["muted"], False, "c")
    append(slides, s)

    return slides


def content_types(slide_count: int) -> str:
    overrides = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="jpg" ContentType="image/jpeg"/>',
        '<Default Extension="jpeg" ContentType="image/jpeg"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>',
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for i in range(1, slide_count + 1):
        overrides.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + "".join(overrides) + "</Types>"


def presentation_xml(slide_count: int) -> str:
    ids = []
    for i in range(1, slide_count + 1):
        ids.append(f'<p:sldId id="{255+i}" r:id="rId{i+1}"/>')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{''.join(ids)}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>'''


def presentation_rels(slide_count: int) -> str:
    rels = [("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml")]
    for i in range(1, slide_count + 1):
        rels.append((f"rId{i+1}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{i}.xml"))
    rels.extend([
        (f"rId{slide_count+2}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"),
        (f"rId{slide_count+3}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps", "presProps.xml"),
        (f"rId{slide_count+4}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps", "viewProps.xml"),
        (f"rId{slide_count+5}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles", "tableStyles.xml"),
    ])
    return rels_xml(rels)


def static_parts(slide_count: int) -> dict[str, str]:
    now = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat()
    return {
        "[Content_Types].xml": content_types(slide_count),
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''',
        "docProps/core.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ICC第一次课课件：AI时代的问题发现与定义</dc:title>
  <dc:creator>Lunaredu</dc:creator>
  <cp:lastModifiedBy>Lunaredu</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>''',
        "docProps/app.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OpenAI Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>{slide_count}</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides>
</Properties>''',
        "ppt/presentation.xml": presentation_xml(slide_count),
        "ppt/_rels/presentation.xml.rels": presentation_rels(slide_count),
        "ppt/presProps.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>''',
        "ppt/viewProps.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>''',
        "ppt/tableStyles.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>''',
        "ppt/theme/theme1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Lunaredu">
  <a:themeElements>
    <a:clrScheme name="Lunaredu"><a:dk1><a:srgbClr val="172033"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="123F72"/></a:dk2><a:lt2><a:srgbClr val="F5F7FB"/></a:lt2><a:accent1><a:srgbClr val="1F6DB1"/></a:accent1><a:accent2><a:srgbClr val="2F7D64"/></a:accent2><a:accent3><a:srgbClr val="B56A14"/></a:accent3><a:accent4><a:srgbClr val="9D3F46"/></a:accent4><a:accent5><a:srgbClr val="5D6A7C"/></a:accent5><a:accent6><a:srgbClr val="D9E2EC"/></a:accent6><a:hlink><a:srgbClr val="1F6DB1"/></a:hlink><a:folHlink><a:srgbClr val="9D3F46"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Lunaredu"><a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont><a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Lunaredu"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/><a:extraClrSchemeLst/>
</a:theme>''',
        "ppt/slideMasters/slideMaster1.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/><a:chOff x="0" y="0"/><a:chExt cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>''',
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels_xml([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "../theme/theme1.xml"),
        ]),
        "ppt/slideLayouts/slideLayout1.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/><a:chOff x="0" y="0"/><a:chExt cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>''',
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels_xml([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "../slideMasters/slideMaster1.xml")
        ]),
    }


def write_pptx(slides: list[Slide]) -> None:
    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, xml in static_parts(len(slides)).items():
            z.writestr(name, xml)
        written_media = set()
        for slide in slides:
            z.writestr(f"ppt/slides/slide{slide.idx}.xml", slide_xml(slide))
            z.writestr(f"ppt/slides/_rels/slide{slide.idx}.xml.rels", rels_xml(slide.rels))
            for media_name, path in slide.images:
                arc = posixpath.join("ppt/media", media_name)
                if arc not in written_media:
                    z.write(path, arc)
                    written_media.add(arc)


def main() -> None:
    img_info = download_images()
    slides = make_slides(img_info)
    write_pptx(slides)
    print(f"exported={OUT}")
    print(f"slides={len(slides)}")
    print(f"images={sum(len(s.images) for s in slides)}")


if __name__ == "__main__":
    main()

