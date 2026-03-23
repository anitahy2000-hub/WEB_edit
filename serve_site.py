#!/usr/bin/env python3

from __future__ import annotations

from cgi import FieldStorage
from html.parser import HTMLParser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse, urljoin, parse_qs, quote
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
import base64
import argparse
import html
import json
import mimetypes
import os
import posixpath
import re
import uuid
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def guess_markdown_heading(style_name: str | None, text: str) -> str:
    if not text:
        return text

    style = (style_name or "").lower()
    if "heading1" in style or "title" in style:
        return f"# {text}"
    if "heading2" in style:
        return f"## {text}"
    return text


def classify_paragraph_style(style_name: str | None) -> str:
    style = (style_name or "").lower()
    if "heading1" in style or "title" in style:
        return "h1"
    if "heading2" in style:
        return "h2"
    return "p"


def style_fragment(tag: str) -> str:
    if tag == "h1":
        return (
            'style="font-size:21px;line-height:1.4;margin:0 0 14px;'
            'font-weight:700;"'
        )
    if tag == "h2":
        return (
            'style="font-size:14px;line-height:1.5;margin:18px 0 10px;'
            'font-weight:700;"'
        )
    if tag == "img":
        return (
            'style="display:block;width:100%;max-width:100%;height:auto;'
            'margin:14px 0;object-fit:cover;"'
        )
    return 'style="font-size:16px;line-height:1.8;margin:0 0 14px;"'


def markdown_to_formatted_html(markdown_text: str) -> str:
    lines = markdown_text.splitlines()
    html_parts: list[str] = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("## "):
            html_parts.append(
                f'<h2 {style_fragment("h2")}>{html.escape(line[3:])}</h2>'
            )
            continue

        if line.startswith("# "):
            html_parts.append(
                f'<h1 {style_fragment("h1")}>{html.escape(line[2:])}</h1>'
            )
            continue

        image_match = re.match(r"!\[(.*?)\]\((.+)\)", line)
        if image_match:
            alt_text = html.escape(image_match.group(1))
            src = html.escape(image_match.group(2), quote=True)
            html_parts.append(
                f'<img {style_fragment("img")} src="{src}" alt="{alt_text}" />'
            )
            continue

        html_parts.append(f'<p {style_fragment("p")}>{html.escape(line)}</p>')

    return "\n".join(html_parts)


def html_to_markdown_text(html_text: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", html_text, flags=re.IGNORECASE)
    text = re.sub(r"</?(strong|b)>", "**", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(em|i)>", "*", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(normalize_whitespace(text))


def is_wechat_article_url(url: str) -> bool:
    parsed = urlparse(url.strip())
    host = (parsed.netloc or "").lower()
    return parsed.scheme in {"http", "https"} and (
        "mp.weixin.qq.com" in host or "weixin.qq.com" in host
    )


def normalize_remote_url(raw_url: str, referer_url: str) -> str:
    cleaned = html.unescape(raw_url.strip())
    if not cleaned:
        return ""
    if cleaned.startswith("//"):
        return f"https:{cleaned}"
    return urljoin(referer_url, cleaned)


def build_proxy_image_url(image_url: str, referer_url: str) -> str:
    normalized_url = normalize_remote_url(image_url, referer_url)
    return (
        "/api/proxy-image"
        f"?src={quote(normalized_url, safe='')}"
        f"&referer={quote(referer_url, safe='')}"
    )


def fetch_remote_image_bytes(
    image_url: str, referer_url: str
) -> tuple[bytes, str, str]:
    normalized_url = normalize_remote_url(image_url, referer_url)
    if not normalized_url:
        raise ValueError("图片链接为空。")

    request = Request(
        normalized_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0 Safari/537.36"
            ),
            "Referer": referer_url,
            "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
        },
    )

    with urlopen(request, timeout=20) as response:
        image_bytes = response.read()
        content_type = (response.headers.get("Content-Type") or "").split(";")[0].strip().lower()

    if not content_type.startswith("image/"):
        guessed_type, _ = mimetypes.guess_type(normalized_url)
        content_type = guessed_type or "image/jpeg"

    return image_bytes, content_type, normalized_url


def localize_remote_images_in_html(
    content_html: str, referer_url: str, temp_dir: Path
) -> str:
    temp_dir.mkdir(parents=True, exist_ok=True)

    def replace_img(match: re.Match[str]) -> str:
        prefix = match.group(1)
        src = html.unescape(match.group(2))
        suffix = match.group(3)

        if not src or src.startswith("data:image/") or src.startswith("/_temp/"):
            return match.group(0)

        try:
            image_bytes, content_type, normalized_url = fetch_remote_image_bytes(src, referer_url)
            guessed_suffix = mimetypes.guess_extension(content_type) or Path(urlparse(normalized_url).path).suffix or ".jpg"
            filename = f"{uuid.uuid4()}{guessed_suffix.lower()}"
            output_path = temp_dir / filename
            output_path.write_bytes(image_bytes)
            local_src = html.escape(f"/_temp/{filename}", quote=True)
            return f'{prefix}{local_src}{suffix}'
        except Exception:
            proxy_src = html.escape(build_proxy_image_url(src, referer_url), quote=True)
            return f'{prefix}{proxy_src}{suffix}'

    return re.sub(r'(<img\b[^>]*\bsrc=")([^"]+)(".*?>)', replace_img, content_html, flags=re.I)


class WechatArticleHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.title_parts: list[str] = []
        self.capture_title = False
        self.capture_depth = 0
        self.skip_depth = 0
        self.content_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = {key: value or "" for key, value in attrs}
        classes = attr_map.get("class", "")
        element_id = attr_map.get("id", "")

        if tag == "h1" and (
            element_id == "activity-name" or "rich_media_title" in classes
        ):
            self.capture_title = True

        if tag == "title" and not self.title_parts:
            self.capture_title = True

        if self.capture_depth == 0 and tag == "div":
            if (
                element_id in {"js_content", "img-content"}
                or "rich_media_content" in classes
            ):
                self.capture_depth = 1
                return

        if self.capture_depth <= 0:
            return

        self.capture_depth += 1

        if tag in {"script", "style"}:
            self.skip_depth += 1
            return

        if self.skip_depth:
            return

        normalized = self.normalize_start_tag(tag, attr_map)
        if normalized:
            self.content_parts.append(normalized)

    def handle_endtag(self, tag: str) -> None:
        if self.capture_title and tag in {"h1", "title"}:
            self.capture_title = False

        if self.capture_depth <= 0:
            return

        if self.skip_depth:
            if tag in {"script", "style"}:
                self.skip_depth -= 1
        else:
            normalized = self.normalize_end_tag(tag)
            if normalized:
                self.content_parts.append(normalized)

        self.capture_depth -= 1

    def handle_data(self, data: str) -> None:
        if self.capture_title:
            self.title_parts.append(data)

        if self.capture_depth > 0 and not self.skip_depth:
            self.content_parts.append(html.escape(data))

    def handle_entityref(self, name: str) -> None:
        self.handle_data(f"&{name};")

    def handle_charref(self, name: str) -> None:
        self.handle_data(f"&#{name};")

    def normalize_start_tag(self, tag: str, attr_map: dict[str, str]) -> str:
        if tag in {"strong", "b"}:
            return "<strong>"
        if tag in {"em", "i"}:
            return "<em>"
        if tag == "br":
            return "<br />"
        if tag in {"p", "h1", "h2", "h3", "blockquote", "ul", "ol", "li"}:
            return f"<{tag}>"
        if tag == "a":
            href = attr_map.get("href", "").strip()
            if href and not href.lower().startswith("javascript:"):
                return f'<a href="{html.escape(href, quote=True)}">'
            return "<a>"
        if tag == "img":
            src = (
                attr_map.get("data-src", "").strip()
                or attr_map.get("src", "").strip()
                or attr_map.get("data-backsrc", "").strip()
            )
            if not src:
                return ""
            alt = attr_map.get("alt", "").strip() or "公众号图片"
            return (
                f'<img src="{html.escape(src, quote=True)}" '
                f'alt="{html.escape(alt, quote=True)}" />'
            )
        return ""

    def normalize_end_tag(self, tag: str) -> str:
        if tag in {"strong", "b"}:
            return "</strong>"
        if tag in {"em", "i"}:
            return "</em>"
        if tag in {"p", "h1", "h2", "h3", "blockquote", "ul", "ol", "li", "a"}:
            return f"</{tag}>"
        return ""

    def get_title(self) -> str:
        raw = normalize_whitespace("".join(self.title_parts))
        return re.sub(r"\s*[-_－]\s*微信公众平台.*$", "", raw).strip()

    def get_content_html(self) -> str:
        content = "".join(self.content_parts).strip()
        content = re.sub(r"(?:<br\s*/?>\s*){3,}", "<br /><br />", content, flags=re.I)
        return content


def fetch_wechat_article(url: str, temp_dir: Path) -> tuple[str, str]:
    request = Request(
        url.strip(),
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0 Safari/537.36"
            )
        },
    )
    with urlopen(request, timeout=15) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        html_text = response.read().decode(charset, errors="replace")

    parser = WechatArticleHTMLParser()
    parser.feed(html_text)
    title = parser.get_title()
    content_html = parser.get_content_html()

    if not content_html:
        raise ValueError("未能识别到公众号正文内容。")

    if title and "<h1" not in content_html.lower():
        content_html = f"<h1>{html.escape(title)}</h1>\n{content_html}"

    content_html = localize_remote_images_in_html(content_html, url.strip(), temp_dir)

    return title, content_html


def normalize_relationship_target(target: str) -> str:
    normalized = target.replace("\\", "/").strip()
    while normalized.startswith("../"):
        normalized = normalized[3:]
    if normalized.startswith("/"):
        normalized = normalized[1:]
    if normalized.startswith("word/"):
        normalized = normalized[5:]
    return normalized


def emu_to_px(value: str | None) -> int | None:
    if not value:
        return None
    try:
        return max(1, round(int(value) / 9525))
    except (TypeError, ValueError):
        return None


def ensure_temp_image_file(
    rel_id: str | None,
    rels: dict[str, str],
    media: dict[str, bytes],
    temp_dir: Path,
    image_cache: dict[str, dict[str, str | int]],
) -> dict[str, str | int] | None:
    if not rel_id:
        return None

    cached = image_cache.get(rel_id)
    if cached:
        return cached

    target = rels.get(rel_id)
    if not target:
        return None

    normalized_target = normalize_relationship_target(target)
    media_bytes = media.get(normalized_target)
    if not media_bytes:
        return None

    suffix = Path(normalized_target).suffix or ".bin"
    filename = f"{uuid.uuid4()}{suffix.lower()}"
    output_path = temp_dir / filename
    output_path.write_bytes(media_bytes)

    asset = {
        "src": f"/_temp/{filename}",
        "alt": "文档图片",
        "source_part": f"word/{normalized_target}",
    }
    image_cache[rel_id] = asset
    return asset


def get_run_text(run: ET.Element) -> str:
    parts: list[str] = []
    for node in list(run):
        if node.tag == f"{{{NS['w']}}}t":
            parts.append(node.text or "")
        elif node.tag == f"{{{NS['w']}}}tab":
            parts.append("\t")
        elif node.tag == f"{{{NS['w']}}}br":
            parts.append("\n")
    return "".join(parts)


def wrap_run_html(text: str, run: ET.Element) -> str:
    escaped = html.escape(text).replace("\n", "<br />")
    if not escaped:
        return ""

    rpr = run.find("w:rPr", NS)
    if rpr is None:
        return escaped

    if rpr.find("w:b", NS) is not None:
        escaped = f"<strong>{escaped}</strong>"
    if rpr.find("w:i", NS) is not None:
        escaped = f"<em>{escaped}</em>"
    return escaped


def get_run_images(
    run: ET.Element,
    rels: dict[str, str],
    media: dict[str, bytes],
    temp_dir: Path,
    image_cache: dict[str, dict[str, str | int]],
) -> list[dict[str, str | int]]:
    rel_ids: list[str] = []
    rel_ids.extend(
        node.attrib.get(f"{{{NS['r']}}}embed")
        for node in run.findall(".//a:blip", NS)
        if node.attrib.get(f"{{{NS['r']}}}embed")
    )
    rel_ids.extend(
        node.attrib.get(f"{{{NS['r']}}}link")
        for node in run.findall(".//a:blip", NS)
        if node.attrib.get(f"{{{NS['r']}}}link")
    )
    rel_ids.extend(
        node.attrib.get(f"{{{NS['r']}}}id")
        for node in run.findall(".//v:imagedata", NS)
        if node.attrib.get(f"{{{NS['r']}}}id")
    )

    extent = run.find(".//wp:extent", NS)
    width = emu_to_px(extent.attrib.get("cx") if extent is not None else None)
    height = emu_to_px(extent.attrib.get("cy") if extent is not None else None)

    images: list[dict[str, str | int]] = []
    for rel_id in rel_ids:
        asset = ensure_temp_image_file(rel_id, rels, media, temp_dir, image_cache)
        if asset:
            image_data = dict(asset)
            if width:
                image_data["width"] = width
            if height:
                image_data["height"] = height
            images.append(image_data)
    return images


def paragraph_to_blocks(
    paragraph: ET.Element,
    style_name: str | None,
    rels: dict[str, str],
    media: dict[str, bytes],
    temp_dir: Path,
    image_cache: dict[str, dict[str, str | int]],
) -> list[dict[str, str]]:
    blocks: list[dict[str, str]] = []
    inline_html_parts: list[str] = []
    inline_text_parts: list[str] = []

    def flush_text_block() -> None:
      nonlocal inline_html_parts, inline_text_parts
      plain = normalize_whitespace("".join(inline_text_parts))
      rich = "".join(inline_html_parts).strip()
      inline_html_parts = []
      inline_text_parts = []
      if plain:
          tag = classify_paragraph_style(style_name)
          blocks.append({"type": tag, "text": plain, "html": rich or html.escape(plain)})

    for child in list(paragraph):
        if child.tag == f"{{{NS['w']}}}r":
            run_text = get_run_text(child)
            if run_text:
                inline_text_parts.append(run_text)
                inline_html_parts.append(wrap_run_html(run_text, child))

            for image_data in get_run_images(child, rels, media, temp_dir, image_cache):
                flush_text_block()
                blocks.append({"type": "img", **image_data})

        elif child.tag == f"{{{NS['w']}}}proofErr":
            continue
        else:
            descendant_text = "".join(node.text or "" for node in child.findall(".//w:t", NS))
            if descendant_text:
                inline_text_parts.append(descendant_text)
                inline_html_parts.append(html.escape(descendant_text))

    flush_text_block()
    return blocks


def blocks_to_markdown(blocks: list[dict[str, str]]) -> str:
    lines: list[str] = []
    for block in blocks:
        if block["type"] == "h1":
            lines.append(f"# {html_to_markdown_text(block.get('html', block.get('text', '')))}")
        elif block["type"] == "h2":
            lines.append(f"## {html_to_markdown_text(block.get('html', block.get('text', '')))}")
        elif block["type"] == "img":
            lines.append(f"![文档图片]({block.get('src', '')})")
        else:
            lines.append(html_to_markdown_text(block.get("html", block.get("text", ""))))
    return "\n\n".join(line for line in lines if line.strip())


def blocks_to_formatted_html(blocks: list[dict[str, str]], source_filename: str = "") -> str:
    html_parts: list[str] = []
    for block in blocks:
        if block["type"] == "h1":
            html_parts.append(f'<h1 {style_fragment("h1")}>{block.get("html", "")}</h1>')
        elif block["type"] == "h2":
            html_parts.append(f'<h2 {style_fragment("h2")}>{block.get("html", "")}</h2>')
        elif block["type"] == "img":
            src = html.escape(block.get("src", ""), quote=True)
            width_attr = f' width="{block["width"]}"' if block.get("width") else ""
            height_attr = f' height="{block["height"]}"' if block.get("height") else ""
            source_doc_attr = (
                f' data-source-doc="{html.escape(source_filename, quote=True)}"'
                if source_filename
                else ""
            )
            source_part_attr = (
                f' data-source-part="{html.escape(str(block.get("source_part", "")), quote=True)}"'
                if block.get("source_part")
                else ""
            )
            html_parts.append(
                f'<img data-layer="photo" src="{src}" align="BOTTOM"{width_attr}{height_attr}{source_doc_attr}{source_part_attr} border="0"/>'
            )
        else:
            html_parts.append(f'<p {style_fragment("p")}>{block.get("html", "")}</p>')
    return "\n".join(html_parts)


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


class SimpleHTMLToBlocksParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.blocks: list[dict[str, str]] = []
        self.current_tag: str | None = None
        self.current_text: list[str] = []
        self.current_attrs: dict[str, str] = {}

    def flush_text(self) -> None:
        if not self.current_tag:
            return
        text = "".join(self.current_text).strip()
        self.current_text = []
        if not text:
            self.current_tag = None
            self.current_attrs = {}
            return
        block_type = "p"
        if self.current_tag == "h1":
            block_type = "h1"
        elif self.current_tag == "h2":
            block_type = "h2"
        block = {"type": block_type, "text": text}
        block.update(self.current_attrs)
        self.blocks.append(block)
        self.current_tag = None
        self.current_attrs = {}

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = {key: value or "" for key, value in attrs}
        if tag in {"h1", "h2", "p"}:
            self.flush_text()
            self.current_tag = tag
            self.current_text = []
            self.current_attrs = {"style": attrs_dict.get("style", "")}
            return
        if tag == "br" and self.current_tag:
            self.current_text.append("\n")
            return
        if tag == "img":
            self.flush_text()
            self.blocks.append(
                {
                    "type": "img",
                    "src": attrs_dict.get("src", ""),
                    "width": attrs_dict.get("width", ""),
                    "height": attrs_dict.get("height", ""),
                    "alt": attrs_dict.get("alt", "图片"),
                }
            )

    def handle_endtag(self, tag: str) -> None:
        if self.current_tag == tag:
            self.flush_text()

    def handle_data(self, data: str) -> None:
        if self.current_tag:
            self.current_text.append(data)


def html_to_export_blocks(html_text: str) -> list[dict[str, str]]:
    parser = SimpleHTMLToBlocksParser()
    parser.feed(html_text)
    parser.flush_text()
    return parser.blocks


def parse_style_map(style_text: str | None) -> dict[str, str]:
    result: dict[str, str] = {}
    if not style_text:
        return result
    for chunk in style_text.split(";"):
        if ":" not in chunk:
            continue
        key, value = chunk.split(":", 1)
        result[key.strip().lower()] = value.strip()
    return result


def px_to_half_points(value: str | None, fallback_px: float) -> int:
    raw = value or ""
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)", raw)
    px = fallback_px
    if match:
        px = float(match.group(1))
    return max(2, round(px * 1.5))


def line_height_to_twips(value: str | None, font_half_points: int) -> int:
    if not value:
        return round((font_half_points / 2) * 20 * 1.6)
    raw = value.strip().lower()
    number_match = re.search(r"([0-9]+(?:\.[0-9]+)?)", raw)
    if not number_match:
        return round((font_half_points / 2) * 20 * 1.6)
    number = float(number_match.group(1))
    if "px" in raw:
        return round(number * 20)
    return round((font_half_points / 2) * 20 * number)


def margin_to_twips(style_map: dict[str, str], side: str, fallback_px: float) -> int:
    direct = style_map.get(f"margin-{side}")
    if direct:
        return round(float(re.search(r"([0-9]+(?:\.[0-9]+)?)", direct).group(1)) * 20) if re.search(r"([0-9]+(?:\.[0-9]+)?)", direct) else round(fallback_px * 20)
    shorthand = style_map.get("margin")
    if shorthand:
        parts = [part for part in shorthand.split() if part]
        values = []
        for part in parts:
            match = re.search(r"([0-9]+(?:\.[0-9]+)?)", part)
            values.append(float(match.group(1)) if match else 0.0)
        if len(values) == 1:
            target = values[0]
        elif len(values) == 2:
            target = values[0] if side in {"top", "bottom"} else values[1]
        elif len(values) == 3:
            target = values[0] if side == "top" else values[2] if side == "bottom" else values[1]
        else:
            target = values[0] if side == "top" else values[2] if side == "bottom" else values[1]
        return round(target * 20)
    return round(fallback_px * 20)


def px_to_emu(value: str | None, fallback_px: int) -> int:
    try:
        if value:
            return max(1, round(float(value) * 9525))
    except ValueError:
        pass
    return fallback_px * 9525


def image_bytes_from_src(src: str, root: Path) -> tuple[bytes, str] | None:
    if not src:
        return None
    if src.startswith("data:image/"):
        match = re.match(r"data:(image/[a-zA-Z0-9.+-]+);base64,(.+)", src, re.DOTALL)
        if not match:
            return None
        mime = match.group(1)
        data = base64.b64decode(match.group(2))
        extension = mimetypes.guess_extension(mime) or ".png"
        return data, extension
    if src.startswith("/_temp/"):
        path = root / src.lstrip("/")
        if path.exists():
            return path.read_bytes(), path.suffix or ".png"
    return None


def export_docx_bytes(html_text: str, root: Path) -> bytes:
    blocks = html_to_export_blocks(html_text)
    relationships: list[str] = []
    media_files: list[tuple[str, bytes]] = []
    body_parts: list[str] = []
    image_index = 1

    def text_paragraph(text: str, kind: str) -> str:
        escaped = xml_escape(text)
        defaults = {
            "h1": {"font_px": 21, "line": "1.4", "before": 0, "after": 14},
            "h2": {"font_px": 14, "line": "1.5", "before": 18, "after": 10},
            "p": {"font_px": 16, "line": "1.8", "before": 0, "after": 14},
        }
        style_map = parse_style_map(block.get("style", ""))
        default = defaults[kind]
        font_size = px_to_half_points(style_map.get("font-size"), default["font_px"])
        line_spacing = line_height_to_twips(style_map.get("line-height"), font_size)
        spacing_before = margin_to_twips(style_map, "top", default["before"])
        spacing_after = margin_to_twips(style_map, "bottom", default["after"])
        bold_xml = "<w:b/>" if kind in {"h1", "h2"} else ""
        if kind == "h1":
            return (
                f'<w:p><w:pPr><w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="{line_spacing}" w:lineRule="exact"/></w:pPr>'
                f'<w:r><w:rPr>{bold_xml}<w:sz w:val="{font_size}"/></w:rPr>'
                f"<w:t xml:space=\"preserve\">{escaped}</w:t></w:r></w:p>"
            )
        if kind == "h2":
            return (
                f'<w:p><w:pPr><w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="{line_spacing}" w:lineRule="exact"/></w:pPr>'
                f'<w:r><w:rPr>{bold_xml}<w:sz w:val="{font_size}"/></w:rPr>'
                f"<w:t xml:space=\"preserve\">{escaped}</w:t></w:r></w:p>"
            )
        return (
            f'<w:p><w:pPr><w:spacing w:before="{spacing_before}" w:after="{spacing_after}" w:line="{line_spacing}" w:lineRule="exact"/></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="{font_size}"/></w:rPr>'
            f"<w:t xml:space=\"preserve\">{escaped}</w:t></w:r></w:p>"
        )

    for block in blocks:
        if block["type"] in {"h1", "h2", "p"}:
            body_parts.append(text_paragraph(block.get("text", ""), block["type"]))
            continue

        if block["type"] == "img":
            image_payload = image_bytes_from_src(block.get("src", ""), root)
            if not image_payload:
                continue
            image_bytes, extension = image_payload
            media_name = f"image{image_index}{extension}"
            media_files.append((f"word/media/{media_name}", image_bytes))
            rel_id = f"rIdImage{image_index}"
            relationships.append(
                f'<Relationship Id="{rel_id}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="media/{media_name}"/>'
            )
            cx = px_to_emu(block.get("width"), 554)
            cy = px_to_emu(block.get("height"), 330)
            body_parts.append(
                "<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" "
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
                "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                f'<wp:extent cx="{cx}" cy="{cy}"/>'
                "<wp:docPr id=\"1\" name=\"Picture\"/>"
                "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Image\"/><pic:cNvPicPr/></pic:nvPicPr>"
                "<pic:blipFill>"
                f'<a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
                "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
                "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
                f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
                "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
            )
            image_index += 1

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels_root = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        + "".join(relationships)
        + "</Relationships>"
    )

    document_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<w:body>"
        + "".join(body_parts)
        + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" "
        "w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>"
        "</w:body></w:document>"
    )

    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels_root)
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/_rels/document.xml.rels", doc_rels)
        for path, data in media_files:
            archive.writestr(path, data)
    return output.getvalue()


def parse_docx_bytes(file_bytes: bytes, temp_dir: Path, source_filename: str = "") -> tuple[str, str]:
    temp_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(BytesIO(file_bytes)) as archive:
        document_xml = archive.read("word/document.xml")
        rels_xml = archive.read("word/_rels/document.xml.rels")
        media_files = {
            normalize_relationship_target(path): archive.read(path)
            for path in archive.namelist()
            if path.startswith("word/media/")
        }

    rels_root = ET.fromstring(rels_xml)
    rels = {}
    for rel in rels_root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rel_id:
            rels[rel_id] = normalize_relationship_target(target)

    document_root = ET.fromstring(document_xml)
    body = document_root.find("w:body", NS)
    if body is None:
        raise ValueError("Word 文档中未找到正文内容。")

    blocks: list[dict[str, str]] = []
    image_cache: dict[str, dict[str, str | int]] = {}

    for child in list(body):
        if child.tag == f"{{{NS['w']}}}p":
            style_name = None
            p_style = child.find("./w:pPr/w:pStyle", NS)
            if p_style is not None:
                style_name = p_style.attrib.get(f"{{{NS['w']}}}val")
            blocks.extend(
                paragraph_to_blocks(child, style_name, rels, media_files, temp_dir, image_cache)
            )

        elif child.tag == f"{{{NS['w']}}}tbl":
            for paragraph in child.findall(".//w:p", NS):
                blocks.extend(
                    paragraph_to_blocks(paragraph, None, rels, media_files, temp_dir, image_cache)
                )

    markdown = blocks_to_markdown(blocks)
    formatted_html = blocks_to_formatted_html(blocks, source_filename=source_filename)
    return markdown, formatted_html


class AppHandler(SimpleHTTPRequestHandler):
    server_version = "DocFormatterHTTP/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.path = "/text_format_tool.html"
            super().do_GET()
            return
        if parsed.path == "/api/proxy-image":
            self.handle_proxy_image(parsed)
            return
        super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/export-docx":
            self.handle_export_docx()
            return
        if parsed.path == "/api/import-wechat":
            self.handle_import_wechat()
            return
        if parsed.path != "/api/convert-docx":
            self.send_error(404, "Unknown endpoint")
            return

        try:
            form = FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )

            upload = form["file"] if "file" in form else None
            if upload is None or not getattr(upload, "file", None):
                self.send_json(
                    400,
                    {
                        "error": "请上传一个 .docx 文件。",
                    },
                )
                return

            filename = getattr(upload, "filename", "") or ""
            if not filename.lower().endswith(".docx"):
                self.send_json(
                    400,
                    {
                        "error": "当前只支持 .docx 格式，暂不支持旧版 .doc。",
                    },
                )
                return

            file_bytes = upload.file.read()
            temp_dir = Path.cwd() / "_temp"
            markdown, formatted_html = parse_docx_bytes(
                file_bytes, temp_dir, source_filename=filename
            )
            self.send_json(
                200,
                {
                    "filename": filename,
                    "markdown": markdown,
                    "formattedHtml": formatted_html,
                },
            )
        except KeyError as exc:
            self.send_json(400, {"error": f"文档结构缺失必要文件：{exc.args[0]}。"})
        except zipfile.BadZipFile:
            self.send_json(400, {"error": "上传的文件不是有效的 .docx 文档。"})
        except Exception as exc:  # noqa: BLE001
            self.send_json(500, {"error": f"处理失败：{exc}"})

    def handle_proxy_image(self, parsed) -> None:
        try:
            params = parse_qs(parsed.query)
            image_url = params.get("src", [""])[0].strip()
            referer_url = params.get("referer", [""])[0].strip()

            if not image_url or not referer_url:
                self.send_error(400, "Missing image src or referer")
                return

            image_bytes, content_type, _normalized_url = fetch_remote_image_bytes(
                image_url, referer_url
            )
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(image_bytes)))
            self.end_headers()
            self.wfile.write(image_bytes)
        except HTTPError as exc:
            self.send_error(exc.code, f"Proxy image fetch failed: {exc.reason}")
        except URLError as exc:
            self.send_error(502, f"Proxy image fetch failed: {exc.reason}")
        except Exception as exc:  # noqa: BLE001
            self.send_error(500, f"Proxy image fetch failed: {exc}")

    def handle_import_wechat(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))
            article_url = str(payload.get("url", "")).strip()

            if not article_url:
                self.send_json(400, {"error": "请输入公众号文章链接。"})
                return

            if not is_wechat_article_url(article_url):
                self.send_json(400, {"error": "请输入有效的微信公众号文章链接。"})
                return

            temp_dir = Path.cwd() / "_temp"
            title, formatted_html = fetch_wechat_article(article_url, temp_dir)
            self.send_json(
                200,
                {
                    "title": title or "公众号文章",
                    "formattedHtml": formatted_html,
                },
            )
        except HTTPError as exc:
            self.send_json(400, {"error": f"抓取失败：远程返回 {exc.code}。"})
        except URLError as exc:
            self.send_json(400, {"error": f"抓取失败：{exc.reason}。"})
        except json.JSONDecodeError:
            self.send_json(400, {"error": "请求格式错误。"})
        except Exception as exc:  # noqa: BLE001
            self.send_json(500, {"error": f"公众号文章导入失败：{exc}"})

    def handle_export_docx(self) -> None:
        try:
            content_type = self.headers.get("Content-Type", "")
            html_text = ""
            filename = "formatted-content"

            if "application/json" in content_type:
                content_length = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8"))
                html_text = payload.get("html", "")
                filename = payload.get("filename", "formatted-content")
            else:
                form = FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ={
                        "REQUEST_METHOD": "POST",
                        "CONTENT_TYPE": content_type,
                    },
                )
                html_text = form.getfirst("html", "")
                filename = form.getfirst("filename", "formatted-content")

            if not html_text.strip():
                self.send_json(400, {"error": "没有可导出的 HTML 内容。"})
                return

            docx_bytes = export_docx_bytes(html_text, Path.cwd())
            safe_name = re.sub(r"[^A-Za-z0-9._-]+", "-", filename).strip("-") or "formatted-content"
            if not safe_name.lower().endswith(".docx"):
                safe_name = f"{safe_name}.docx"

            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{safe_name}"',
            )
            self.end_headers()
            self.wfile.write(docx_bytes)
        except Exception as exc:  # noqa: BLE001
            self.send_json(500, {"error": f"DOCX 导出失败：{exc}"})

    def send_json(self, status: int, payload: dict[str, str]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def translate_path(self, path: str) -> str:
        path = urlparse(path).path
        path = posixpath.normpath(path)
        words = [word for word in path.split("/") if word]
        translated = str(Path.cwd())
        for word in words:
            translated = os.path.join(translated, word)
        return translated
    
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run a local web server with DOCX-to-Markdown conversion."
    )
    parser.add_argument(
        "--host",
        default=os.environ.get("HOST", "0.0.0.0"),
        help="Host to bind to. Default: HOST env var or 0.0.0.0",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("PORT", "8000")),
        help="Port to bind to. Default: PORT env var or 8000",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    os.chdir(root)

    server = ThreadingHTTPServer((args.host, args.port), AppHandler)
    display_host = args.host
    if display_host == "0.0.0.0":
        display_host = "127.0.0.1"

    url = f"http://{display_host}:{args.port}/"

    print(f"Serving folder: {root}")
    print(f"Open in browser: {url}")
    print("Press Ctrl+C to stop.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
