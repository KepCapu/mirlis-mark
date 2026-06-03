#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка PDF-руководства из Markdown. Запуск: python3 build_pdf.py"""
import re
import pathlib
import markdown
from bs4 import BeautifulSoup
from weasyprint import HTML

BASE = pathlib.Path(__file__).resolve().parent
MD = BASE / "Руководство_пользователя.md"
PDF = BASE / "Руководство_пользователя.pdf"

md_text = MD.read_text(encoding="utf-8")
html_body = markdown.markdown(
    md_text,
    extensions=["extra", "sane_lists", "smarty"],
    output_format="html5",
)

soup = BeautifulSoup(html_body, "html.parser")

# 0) Проставляем id заголовкам (GitHub-style слаги): без них ссылки
#    содержания никуда не ведут, и target-counter не может взять номер страницы.
def _gh_slug(text):
    s = text.strip().lower().replace("«", "").replace("»", "")
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "-", s)
    return s

_seen = {}
for _h in soup.find_all(["h1", "h2", "h3", "h4"]):
    _base = _gh_slug(_h.get_text())
    if _base in _seen:
        _seen[_base] += 1
        _h["id"] = f"{_base}-{_seen[_base]}"
    else:
        _seen[_base] = 0
        _h["id"] = _base

# Помечаем список содержания классом toc (номера страниц с отточием).
_toc_h = soup.find(lambda t: t.name == "h2" and "Содержание" in t.get_text())
if _toc_h:
    _toc_ol = _toc_h.find_next("ol")
    if _toc_ol:
        _toc_ol["class"] = _toc_ol.get("class", []) + ["toc"]
        # markdown не вложил подпункты в <ul>: они пришли как литералы
        # "- <a>…</a>" внутри родительского <li>. Превращаем в настоящий
        # вложенный список и убираем текстовые узлы-«дефисы».
        for _li in list(_toc_ol.find_all("li", recursive=False)):
            _links = _li.find_all("a", recursive=False)
            if len(_links) > 1:
                _nested = soup.new_tag("ul")
                for _a in _links[1:]:
                    _a.extract()
                    _li2 = soup.new_tag("li")
                    _li2.append(_a)
                    _nested.append(_li2)
                for _child in list(_li.children):
                    if isinstance(_child, str) and set(_child.strip()) <= {"-"}:
                        _child.extract()
                _li.append(_nested)

# 1) Абзац с единственной картинкой -> <figure> с подписью из alt.
for p in soup.find_all("p"):
    imgs = p.find_all("img")
    if len(imgs) == 1 and not p.get_text(strip=True):
        img = imgs[0]
        fig = soup.new_tag("figure")
        img.extract()
        fig.append(img)
        alt = img.get("alt", "").strip()
        if alt:
            cap = soup.new_tag("figcaption")
            cap.string = alt
            fig.append(cap)
        p.replace_with(fig)

# 2) Группируем заголовок (h2/h3/h4) вместе с идущими следом
#    абзацами/таблицами/списками/цитатами — до ближайшего заголовка,
#    рисунка (figure) или разделителя. Так название НЕ отрывается от
#    своего описания и таблицы. Рисунки остаются вне группы и могут
#    свободно переноситься на новую страницу.
HEAD = {"h1", "h2", "h3", "h4"}
STOP = HEAD | {"figure", "hr"}
GLUE = {"p", "table", "ul", "ol", "blockquote"}

top = list(soup.children)
i = 0
while i < len(top):
    node = getattr(top[i], "name", None)
    if node in {"h2", "h3", "h4"}:
        group = [top[i]]
        j = i + 1
        while j < len(top):
            nm = getattr(top[j], "name", None)
            if nm in STOP:
                break
            if nm in GLUE:
                group.append(top[j])
                j += 1
            else:
                # пробелы между тегами
                group.append(top[j])
                j += 1
        # оборачиваем только если в группе есть что-то кроме заголовка
        meaningful = [g for g in group[1:] if getattr(g, "name", None) in GLUE]
        if meaningful:
            wrapper = soup.new_tag("div", **{"class": "keep"})
            group[0].insert_before(wrapper)
            for g in group:
                wrapper.append(g.extract())
            top = list(soup.children)
            i = top.index(wrapper) + 1
            continue
    i += 1

html_body = str(soup)

CSS = """
@page {
  size: A4;
  margin: 18mm 16mm 20mm 16mm;
  @bottom-center {
    content: "Mark · Руководство пользователя · стр. " counter(page) " из " counter(pages);
    font-family: "DejaVu Sans"; font-size: 8pt; color: #8a8f98;
  }
}
* { box-sizing: border-box; }
body {
  font-family: "DejaVu Sans", sans-serif;
  font-size: 10.5pt; line-height: 1.5; color: #1f2430;
}
h1 {
  font-size: 22pt; color: #11151c; margin: 0 0 4pt;
  border-bottom: 3px solid #f5b301; padding-bottom: 6pt;
}
h2 {
  font-size: 15pt; color: #11151c; margin: 22pt 0 8pt;
  padding-left: 8pt; border-left: 4px solid #f5b301;
  page-break-after: avoid;
}
h3 {
  font-size: 12.5pt; color: #2a3140; margin: 16pt 0 6pt;
  page-break-after: avoid;
}
h4 {
  font-size: 11pt; color: #2a3140; margin: 13pt 0 4pt;
  page-break-after: avoid;
}
blockquote {
  margin: 8pt 0; padding: 6pt 12pt; background: #faf3dd;
  border-left: 4px solid #f5b301; border-radius: 0 4px 4px 0;
  font-size: 9.8pt; color: #4a4f59; page-break-inside: avoid;
}
blockquote p { margin: 0; }
p { margin: 6pt 0; }
strong { color: #11151c; }
ul, ol { margin: 6pt 0 6pt 0; padding-left: 20pt; }
li { margin: 3pt 0; }
a { color: #1f2430; text-decoration: none; }

/* Содержание: номера страниц с отточием (точками) */
ol.toc { padding-left: 24pt; }
ol.toc li { margin: 3.5pt 0; }
ol.toc a {
  display: block; color: #1f2430; text-decoration: none;
}
ol.toc a::after {
  content: leader(". ") target-counter(attr(href url), page);
  color: #8a8f98;
}
ol.toc ul { list-style: none; padding-left: 18pt; margin: 2.5pt 0; }
ol.toc ul a { color: #4a4f59; font-size: 9.8pt; }
ol.toc ul a::after { color: #aab0ba; }
hr { border: 0; border-top: 1px solid #e3e6ea; margin: 16pt 0; }
code { font-family: "DejaVu Sans Mono"; font-size: 9pt;
       background: #f2f3f5; padding: 0 3px; border-radius: 3px; }

/* Заголовок + его описание/таблица держим вместе */
.keep { page-break-inside: avoid; }

figure {
  margin: 12pt auto; text-align: center; page-break-inside: avoid;
}
figure img {
  max-width: 100%; max-height: 175mm;
  border: 1px solid #e3e6ea; border-radius: 6px;
}
figcaption {
  font-size: 8.5pt; color: #8a8f98; margin-top: 4pt; font-style: italic;
}
table {
  border-collapse: collapse; margin: 10pt 0; width: 100%;
  page-break-inside: avoid; font-size: 9.8pt;
}
th, td {
  border: 1px solid #e3e6ea; padding: 5pt 9pt; text-align: left;
  vertical-align: top;
}
th { background: #faf3dd; }
td img { max-width: 100%; height: auto; border: 1px solid #e3e6ea; border-radius: 4px; }
td:first-child, th:first-child { white-space: nowrap; }
"""

full_html = f"""<!doctype html><html lang="ru"><head><meta charset="utf-8">
<style>{CSS}</style></head><body>{html_body}</body></html>"""

(BASE / "_preview.html").write_text(full_html, encoding="utf-8")
HTML(string=full_html, base_url=str(BASE)).write_pdf(str(PDF))
print("PDF собран:", PDF)
