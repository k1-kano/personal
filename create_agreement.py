"""
裁量労働制から普通勤務制（みなし残業45時間）への変更に関する合意書を作成する。
"""

from docx import Document
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_japanese_font(run, size=10.5, bold=False):
    run.font.name = "ＭＳ 明朝"
    run.font.size = Pt(size)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "ＭＳ 明朝")
    rFonts.set(qn("w:ascii"), "ＭＳ 明朝")
    rFonts.set(qn("w:hAnsi"), "ＭＳ 明朝")


def add_para(doc, text="", *, size=10.5, bold=False, align=None, indent=None, space_after=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    if indent is not None:
        pf.first_line_indent = Pt(indent)
    if space_after is not None:
        pf.space_after = Pt(space_after)
    run = p.add_run(text)
    set_japanese_font(run, size=size, bold=bold)
    return p


def add_article_heading(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_japanese_font(run, size=10.5, bold=True)
    return p


def build():
    doc = Document()

    # ページ余白
    for section in doc.sections:
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(25)
        section.right_margin = Mm(25)

    # 既定スタイル
    style = doc.styles["Normal"]
    style.font.name = "ＭＳ 明朝"
    style.font.size = Pt(10.5)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "ＭＳ 明朝")

    # タイトル
    add_para(
        doc,
        "労働条件変更に関する合意書",
        size=16,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=12,
    )

    # 前文
    add_para(
        doc,
        "株式会社○○○○（以下「甲」という。）と○○○○（以下「乙」という。）"
        "とは、乙の勤務制度の変更に関し、以下のとおり合意する。",
        indent=10.5,
        space_after=6,
    )

    # 第1条（勤務制度の変更）
    add_article_heading(doc, "第1条（勤務制度の変更）")
    add_para(
        doc,
        "甲及び乙は、現在乙に適用されている専門業務型裁量労働制（又は企画業務型裁量労働制）"
        "を廃止し、第2条に定める所定労働時間を基準とする通常勤務制度に変更することに合意する。",
        indent=10.5,
    )

    # 第2条（所定労働時間及び勤務時間）
    add_article_heading(doc, "第2条（所定労働時間及び勤務時間）")
    add_para(
        doc,
        "1. 乙の所定労働時間は、1日8時間、1週40時間とする。",
        indent=10.5,
    )
    add_para(
        doc,
        "2. 乙の始業時刻は午前9時00分、終業時刻は午後6時00分とし、休憩時間は正午から午後1時00分までの60分とする。",
        indent=10.5,
    )
    add_para(
        doc,
        "3. 休日は、土曜日、日曜日、国民の祝日に関する法律に定める休日、年末年始（12月29日から翌年1月3日まで）"
        "その他甲が別途定める日とする。",
        indent=10.5,
    )

    # 第3条（固定残業代／みなし残業手当）
    add_article_heading(doc, "第3条（固定残業手当）")
    add_para(
        doc,
        "1. 甲は乙に対し、時間外労働及び深夜労働の対価として、毎月、固定残業手当（以下「本手当」という。）"
        "を支給する。本手当は、月45時間分の時間外労働に対する割増賃金に相当する額として定額で支給する。",
        indent=10.5,
    )
    add_para(
        doc,
        "2. 本手当の額は、金○○○，○○○円（月額）とする。",
        indent=10.5,
    )
    add_para(
        doc,
        "3. 1か月の時間外労働が45時間を超えた場合、甲は超過部分に相当する割増賃金を別途支払う。"
        "また、法定休日労働及び深夜労働（22時から翌5時まで）に対する割増賃金は、本手当に含まれず、別途支給する。",
        indent=10.5,
    )
    add_para(
        doc,
        "4. 1か月の実際の時間外労働が45時間に満たない場合であっても、本手当の額は減額しないものとする。",
        indent=10.5,
    )

    # 第4条（賃金）
    add_article_heading(doc, "第4条（賃金）")
    add_para(
        doc,
        "本合意による変更後の乙の賃金は、別紙「賃金内訳」のとおりとし、"
        "基本給、諸手当及び前条の固定残業手当により構成される。",
        indent=10.5,
    )

    # 第5条（時間外労働及び休日労働に関する協定）
    add_article_heading(doc, "第5条（時間外労働及び休日労働）")
    add_para(
        doc,
        "時間外労働及び休日労働は、労働基準法第36条に基づき甲が締結し所轄労働基準監督署長に届け出た"
        "労使協定（いわゆる36協定）の範囲内で、甲の業務上の必要に応じて命じることがあり、乙はこれに従うものとする。",
        indent=10.5,
    )

    # 第6条（労働時間の管理）
    add_article_heading(doc, "第6条（労働時間の管理）")
    add_para(
        doc,
        "乙は、甲の定める勤怠管理の方法に従い、始業・終業時刻、休憩時間及び時間外労働時間を正確に記録するものとする。",
        indent=10.5,
    )

    # 第7条（適用開始日）
    add_article_heading(doc, "第7条（適用開始日）")
    add_para(
        doc,
        "本合意は、2026年5月1日より適用する。",
        indent=10.5,
    )

    # 第8条（その他の労働条件）
    add_article_heading(doc, "第8条（その他の労働条件）")
    add_para(
        doc,
        "本合意に定めのない事項については、甲の就業規則、賃金規程その他関係諸規程の定めるところによる。"
        "本合意の内容と就業規則等の定めとの間に齟齬が生じた場合は、本合意の内容が優先するものとする。",
        indent=10.5,
    )

    # 第9条（協議事項）
    add_article_heading(doc, "第9条（協議事項）")
    add_para(
        doc,
        "本合意の解釈について疑義が生じた場合、又は本合意に定めのない事項については、甲乙誠実に協議の上これを解決する。",
        indent=10.5,
        space_after=12,
    )

    # 結語
    add_para(
        doc,
        "以上、本合意の成立を証するため本書2通を作成し、甲乙記名押印の上、各自1通を保有する。",
        indent=10.5,
        space_after=18,
    )

    # 日付
    add_para(
        doc,
        "2026年　　月　　日",
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        space_after=12,
    )

    # 署名欄（テーブル）
    table = doc.add_table(rows=2, cols=2)
    table.autofit = True
    widths = [Cm(2.5), Cm(13.0)]
    rows_data = [
        ("（甲）", "所在地：\n名　称：株式会社○○○○\n代表者：代表取締役　○○　○○　　　　　　印"),
        ("（乙）", "住　所：\n氏　名：○○　○○　　　　　　　　　　　　印"),
    ]
    for row_idx, (label, content) in enumerate(rows_data):
        cells = table.rows[row_idx].cells
        cells[0].width = widths[0]
        cells[1].width = widths[1]

        for cell in cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            # セル内の既定段落を消去してから追記
            cell.text = ""

        p0 = cells[0].paragraphs[0]
        run0 = p0.add_run(label)
        set_japanese_font(run0, size=10.5, bold=True)

        # 改行込みのテキスト
        lines = content.split("\n")
        first = True
        for line in lines:
            if first:
                p = cells[1].paragraphs[0]
                first = False
            else:
                p = cells[1].add_paragraph()
            run = p.add_run(line)
            set_japanese_font(run, size=10.5)

    # 表の罫線を消す
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{border_name}")
        b.set(qn("w:val"), "nil")
        tblBorders.append(b)
    tblPr.append(tblBorders)

    # 改ページ後に賃金内訳の別紙
    doc.add_page_break()
    add_para(
        doc,
        "別紙　賃金内訳",
        size=14,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=12,
    )
    add_para(doc, "適用開始日：2026年5月1日", space_after=6)

    salary_table = doc.add_table(rows=6, cols=2)
    salary_table.style = "Table Grid"
    salary_rows = [
        ("項目", "金額（円）"),
        ("基本給", ""),
        ("役職手当・その他手当", ""),
        ("固定残業手当（時間外労働45時間相当分）", ""),
        ("通勤手当", "実費支給"),
        ("合計（月額）", ""),
    ]
    for i, (k, v) in enumerate(salary_rows):
        row = salary_table.rows[i]
        for j, val in enumerate((k, v)):
            cell = row.cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            if i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            set_japanese_font(run, size=10.5, bold=(i == 0))

    add_para(doc, "")
    add_para(
        doc,
        "※固定残業手当は、月45時間分の時間外労働に対する割増賃金として支給する定額手当である。"
        "実際の時間外労働が当該時間数を超過した場合は、超過分の割増賃金を別途支給する。",
        size=9,
    )
    add_para(
        doc,
        "※深夜労働（22時～翌5時）及び法定休日労働に対する割増賃金は、上記固定残業手当に含まれず、別途支給する。",
        size=9,
    )

    out_path = "/home/user/personal/労働条件変更合意書_裁量労働制から通常勤務制への変更.docx"
    doc.save(out_path)
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    build()
