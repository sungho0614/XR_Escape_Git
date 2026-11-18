from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(r"C:\Unity Project\XR_Escape")
OUT = ROOT / "output" / "documents" / "표류_우주선_탈출_XR콘텐츠_기획서.docx"
SPACE_IMAGE = ROOT / "Assets" / "space.jpg"
KIT_IMAGE = ROOT / "Assets" / "Creepy_Cat" / "3D Scifi Kit Starter Kit HD - HDRP.png"

NAVY = "18324A"
PALE_BLUE = "EAF1F6"
PALE_GRAY = "F4F5F6"
MID_GRAY = "D9D9D9"
BLACK = "000000"
WHITE = "FFFFFF"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=110, start=130, bottom=110, end=130):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color=MID_GRAY, size="6"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        node = borders.find(tag)
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_text(cell, text, bold=False, color=BLACK, align=WD_ALIGN_PARAGRAPH.LEFT, size=9.5):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    r.bold = bold
    r.font.name = "Malgun Gothic"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, header in enumerate(headers):
        set_cell_shading(hdr.cells[i], NAVY)
        set_cell_text(hdr.cells[i], header, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        if widths:
            hdr.cells[i].width = Cm(widths[i])
    for row_index, row_data in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row_data):
            set_cell_text(cells[i], str(value), align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            if row_index % 2 == 1:
                set_cell_shading(cells[i], PALE_BLUE)
            if widths:
                cells[i].width = Cm(widths[i])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def add_field(paragraph, field_code):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field_code
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])


def add_body(doc, text, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.line_spacing = 1.35
    p.paragraph_format.first_line_indent = Cm(0.6)
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        r1.bold = True
        r2 = p.add_run(text[len(bold_lead):])
        for r in (r1, r2):
            r.font.name = "Malgun Gothic"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
            r.font.size = Pt(10.2)
    else:
        r = p.add_run(text)
        r.font.name = "Malgun Gothic"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
        r.font.size = Pt(10.2)
    return p


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.2
        r = p.add_run(item)
        r.font.name = "Malgun Gothic"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
        r.font.size = Pt(10)


def page_break(doc):
    doc.add_page_break()


doc = Document()
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(1.8)
section.left_margin = Cm(2.1)
section.right_margin = Cm(2.1)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Malgun Gothic"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
normal.font.size = Pt(10.2)
normal.font.color.rgb = RGBColor(0, 0, 0)

title_style = styles["Title"]
title_style.font.name = "Malgun Gothic"
title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
title_style.font.size = Pt(28)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)

for style_name, size in (("Heading 1", 18), ("Heading 2", 13), ("Heading 3", 11)):
    style = styles[style_name]
    style.font.name = "Malgun Gothic"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(8)
    style.paragraph_format.space_after = Pt(7)
    style.paragraph_format.keep_with_next = True

# Footer with page number
footer = section.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = fp.add_run("XR콘텐츠응용 중간과제 기획서  ")
fr.font.name = "Malgun Gothic"
fr._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
fr.font.size = Pt(8)
fr.font.color.rgb = RGBColor(90, 90, 90)
add_field(fp, "PAGE")

# Page 1 Cover
p = doc.add_paragraph(style="Title")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(18)
p.add_run("표류 우주선 탈출 XR 콘텐츠 기획서")

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(14)
r = sub.add_run("중간 데모 구현 결과와 기말 프로젝트 개발 계획")
r.font.name = "Malgun Gothic"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(0, 0, 0)

if SPACE_IMAGE.exists():
    pic_p = doc.add_paragraph()
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.add_run().add_picture(str(SPACE_IMAGE), width=Cm(14.8))

meta = doc.add_table(rows=4, cols=2)
meta.alignment = WD_TABLE_ALIGNMENT.CENTER
meta.autofit = False
set_table_borders(meta)
for row, (label, value) in zip(meta.rows, [
    ("교과목", "XR콘텐츠응용"),
    ("콘텐츠 제목", "표류 우주선 탈출"),
    ("학번", ""),
    ("이름", ""),
]):
    set_cell_shading(row.cells[0], PALE_BLUE)
    set_cell_text(row.cells[0], label, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row.cells[1], value, align=WD_ALIGN_PARAGRAPH.CENTER)
    row.cells[0].width = Cm(4.0)
    row.cells[1].width = Cm(10.0)

date_p = doc.add_paragraph()
date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
date_p.paragraph_format.space_before = Pt(14)
date_p.add_run("2026년 10월").font.size = Pt(10)

# Page 2 Contents and summary
page_break(doc)
doc.add_heading("문서 구성", level=1)
add_table(doc, ["번호", "내용", "핵심 확인 사항"], [
    ("1", "콘텐츠 개요", "장르와 목표 및 핵심 경험"),
    ("2", "기획 배경", "시각적 의도와 개발 필요성"),
    ("3", "레퍼런스와 차별화", "기존 콘텐츠와 구현 방식 비교"),
    ("4", "제작 재료", "에셋 이름과 출처 및 활용 위치"),
    ("5", "제작 과정", "공간과 캐릭터 및 상호작용 구현"),
    ("6", "중간 데모 구현 결과", "필수 기술 요소와 플레이 흐름"),
    ("7", "기말 프로젝트 개발 계획", "XR 확장 범위와 일정 및 위험 관리"),
], widths=[1.3, 5.1, 9.0])

doc.add_heading("기획 요약", level=2)
add_body(doc, "사고로 정지한 우주선에서 플레이어가 전력 장치와 압력 버튼을 조작하고 정비 로봇의 움직임을 활용하여 잠긴 출구를 여는 SF 방탈출 콘텐츠를 제작한다. 중간 데모에서는 Third Person 캐릭터와 NavMesh AI, 물리 오브젝트, 트리거, 조명, 파티클, 2D 및 3D 사운드를 연결해 핵심 게임 흐름을 검증했다.")
add_body(doc, "기말 프로젝트에서는 현재 데모의 구조를 XR Interaction Toolkit 기반의 손 상호작용으로 확장한다. 에너지 코어 삽입, 양손 밸브 조작, 공간 음향 추적과 같은 퍼즐을 추가하고, 각 퍼즐의 결과가 우주선 전력 복구와 탈출 포드 개방으로 이어지도록 구성할 계획이다.")

doc.add_heading("중간 데모와 기말 목표", level=2)
add_table(doc, ["구분", "중간 데모", "기말 프로젝트"], [
    ("플레이 방식", "Third Person 캐릭터로 공간 탐색", "XR 헤드셋과 컨트롤러를 이용한 직접 조작"),
    ("퍼즐 범위", "압력 버튼과 AI 이동을 이용한 문 개방", "에너지 코어와 밸브 및 공간 음향 퍼즐"),
    ("콘텐츠 범위", "단일 우주선 씬과 짧은 탈출 과정", "복수 구역과 단계형 전력 복구 과정"),
    ("검증 목표", "수업 필수 기술의 작동과 연결", "XR 환경에서의 조작성과 몰입감 개선"),
], widths=[2.5, 6.5, 6.5])

# Page 3 Overview
page_break(doc)
doc.add_heading("1 콘텐츠 개요", level=1)
doc.add_heading("1.1 콘텐츠 설정", level=2)
add_body(doc, "플레이어는 항로를 이탈해 정지한 소형 우주선 내부에서 깨어난다. 주 전원은 차단되어 있으며 출입문과 탈출 포드가 잠겨 있다. 플레이어는 비상 조명과 기계음이 남아 있는 내부를 탐색하고, 이동 가능한 물체와 압력 버튼, 정비 로봇을 이용해 전력을 단계적으로 복구한다.")
add_body(doc, "중간 데모의 목표는 모든 기능을 완성하는 것이 아니라 핵심 상호작용이 하나의 게임 경험으로 연결되는지 확인하는 것이다. 플레이어 조작, AI의 자동 이동, 두 개 버튼의 동시 작동, 문 비활성화, 연기와 폭발 효과, 사운드 재생을 한 장면 안에서 확인할 수 있도록 제작했다.")

doc.add_heading("1.2 목표 사용자 경험", level=2)
add_bullets(doc, [
    "우주 Skybox와 밀폐형 선실 구조를 통해 현재 위치와 위기 상황을 즉시 이해한다.",
    "빛과 소리를 단서로 삼아 상호작용 가능한 장치의 위치를 추론한다.",
    "플레이어와 AI가 서로 다른 버튼을 담당하면서 환경을 이용한 해결 방식을 경험한다.",
    "문이 열리고 파티클과 사운드가 재생되는 결과를 통해 퍼즐 해결 여부를 명확히 확인한다.",
])

doc.add_heading("1.3 플레이 흐름", level=2)
add_table(doc, ["단계", "플레이 행동", "시스템 반응"], [
    ("1", "우주선 내부 탐색", "FreeLook 카메라가 플레이어를 추적하고 배경음이 재생됨"),
    ("2", "버튼과 정비 로봇 위치 확인", "NavMesh AI가 지정 경로를 따라 이동함"),
    ("3", "플레이어와 AI가 버튼을 동시에 누름", "DualButtonManager가 두 버튼의 상태를 확인함"),
    ("4", "조건 충족", "좌우 문이 사라지고 연기 및 폭발 이펙트가 활성화됨"),
    ("5", "열린 구역으로 이동", "탈출 구역과 다음 확장 퍼즐로 진행함"),
], widths=[1.4, 7.0, 8.0])

# Page 4 Background
page_break(doc)
doc.add_heading("2 기획 배경", level=1)
doc.add_heading("2.1 기획 의도", level=2)
add_body(doc, "시각적 의도는 기본 Cube로 만든 기능 시험 공간을 실제 우주선 내부처럼 보이게 만드는 데 있다. 금속 벽과 발광 패널, 창밖 우주 이미지, 제한된 조명을 조합하여 정전된 선실의 분위기를 표현했다. 버튼 작동 전후에는 문, 조명, 연기와 폭발 효과가 달라지도록 구성하여 플레이어의 행동이 공간에 영향을 주는 모습을 강조했다.")
add_body(doc, "게임 기획에서는 XR에서 가능한 신체적 행동을 퍼즐의 규칙으로 사용하는 것을 목표로 했다. 현재 데모는 키보드 기반 Third Person 방식이지만, 압력판에 올라가기, 물체를 옮기기, 서로 떨어진 장치를 동시에 작동하기와 같은 구조는 기말 단계에서 손과 몸을 사용하는 XR 상호작용으로 전환하기 쉽다.")

doc.add_heading("2.2 개발 필요성", level=2)
add_body(doc, "방탈출 장르는 사용자가 공간을 관찰하고 원인을 추론한 뒤 직접 행동한다는 점에서 XR 매체와 잘 맞는다. 단순히 화면의 버튼을 선택하는 대신 사용자가 손으로 장치를 조작하고, 소리가 들리는 방향으로 이동하며, 다른 위치의 장치를 동시에 다루게 하면 공간 자체가 퍼즐 인터페이스가 된다.")
add_body(doc, "수업에서 다룬 Standard Assets와 NavMesh, Collider, Rigidbody, Trigger, 2D 및 3D 사운드를 하나의 목표 아래 연결할 수 있다는 점도 개발 이유다. 각 기능을 독립적으로 배치하지 않고 우주선 전력 복구라는 상황에 포함하면 구현 요소의 역할을 설명하기 쉽고 기말 프로젝트로 확장할 구조도 확보할 수 있다.")

doc.add_heading("2.3 개발 범위", level=2)
add_table(doc, ["범위", "포함 내용", "제외 또는 이후 개발"], [
    ("중간", "우주선 단일 씬과 플레이어 및 AI 이동, 버튼 퍼즐, 문과 이펙트, 사운드", "실제 XR 입력과 복수 스테이지"),
    ("기말", "XR 손 조작, 에너지 코어와 밸브 퍼즐, 공간 음향 단서, 탈출 포드", "온라인 멀티플레이는 선택 확장"),
], widths=[2.0, 8.0, 6.4])

# Page 5 Reference and differentiation
page_break(doc)
doc.add_heading("3 레퍼런스와 차별화 요소", level=1)
doc.add_heading("3.1 레퍼런스 콘텐츠", level=2)
add_table(doc, ["레퍼런스", "참고한 요소", "본 콘텐츠의 적용 방식"], [
    ("The Room VR A Dark Matter", "공간 관찰과 장치 조작을 중심으로 한 VR 퍼즐", "우주선 장치와 빛 및 소리를 단서로 사용"),
    ("I Expect You To Die", "밀폐된 공간에서 연속 행동으로 위기를 해결하는 구조", "하나의 선실 안에서 전력 복구와 탈출을 연결"),
    ("Keep Talking and Nobody Explodes", "서로 다른 역할과 동시 행동이 필요한 협동 규칙", "플레이어와 AI가 두 압력 버튼을 동시에 담당"),
], widths=[4.3, 6.0, 6.1])

doc.add_heading("3.2 기획 차별화", level=2)
add_body(doc, "본 콘텐츠는 정답 번호를 입력하는 방식보다 공간 안의 물체와 캐릭터가 조건을 만족해야 다음 상태로 전환되는 퍼즐을 중심에 둔다. 특히 정비 로봇은 배경 캐릭터가 아니라 압력 버튼을 담당하는 퍼즐 요소로 사용된다. 플레이어는 AI의 이동 경로를 관찰하고 자신이 담당할 버튼을 선택해야 한다.")
add_body(doc, "기말 단계에서는 두 버튼 구조를 XR의 양손 조작과 공간 분할 퍼즐로 확장한다. 한 손으로 전력 레버를 유지한 상태에서 다른 손으로 에너지 코어를 삽입하거나, 소리가 강해지는 방향을 따라 손전등과 장치를 정렬하는 방식으로 발전시킬 계획이다.")

doc.add_heading("3.3 구현 차별화", level=2)
add_bullets(doc, [
    "DualButtonManager가 두 Trigger의 상태를 동시에 검사하여 단일 버튼 방식과 다른 조건형 퍼즐을 구성한다.",
    "NavMesh AI의 이동 결과가 문 개방 조건에 포함되어 캐릭터 이동 기능과 퍼즐 진행이 연결된다.",
    "문 비활성화와 파티클 활성화를 분리하여 하나의 조건에서 서로 반대되는 공간 변화가 발생한다.",
    "DisableTriggerAfterUse를 이용해 일회성 장치가 반복 실행되며 오류를 만드는 상황을 방지한다.",
])

# Page 6 Materials
page_break(doc)
doc.add_heading("4 제작에 활용된 재료", level=1)
doc.add_heading("4.1 주요 에셋", level=2)
add_table(doc, ["에셋과 출처", "활용 위치", "활용 내용"], [
    ("Unity Standard Assets", "플레이어와 AI 및 카메라", "ThirdPersonController, AIThirdPersonController, FreeLookCameraRig, Particle System 활용"),
    ("3D Sci Fi Kit Starter Kit HD 제작사 Creepy Cat", "우주선 내부", "벽, 문, 바닥, 창문과 발광 재질을 이용한 선실 구성"),
    ("Polygon Starter", "보조 모델", "테스트용 오브젝트와 공간 장식에 활용"),
    ("space.jpg 프로젝트 이미지", "Skybox와 우주 배경", "선실 외부가 우주 공간임을 전달하는 배경 이미지"),
    ("XR Sounds 프로젝트 폴더", "배경음과 상호작용 사운드", "loop, ding, explosion, smoke 계열 음원을 장면 연출에 활용"),
], widths=[5.3, 4.0, 7.1])

image_table = doc.add_table(rows=1, cols=2)
image_table.alignment = WD_TABLE_ALIGNMENT.CENTER
image_table.autofit = False
set_table_borders(image_table)
for cell in image_table.rows[0].cells:
    set_cell_margins(cell, 140, 140, 140, 140)
if SPACE_IMAGE.exists():
    p1 = image_table.cell(0, 0).paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run().add_picture(str(SPACE_IMAGE), width=Cm(7.2))
    c1 = image_table.cell(0, 0).add_paragraph("우주 배경 이미지")
    c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
if KIT_IMAGE.exists():
    p2 = image_table.cell(0, 1).paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run().add_picture(str(KIT_IMAGE), width=Cm(7.2))
    c2 = image_table.cell(0, 1).add_paragraph("3D Sci Fi Kit Starter Kit HD")
    c2.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_heading("4.2 직접 제작한 구성 요소", level=2)
add_body(doc, "에셋을 그대로 배치하는 대신 우주선 실내 구조와 버튼 위치, AI 이동 경로, 문 개방 조건을 직접 구성했다. PressureButton과 DualButtonManager를 추가하여 플레이어와 AI가 동시에 상호작용해야 하는 조건을 만들었고, 반복 발동으로 파티클 참조 오류가 발생하지 않도록 일회성 Trigger 제어 스크립트도 제작했다.")

# Page 7 Process
page_break(doc)
doc.add_heading("5 제작 과정", level=1)
doc.add_heading("5.1 공간 제작", level=2)
add_body(doc, "새 프로젝트의 SampleScene을 기준으로 Plane과 Cube를 이용해 이동 가능 영역과 충돌 벽을 먼저 만들었다. 이후 SF Kit의 벽과 문, 창문 모델을 배치하고 우주 배경이 창밖으로 보이도록 Skybox를 설정했다. Directional Light와 여러 Point Light를 사용해 선실의 기본 조명과 장치별 강조 조명을 구분했다.")

doc.add_heading("5.2 캐릭터와 카메라", level=2)
add_body(doc, "Standard Assets의 ThirdPersonController를 플레이어로 배치하고 FreeLookCameraRig가 캐릭터를 자연스럽게 추적하도록 연결했다. AIThirdPersonController에는 NavMeshAgent와 순찰 지점을 설정해 목표 위치로 이동하게 했다. NavMesh의 Stopping Distance와 버튼 Collider 범위를 조정하여 AI가 Trigger 영역까지 실제로 진입하도록 맞췄다.")

doc.add_heading("5.3 상호작용 제작", level=2)
add_table(doc, ["구현 순서", "문제와 조정 내용"], [
    ("Trigger 기본 동작", "Player 태그에만 반응하던 구조를 오브젝트와 AI 상호작용까지 확장"),
    ("두 버튼 조건", "PressureButton 두 개가 같은 DualButtonManager에 0번과 1번 상태를 전달하도록 구성"),
    ("문과 이펙트 분리", "문은 SetActive false, 이펙트는 SetActive true로 처리하도록 배열 분리"),
    ("반복 실행 방지", "일회성 버튼은 첫 실행 다음 프레임에 Collider와 관련 Behaviour를 비활성화"),
    ("파티클 수명 조정", "자동 Destroy 구성과 반복 활성화의 충돌을 확인하고 유지 또는 일회성 실행 방식으로 정리"),
], widths=[4.0, 12.4])

doc.add_heading("5.4 제작 과정에서 확인한 어려움", level=2)
add_body(doc, "Trigger는 충돌하는 오브젝트 중 적어도 하나에 Rigidbody가 있어야 정상적으로 이벤트를 전달한다. 또한 NavMeshAgent가 버튼 중심보다 앞에서 멈추면 Collider가 겹치지 않아 코드가 정상이어도 버튼 상태가 변경되지 않았다. 구현 과정에서는 스크립트 조건뿐 아니라 Collider의 실제 범위와 Rigidbody, Stopping Distance를 함께 확인했다.")

# Page 8 Implemented tech
page_break(doc)
doc.add_heading("6 중간 데모 구현 결과", level=1)
doc.add_heading("6.1 필수 기술 구현", level=2)
add_table(doc, ["평가 요소", "구현 내용", "영상에서 보여줄 장면"], [
    ("공간 구성", "우주선 벽과 바닥 및 문에 Collider 적용", "캐릭터가 벽과 바닥을 통과하지 않는 장면"),
    ("물리", "Rigidbody를 가진 플레이어와 AI 및 상호작용 오브젝트", "중력과 충돌 또는 버튼 진입 장면"),
    ("분위기", "우주 Skybox와 Point Light 및 발광 재질", "창밖 우주와 선실 조명 전경"),
    ("카메라", "FreeLookCameraRig가 플레이어 추적", "이동 중 카메라 회전과 추적"),
    ("플레이어", "ThirdPersonController 이동", "키보드 이동과 점프"),
    ("AI", "NavMeshAgent 기반 AIThirdPersonController", "정비 로봇이 목표 버튼으로 이동"),
    ("트리거", "두 버튼 동시 입력과 일회성 버튼", "문과 이펙트가 조건에 따라 변화"),
    ("사운드", "2D 배경음과 3D 상호작용 음원", "배경음과 거리 기반 효과음"),
], widths=[2.8, 7.6, 6.0])

doc.add_heading("6.2 완성도와 현재 한계", level=2)
add_body(doc, "현재 데모는 우주선 내부 탐색에서 버튼 퍼즐과 문 개방까지 이어지는 최소 게임 흐름을 완성했다. 공간, 캐릭터, AI, 조명, 사운드와 Trigger가 같은 목표 안에서 작동하므로 기능 시험용 배치보다 하나의 콘텐츠에 가까운 구조를 갖췄다.")
add_body(doc, "다만 현재 조작은 Third Person 방식이며 실제 XR 컨트롤러 입력은 적용하지 않았다. 퍼즐 수도 제한되어 있고 성공 UI와 실패 조건, 사용자 안내가 충분하지 않다. 이 한계는 기말 개발에서 XR 손 상호작용과 단계형 퍼즐, 상태 안내 UI를 추가하는 방향으로 보완한다.")

# Page 9 Final plan
page_break(doc)
doc.add_heading("7 기말 프로젝트 개발 계획", level=1)
doc.add_heading("7.1 개발 목표", level=2)
add_body(doc, "기말 프로젝트의 목표는 현재 데모의 조건형 퍼즐 구조를 실제 XR 조작으로 전환하고, 우주선의 전력 복구 과정이 여러 구역을 연결하도록 확장하는 것이다. 플레이어가 손으로 잡고 돌리고 삽입하는 행동을 수행하며, 조명과 사운드 변화로 진행 상태를 이해할 수 있도록 한다.")

doc.add_heading("7.2 추가 기능", level=2)
add_table(doc, ["우선순위", "기능", "구현 계획"], [
    ("필수", "XR 이동과 손 상호작용", "XR Interaction Toolkit을 적용하고 Grab 및 Socket 상호작용 구성"),
    ("필수", "에너지 코어 퍼즐", "Rigidbody 코어를 전력 장치 Socket에 삽입하면 조명과 문 상태 변경"),
    ("필수", "양손 밸브 퍼즐", "두 손의 위치와 회전을 이용해 비상 밸브 조작"),
    ("권장", "공간 음향 퍼즐", "3D 사운드의 방향과 거리를 이용해 고장 장치 위치 탐색"),
    ("권장", "진행 UI", "산소와 전력 상태 및 현재 목표를 월드 스페이스 UI로 표시"),
    ("선택", "협동 모드", "두 사용자가 서로 다른 구역의 장치를 동시에 조작"),
], widths=[2.0, 4.7, 9.7])

doc.add_heading("7.3 개발 일정", level=2)
add_table(doc, ["단계", "주요 작업", "완료 기준"], [
    ("1주차", "XR Rig와 입력 환경 구축", "헤드셋 이동과 양손 컨트롤러 입력 작동"),
    ("2주차", "잡기와 Socket 상호작용", "에너지 코어를 집고 지정 위치에 삽입 가능"),
    ("3주차", "추가 퍼즐과 공간 음향", "밸브와 음향 퍼즐이 전력 상태와 연결"),
    ("4주차", "연출과 UI 및 사운드", "목표와 성공 상태를 사용자가 명확히 파악"),
    ("5주차", "테스트와 최적화 및 영상 제작", "주요 버그 제거와 전체 플레이 녹화"),
], widths=[2.2, 7.2, 7.0])

doc.add_heading("7.4 위험 요소와 대응", level=2)
add_table(doc, ["위험 요소", "대응 방법"], [
    ("XR 장비 테스트 시간이 부족함", "데스크톱 XR Device Simulator로 먼저 검증하고 실제 장비 테스트 항목을 별도로 관리"),
    ("물리 오브젝트가 튀거나 끼임", "질량과 Collider 크기 및 Fixed Timestep을 조정하고 Reset 지점 제공"),
    ("퍼즐 조건이 복잡해짐", "각 퍼즐 상태를 독립 Manager로 분리하고 Inspector에서 참조를 명확히 연결"),
    ("멀미와 조작성 문제", "이동 속도를 제한하고 Snap Turn과 Teleport 옵션을 우선 제공"),
], widths=[5.0, 11.4])

doc.add_heading("7.5 제출 전 확인", level=2)
add_bullets(doc, [
    "실행 영상에서 공간, 물리, 조명, 카메라, 플레이어, AI, Trigger, 2D 및 3D 사운드를 모두 보여준다.",
    "사용한 외부 에셋의 이름과 제작사 및 활용 위치를 최종 문서에 다시 확인한다.",
    "Unity Package를 새 프로젝트에서 열어 누락된 참조와 컴파일 오류가 없는지 확인한다.",
    "기말 단계에서는 중간 데모와 XR 확장 기능을 별도 씬 또는 브랜치로 관리한다.",
])

OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(OUT)
