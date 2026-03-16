from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
import os

# win32 (pywin32) is optional — used only for accurate page counts when available
try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except Exception:
    win32 = None
    WIN32_AVAILABLE = False

# ----------------------
# Word 페이지 수 확인
# ----------------------
def get_page_count(doc_path):
    if not WIN32_AVAILABLE:
        raise RuntimeError("win32com is not available on this system; cannot compute page count.")

    # Ensure doc_path is absolute
    abs_doc_path = os.path.abspath(doc_path)

    # Initialize Word application FIRST, before opening documents
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    # Open the document only once, using absolute path
    doc = word.Documents.Open(abs_doc_path)
    doc.Repaginate()
    pages = doc.ComputeStatistics(2)  # 2: wdStatisticPages

    doc.Close(False)
    word.Quit()

    return pages

def set_document_font(doc, font_name):
    """문서 전체 글꼴 변경"""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(12)

    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            for run in p.runs:
                run.font.name = font_name
                run.font.size = Pt(12)

# ----------------------
# 페이지 번호 추가
# ----------------------
def add_page_number(header, name=None):
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run(f"{name} " if name else "")

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

    # Ensure header font consistent
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)


def apply_common_styles(doc):
    """Apply common Normal style settings to the document."""
    if 'Normal' in doc.styles:
        style = doc.styles['Normal']
        set_document_font(doc, "Cambria")
        font = style.font
        font.name = "Times New Roman"
        font.size = Pt(12)
        pformat = style.paragraph_format
        pformat.line_spacing = 2
        pformat.space_before = Pt(0)
        pformat.space_after = Pt(0)
        pformat.first_line_indent = Inches(0.5)

# ----------------------
# 문서 포맷팅
# ----------------------
def mla_format(input_file, output_file, title, student_name, professor_name, course_name, page_limit=None):

    doc = Document(input_file)

    # 섹션 마진 설정 (기본 MLA: 1-inch margins)
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    # 페이지 헤더 (MLA: 성 + 페이지)
    header = section.header
    add_page_number(header, student_name)

    # 공통 스타일 적용
    apply_common_styles(doc)

    # 문서 맨 앞 정보 삽입 (MLA 스타일 상단 왼쪽에 배치)
    header_info = [
        student_name,
        professor_name,
        course_name,
        datetime.today().strftime("%d %B %Y")
    ]
    for text in reversed(header_info):
        p = doc.paragraphs[0].insert_paragraph_before(text)

    # 제목 추가 (가운데 정렬)
    title_paragraph = doc.paragraphs[4].insert_paragraph_before(title)
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 문단 들여쓰기
    for p in doc.paragraphs:
        if p.text.strip() != "":
            p.paragraph_format.first_line_indent = Inches(0.5)

    # 페이지 제한 체크
    if page_limit is not None:

        doc.save(output_file)

        pages = get_page_count(output_file)

        if pages > page_limit:
            print(f"Page limit exceeded ({pages} > {page_limit})")
            print("Switching font from Cambria to Times New Roman...")

            set_document_font(doc, "Times New Roman")

            doc.save(output_file)
            pages = get_page_count(output_file)

            if pages > page_limit:
                print("⚠ WARNING: Even with Times New Roman the document exceeds the page limit.")
            else:
                print("Page count fixed by switching to Times New Roman.")
        else:
            print(f"Pages within limit ({pages} <= {page_limit})")

    # 최종 저장
    doc.save(output_file)
    print(f"MLA-formatted document saved as {output_file}")

# ----------------------
# 유저 입력
# ----------------------
def main():
    print("=== Word Formatter ===")
    style = input("Choose style (MLA/Chicago/APA/Harvard/IEEE) [MLA]: ").strip().lower() or "mla"
    title = input("Enter your essay title: ")
    student_name = input("Enter your name (optional for Chicago): ")
    professor_name = input("Enter professor's name: ")
    course_name = input("Enter course name: ")
    page_limit_input = input("Enter page limit (or press Enter to skip): ")
    page_limit = int(page_limit_input) if page_limit_input.strip() else None
    input_file = input("Enter path to your Word (.docx) file: ")

    # 출력 파일 이름
    output_file = os.path.splitext(input_file)[0] + "_formatted.docx"

    if style.startswith('m'):
        mla_format(input_file, output_file, title, student_name, professor_name, course_name, page_limit)
    elif style.startswith('c'):
        chicago_format(input_file, output_file, title, student_name or None, professor_name or None, course_name or None, page_limit)
    elif style.startswith('a'):
        apa_format(input_file, output_file, title, student_name or None, professor_name or None, course_name or None, page_limit)
    elif style.startswith('h'):
        harvard_format(input_file, output_file, title, student_name or None, professor_name or None, course_name or None, page_limit)
    elif style.startswith('i'):
        ieee_format(input_file, output_file, title, student_name or None, professor_name or None, course_name or None, page_limit)
    else:
        print("Unknown style specified — defaulting to MLA.")
        mla_format(input_file, output_file, title, student_name, professor_name, course_name, page_limit)

if __name__ == "__main__":
    main()