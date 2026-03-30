from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from lxml import etree
import copy

doc = Document()

# 기본 스타일 설정
style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

# 여백 설정
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# ─────────────────────────────────────
# 미주(Endnote) 기능 구현
# ─────────────────────────────────────
WPML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

endnote_counter = [0]  # mutable counter

def setup_endnotes_part(document):
    """문서에 endnotes.xml 파트를 생성"""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    # endnotes.xml 내용 생성 (separator와 continuationSeparator 포함)
    endnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:endnotes xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:mv="urn:schemas-microsoft-com:mac:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14">'
        '<w:endnote w:type="separator" w:id="-1">'
        '<w:p><w:r><w:separator/></w:r></w:p>'
        '</w:endnote>'
        '<w:endnote w:type="continuationSeparator" w:id="0">'
        '<w:p><w:r><w:continuationSeparator/></w:r></w:p>'
        '</w:endnote>'
        '</w:endnotes>'
    )

    # Part 생성
    part_name = PackURI('/word/endnotes.xml')
    content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'
    rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes'

    endnotes_part = Part(
        part_name, content_type, endnotes_xml.encode('utf-8'),
        document.part.package
    )

    document.part.relate_to(endnotes_part, rel_type)

    return endnotes_part

def add_endnote_to_part(endnotes_part, endnote_id, text):
    """endnotes.xml에 미주 내용 추가"""
    endnotes_elem = etree.fromstring(endnotes_part.blob)

    # 새 endnote 요소 생성
    endnote = etree.SubElement(endnotes_elem, qn('w:endnote'))
    endnote.set(qn('w:id'), str(endnote_id))

    p = etree.SubElement(endnote, qn('w:p'))

    # 미주 번호 참조 run
    r1 = etree.SubElement(p, qn('w:r'))
    rPr1 = etree.SubElement(r1, qn('w:rPr'))
    rStyle1 = etree.SubElement(rPr1, qn('w:rStyle'))
    rStyle1.set(qn('w:val'), 'EndnoteReference')
    endnoteRef = etree.SubElement(r1, qn('w:endnoteRef'))

    # 미주 텍스트 run
    r2 = etree.SubElement(p, qn('w:r'))
    t2 = etree.SubElement(r2, qn('w:t'))
    t2.set(qn('xml:space'), 'preserve')
    t2.text = ' ' + text

    # rPr에 폰트 설정
    rPr2 = etree.SubElement(r2, qn('w:rPr'))
    rFonts2 = etree.SubElement(rPr2, qn('w:rFonts'))
    rFonts2.set(qn('w:eastAsia'), '맑은 고딕')
    sz2 = etree.SubElement(rPr2, qn('w:sz'))
    sz2.set(qn('w:val'), '18')  # 9pt

    endnotes_part._blob = etree.tostring(endnotes_elem, xml_declaration=True, encoding='UTF-8', standalone=True)

def add_endnote_reference(paragraph, endnote_id):
    """본문 paragraph에 미주 참조 마크 삽입"""
    run = OxmlElement('w:r')

    # 위첨자 스타일
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'EndnoteReference')
    vertAlign = OxmlElement('w:vertAlign')
    vertAlign.set(qn('w:val'), 'superscript')
    rPr.append(rStyle)
    rPr.append(vertAlign)

    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '18')
    rPr.append(sz)

    run.append(rPr)

    endnoteReference = OxmlElement('w:endnoteReference')
    endnoteReference.set(qn('w:id'), str(endnote_id))
    run.append(endnoteReference)

    paragraph._element.append(run)

# ─────────────────────────────────────
# 헬퍼 함수
# ─────────────────────────────────────
def add_heading_text(text, bold=True, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    return p

def add_mixed_run(paragraph, text, bold=False, underline=False, size=11):
    run = paragraph.add_run(text)
    run.bold = bold
    run.underline = underline
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    return run

def add_empty_line():
    p = doc.add_paragraph()
    p.space_after = Pt(0)
    p.space_before = Pt(0)
    return p

# ─────────────────────────────────────
# endnotes 파트 설정
# ─────────────────────────────────────
endnotes_part = setup_endnotes_part(doc)

# 정답 정보
answers = {
    1: '정답: ④',
    2: '정답: ②, ⑤',
    3: '정답: ④ (ㄱ, ㄴ, ㅁ)',
    4: '정답: ④',
    5: '정답: ④'
}

# ─────────────────────────────────────
# 지문 (공통)
# ─────────────────────────────────────
passage = (
    "One of the most important aspects of sustaining long-term relationships is communication. "
    "It's easy to connect with someone and then let the relationship get stuck due to a lack of follow-up. "
    "To keep the connection alive, make a conscious effort to stay in touch. "
    "This doesn't mean constantly reaching out with requests or updates but rather maintaining "
    "a friendly and consistent line of communication. A simple message to check in or share something "
    "of value can go a long way in reinforcing your relationship. For example, if you come across an "
    "article or resource that you think might interest a connection, share it with them, even if you "
    "haven't spoken in a while. This shows that you're thinking of them and are invested in maintaining "
    "the relationship."
)

def add_passage(doc):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    run = p.add_run(passage)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def add_choices(choices):
    for c in choices:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.3)
        run = p.add_run(c)
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

# ─────────────────────────────────────
# 제목
# ─────────────────────────────────────
add_heading_text('배재고등학교 1학년 영어 내신 대비', size=16)
add_heading_text('일치 / 불일치 유형 문항', size=14)
add_empty_line()

# ─────────────────────────────────────
# 1번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '1. ', bold=True, size=12)
add_mixed_run(p, '다음 글의 내용과 일치하지 ', bold=True, size=12)
add_mixed_run(p, '않는', bold=True, underline=True, size=12)
add_mixed_run(p, ' 것은? [3.1점]', bold=True, size=12)
# 미주 삽입
add_endnote_to_part(endnotes_part, 1, answers[1])
add_endnote_reference(p, 1)

add_empty_line()
add_passage(doc)
add_empty_line()

add_choices([
    '① Communication plays a crucial role in keeping long-term relationships going.',
    '② Relationships can become stagnant when there is insufficient follow-up.',
    '③ Staying in touch requires a deliberate and mindful approach.',
    '④ Reaching out frequently with various requests helps strengthen the bond between people.',
    '⑤ Sharing a useful article with someone you haven\'t contacted recently can help sustain the relationship.'
])

add_empty_line()
add_empty_line()

# ─────────────────────────────────────
# 2번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '2. ', bold=True, size=12)
add_mixed_run(p, '다음 글의 내용과 일치하지 ', bold=True, size=12)
add_mixed_run(p, '않는', bold=True, underline=True, size=12)
add_mixed_run(p, ' 것을 2개 고르면? [3.5점]', bold=True, size=12)
add_endnote_to_part(endnotes_part, 2, answers[2])
add_endnote_reference(p, 2)

add_empty_line()
add_passage(doc)
add_empty_line()

add_choices([
    '① 장기적인 관계를 유지하는 데 있어 가장 중요한 요소 중 하나는 소통이다.',
    '② 누군가와 연결된 후에도 지속적으로 후속 조치를 취하면 관계가 자연스럽게 깊어진다.',
    '③ 연락을 유지한다는 것은 끊임없이 근황이나 요청 사항을 전달하는 것이 아니라 우호적이고 일관된 소통을 이어가는 것이다.',
    '④ 안부를 확인하거나 가치 있는 것을 공유하는 간단한 메시지만으로도 관계를 강화하는 데 충분하다.',
    '⑤ 오랫동안 연락하지 않았던 사람에게 유용한 자료를 공유하는 것은 상대방이 자신에게 투자하고 있다는 것을 보여준다.'
])

add_empty_line()
add_empty_line()

# ─────────────────────────────────────
# 3번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '3. ', bold=True, size=12)
add_mixed_run(p, '다음 중 윗글의 내용과 일치하는 것을 ', bold=True, size=12)
add_mixed_run(p, '모두', bold=True, underline=True, size=12)
add_mixed_run(p, ' 고르면? [3.4점]', bold=True, size=12)
add_endnote_to_part(endnotes_part, 3, answers[3])
add_endnote_reference(p, 3)

add_empty_line()
add_passage(doc)
add_empty_line()

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_mixed_run(p_title, '< 보 기 >', bold=True, size=11)

for b in [
    'ㄱ. Following up after making a connection is essential to prevent the relationship from losing momentum.',
    'ㄴ. The passage suggests that the quality of communication matters more than its frequency when it comes to sustaining relationships.',
    'ㄷ. Sharing valuable information only with those you regularly keep in touch with is recommended for relationship building.',
    'ㄹ. A conscious effort to maintain communication implies that relationship maintenance happens naturally without much thought.',
    'ㅁ. The act of sharing a relevant resource demonstrates genuine care for the other person.'
]:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(b)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
for item in ['① ㄱ, ㄴ          ', '② ㄱ, ㅁ          ', '③ ㄴ, ㄷ']:
    add_mixed_run(p, item, size=11)
p2 = doc.add_paragraph()
p2.paragraph_format.left_indent = Cm(0.3)
for item in ['④ ㄱ, ㄴ, ㅁ          ', '⑤ ㄴ, ㄷ, ㅁ']:
    add_mixed_run(p2, item, size=11)

add_empty_line()
add_empty_line()

# ─────────────────────────────────────
# 4번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '4. ', bold=True, size=12)
add_mixed_run(p, '다음 글에서 추론할 수 있는 것으로 가장 적절한 것은? [3.3점]', bold=True, size=12)
add_endnote_to_part(endnotes_part, 4, answers[4])
add_endnote_reference(p, 4)

add_empty_line()
add_passage(doc)
add_empty_line()

add_choices([
    '① Relationships that lack regular communication are bound to end permanently.',
    '② The best way to reconnect with someone is to ask them for help with a personal matter.',
    '③ People who frequently share news and requests with their connections are the most effective communicators.',
    '④ Maintaining a relationship requires ongoing attention, even when there is no immediate reason to reach out.',
    '⑤ Sharing an article with a connection is effective only when the content directly benefits both parties involved.'
])

add_empty_line()
add_empty_line()

# ─────────────────────────────────────
# 5번 문항 (대화문)
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '5. ', bold=True, size=12)
add_mixed_run(p, '윗글을 읽은 두 학생이 나눈 대화로 옳지 ', bold=True, size=12)
add_mixed_run(p, '않은', bold=True, underline=True, size=12)
add_mixed_run(p, ' 것은? [3.5점]', bold=True, size=12)
add_endnote_to_part(endnotes_part, 5, answers[5])
add_endnote_reference(p, 5)

add_empty_line()
add_passage(doc)
add_empty_line()

# 대화문
dialog_data = [
    ('A:', 'I think the passage is saying that we need to put effort into staying connected with people.', None, None),
    ('B:', 'Right. ', '①', 'It says relationships can easily get stuck if we don\'t follow up after making a connection.'),
    ('A:', 'So does that mean we should message people all the time?', None, None),
    ('B:', 'No, actually the opposite. ', '②', 'The key is to keep communication friendly and steady, not to bombard people with constant messages.'),
    ('A:', 'Oh, I see. So even a short check-in message would help?', None, None),
    ('B:', 'Exactly. ', '③', 'The passage says that even a brief message can significantly contribute to making the relationship stronger.'),
    ('A:', 'What about sharing articles? The passage mentions that too, right?', None, None),
    ('B:', 'Yes. ', '④', 'It suggests sharing something useful, but only after you\'ve confirmed that the other person wants to receive it.'),
    ('A:', 'That makes sense. It really shows you care.', None, None),
    ('B:', '', '⑤', 'That\'s what the passage says \u2014 sharing something relevant shows you\'re thinking about them and committed to the relationship.'),
]

for speaker, text, num, underlined_text in dialog_data:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    add_mixed_run(p, speaker + ' ', bold=True, size=11)
    add_mixed_run(p, text, size=11)
    if num:
        add_mixed_run(p, num, bold=True, size=11)
        add_mixed_run(p, ' ', size=11)
        add_mixed_run(p, underlined_text, underline=True, size=11)

# ─────────────────────────────────────
# 페이지 나누기 → 정답 및 해설
# ─────────────────────────────────────
doc.add_page_break()

add_heading_text('정답 및 해설', size=16)
add_empty_line()

# 해설 데이터
explanations = [
    ('1번  정답: ④',
     '지문에서 부정 구문(doesn\'t mean ~)으로 명시적으로 배제한 내용을 선지에서 긍정적으로 서술할 때, 이를 정확히 판별하는 능력을 평가하는 문항.',
     '지문은 "This doesn\'t mean constantly reaching out with requests or updates but rather maintaining a friendly and consistent line of communication"이라고 명시했다. ④번은 지문이 명시적으로 부정한 내용(requests를 통한 빈번한 연락)을 일치하는 것처럼 서술하였다.',
     '부정어 삭제 함정 — doesn\'t mean을 제거하고 그 내용을 사실처럼 제시.'),
    ('2번  정답: ②, ⑤',
     '지문의 인과관계를 정확히 파악하는 능력(②)과 행위 주체를 정밀하게 구별하는 능력(⑤)을 동시에 평가하는 고난도 문항.',
     '② 지문은 후속 조치가 부족하면 관계가 정체(get stuck)된다고 했을 뿐, "자연스럽게 깊어진다"고 하지 않았다. 인과 방향 뒤집기 + 무근거 추가.\n'
     '⑤ 지문은 "you\'re thinking of them and are invested"라고 하여 투자의 주체가 \'나(you)\'이다. 선지는 "상대방이 자신에게 투자하고 있다"로 주체와 대상을 바꿔치기했다.',
     '② 인과 방향 뒤집기  ⑤ 주체 바꿔치기'),
    ('3번  정답: ④ (ㄱ, ㄴ, ㅁ)',
     '핵심 메시지(소통의 질 > 빈도)를 파악하고, 한정사 삽입(ㄷ)과 반의어 교체(ㄹ) 함정을 판별하는 능력 평가.',
     'ㄷ 불일치: "only with those you regularly keep in touch with"는 한정사 only를 삽입하여 의미를 반대로 전환.\n'
     'ㄹ 불일치: "conscious effort(의식적 노력)"와 "naturally without much thought"는 정반대. 반의어 교체.',
     'ㄷ 한정사 only 삽입  ㄹ 반의어 교체'),
    ('4번  정답: ④',
     '지문 전체의 핵심 메시지를 종합하여 일반화된 원칙을 도출하는 추론 능력 평가.',
     '④ "conscious effort" + "even if you haven\'t spoken in a while"에서, 즉각적 이유가 없어도 지속적 관심이 필요하다는 추론이 가능. ① stuck≠end permanently 정도 과장. ②③ 지문이 부정한 내용을 긍정적으로 전환. ⑤ "only when ~ both parties" 무근거 조건 추가.',
     '① 정도 과장  ②③ 부정 내용 긍정화  ⑤ 무근거 조건 추가'),
    ('5번  정답: ④',
     '대화문 형식에서 지문의 세부 조건을 정밀하게 판별하는 능력 평가.',
     '지문은 사전 확인 없이도 공유를 권장한다. ④의 "but only after you\'ve confirmed that the other person wants to receive it"은 지문에 없는 조건을 추가한 것이다.',
     '무근거 조건 추가 + 대화문 경계심 저하 효과 활용')
]

for title, intent, explanation, trap in explanations:
    p = doc.add_paragraph()
    add_mixed_run(p, title, bold=True, size=12)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
    add_mixed_run(p, intent, size=10)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    add_mixed_run(p, '[해설] ', bold=True, size=10)
    add_mixed_run(p, explanation, size=10)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    add_mixed_run(p, '[오답 설계] ', bold=True, size=10)
    add_mixed_run(p, trap, size=10)

    add_empty_line()

# ─────────────────────────────────────
# 저장
# ─────────────────────────────────────
output_path = '/home/user/Paichai-1st/배재고_일치불일치_문항.docx'
doc.save(output_path)
print(f'파일 저장 완료: {output_path}')
print('각 문항 발문 끝에 미주(endnote)로 정답 번호가 삽입되었습니다.')
