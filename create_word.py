from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# 기본 스타일 설정
style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

# 여백 설정
sections = doc.sections
for section in sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

def add_heading_text(text, bold=True, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    return p

def add_text(text, bold=False, size=11, alignment=None):
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment
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
# 제목
# ─────────────────────────────────────
add_heading_text('배재고등학교 1학년 영어 내신 대비', size=16)
add_heading_text('일치 / 불일치 유형 문항', size=14)
add_empty_line()

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

# ─────────────────────────────────────
# 1번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '1. ', bold=True, size=12)
add_mixed_run(p, '다음 글의 내용과 일치하지 ', bold=True, size=12)
add_mixed_run(p, '않는', bold=True, underline=True, size=12)
add_mixed_run(p, ' 것은? [3.1점]', bold=True, size=12)

add_empty_line()

# 지문 박스 (들여쓰기로 표현)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
p.paragraph_format.right_indent = Cm(0.5)
run = p.add_run(passage)
run.font.size = Pt(11)
run.font.name = 'Times New Roman'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

choices_1 = [
    '① Communication plays a crucial role in keeping long-term relationships going.',
    '② Relationships can become stagnant when there is insufficient follow-up.',
    '③ Staying in touch requires a deliberate and mindful approach.',
    '④ Reaching out frequently with various requests helps strengthen the bond between people.',
    '⑤ Sharing a useful article with someone you haven\'t contacted recently can help sustain the relationship.'
]

for c in choices_1:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    run = p.add_run(c)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

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

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
p.paragraph_format.right_indent = Cm(0.5)
run = p.add_run(passage)
run.font.size = Pt(11)
run.font.name = 'Times New Roman'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

choices_2 = [
    '① 장기적인 관계를 유지하는 데 있어 가장 중요한 요소 중 하나는 소통이다.',
    '② 누군가와 연결된 후에도 지속적으로 후속 조치를 취하면 관계가 자연스럽게 깊어진다.',
    '③ 연락을 유지한다는 것은 끊임없이 근황이나 요청 사항을 전달하는 것이 아니라 우호적이고 일관된 소통을 이어가는 것이다.',
    '④ 안부를 확인하거나 가치 있는 것을 공유하는 간단한 메시지만으로도 관계를 강화하는 데 충분하다.',
    '⑤ 오랫동안 연락하지 않았던 사람에게 유용한 자료를 공유하는 것은 상대방이 자신에게 투자하고 있다는 것을 보여준다.'
]

for c in choices_2:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    run = p.add_run(c)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

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

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
p.paragraph_format.right_indent = Cm(0.5)
run = p.add_run(passage)
run.font.size = Pt(11)
run.font.name = 'Times New Roman'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

# 보기 박스
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_mixed_run(p_title, '< 보 기 >', bold=True, size=11)

bogi = [
    'ㄱ. Following up after making a connection is essential to prevent the relationship from losing momentum.',
    'ㄴ. The passage suggests that the quality of communication matters more than its frequency when it comes to sustaining relationships.',
    'ㄷ. Sharing valuable information only with those you regularly keep in touch with is recommended for relationship building.',
    'ㄹ. A conscious effort to maintain communication implies that relationship maintenance happens naturally without much thought.',
    'ㅁ. The act of sharing a relevant resource demonstrates genuine care for the other person.'
]

for b in bogi:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(b)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
items = ['① ㄱ, ㄴ          ', '② ㄱ, ㅁ          ', '③ ㄴ, ㄷ']
for item in items:
    add_mixed_run(p, item, size=11)
p2 = doc.add_paragraph()
p2.paragraph_format.left_indent = Cm(0.3)
items2 = ['④ ㄱ, ㄴ, ㅁ          ', '⑤ ㄴ, ㄷ, ㅁ']
for item in items2:
    add_mixed_run(p2, item, size=11)

add_empty_line()
add_empty_line()

# ─────────────────────────────────────
# 4번 문항
# ─────────────────────────────────────
p = doc.add_paragraph()
add_mixed_run(p, '4. ', bold=True, size=12)
add_mixed_run(p, '다음 글에서 추론할 수 있는 것으로 가장 적절한 것은? [3.3점]', bold=True, size=12)

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
p.paragraph_format.right_indent = Cm(0.5)
run = p.add_run(passage)
run.font.size = Pt(11)
run.font.name = 'Times New Roman'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

choices_4 = [
    '① Relationships that lack regular communication are bound to end permanently.',
    '② The best way to reconnect with someone is to ask them for help with a personal matter.',
    '③ People who frequently share news and requests with their connections are the most effective communicators.',
    '④ Maintaining a relationship requires ongoing attention, even when there is no immediate reason to reach out.',
    '⑤ Sharing an article with a connection is effective only when the content directly benefits both parties involved.'
]

for c in choices_4:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    run = p.add_run(c)
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

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

add_empty_line()

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
p.paragraph_format.right_indent = Cm(0.5)
run = p.add_run(passage)
run.font.size = Pt(11)
run.font.name = 'Times New Roman'
run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

add_empty_line()

# 대화문
conversations = [
    ('A:', 'I think the passage is saying that we need to put effort into staying connected with people.'),
    ('B:', None),  # special: ① 포함
    ('A:', 'So does that mean we should message people all the time?'),
    ('B:', None),  # special: ② 포함
    ('A:', 'Oh, I see. So even a short check-in message would help?'),
    ('B:', None),  # special: ③ 포함
    ('A:', 'What about sharing articles? The passage mentions that too, right?'),
    ('B:', None),  # special: ④ 포함
    ('A:', 'That makes sense. It really shows you care.'),
    ('B:', None),  # special: ⑤ 포함
]

# ① 대화
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'A: ', bold=True, size=11)
add_mixed_run(p, 'I think the passage is saying that we need to put effort into staying connected with people.', size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'B: ', bold=True, size=11)
add_mixed_run(p, 'Right. ', size=11)
add_mixed_run(p, '①', bold=True, size=11)
add_mixed_run(p, ' ', size=11)
add_mixed_run(p, 'It says relationships can easily get stuck if we don\'t follow up after making a connection.', underline=True, size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'A: ', bold=True, size=11)
add_mixed_run(p, 'So does that mean we should message people all the time?', size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'B: ', bold=True, size=11)
add_mixed_run(p, 'No, actually the opposite. ', size=11)
add_mixed_run(p, '②', bold=True, size=11)
add_mixed_run(p, ' ', size=11)
add_mixed_run(p, 'The key is to keep communication friendly and steady, not to bombard people with constant messages.', underline=True, size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'A: ', bold=True, size=11)
add_mixed_run(p, 'Oh, I see. So even a short check-in message would help?', size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'B: ', bold=True, size=11)
add_mixed_run(p, 'Exactly. ', size=11)
add_mixed_run(p, '③', bold=True, size=11)
add_mixed_run(p, ' ', size=11)
add_mixed_run(p, 'The passage says that even a brief message can significantly contribute to making the relationship stronger.', underline=True, size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'A: ', bold=True, size=11)
add_mixed_run(p, 'What about sharing articles? The passage mentions that too, right?', size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'B: ', bold=True, size=11)
add_mixed_run(p, 'Yes. ', size=11)
add_mixed_run(p, '④', bold=True, size=11)
add_mixed_run(p, ' ', size=11)
add_mixed_run(p, 'It suggests sharing something useful, but only after you\'ve confirmed that the other person wants to receive it.', underline=True, size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'A: ', bold=True, size=11)
add_mixed_run(p, 'That makes sense. It really shows you care.', size=11)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.5)
add_mixed_run(p, 'B: ', bold=True, size=11)
add_mixed_run(p, '⑤', bold=True, size=11)
add_mixed_run(p, ' ', size=11)
add_mixed_run(p, 'That\'s what the passage says — sharing something relevant shows you\'re thinking about them and committed to the relationship.', underline=True, size=11)

# ─────────────────────────────────────
# 페이지 나누기 → 정답 및 해설
# ─────────────────────────────────────
doc.add_page_break()

add_heading_text('정답 및 해설', size=16)
add_empty_line()

# 1번 해설
p = doc.add_paragraph()
add_mixed_run(p, '1번  정답: ④', bold=True, size=12)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
add_mixed_run(p, '지문에서 부정 구문(doesn\'t mean ~)으로 명시적으로 배제한 내용을 선지에서 긍정적으로 서술할 때, 이를 정확히 판별하는 능력을 평가하는 문항.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설] ', bold=True, size=10)
add_mixed_run(p, '지문은 "This ', size=10)
add_mixed_run(p, 'doesn\'t mean', bold=True, size=10)
add_mixed_run(p, ' constantly reaching out with requests or updates but rather maintaining a friendly and consistent line of communication"이라고 명시했다. '
    '즉, 자주 요청을 하며 연락하는 것이 관계를 강화한다는 것이 아니라, 우호적이고 일관된 소통을 유지하는 것이 핵심이다. '
    '④번은 지문이 명시적으로 부정한 내용(requests를 통한 빈번한 연락)을 일치하는 것처럼 서술하였다.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[오답 설계] ', bold=True, size=10)
add_mixed_run(p, '부정어 삭제 함정 — doesn\'t mean을 제거하고 그 내용을 사실처럼 제시.', size=10)

add_empty_line()

# 2번 해설
p = doc.add_paragraph()
add_mixed_run(p, '2번  정답: ②, ⑤', bold=True, size=12)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
add_mixed_run(p, '지문의 인과관계를 정확히 파악하는 능력(②)과 행위 주체를 정밀하게 구별하는 능력(⑤)을 동시에 평가하는 고난도 문항.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설 - ②] ', bold=True, size=10)
add_mixed_run(p, '지문은 후속 조치가 부족하면 관계가 정체(get stuck)된다고 했을 뿐, "지속적으로 후속 조치를 취하면 자연스럽게 깊어진다"고 하지 않았다. '
    '인과 방향 뒤집기 + 무근거 추가.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설 - ⑤] ', bold=True, size=10)
add_mixed_run(p, '지문은 "you\'re thinking of them and are invested"라고 하여 투자의 주체가 \'나(you)\'이다. '
    '선지는 "상대방이 자신에게 투자하고 있다"로 주체와 대상을 바꿔치기했다.', size=10)

add_empty_line()

# 3번 해설
p = doc.add_paragraph()
add_mixed_run(p, '3번  정답: ④ (ㄱ, ㄴ, ㅁ)', bold=True, size=12)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
add_mixed_run(p, '핵심 메시지(소통의 질 > 빈도)를 파악하고, 한정사 삽입(ㄷ)과 반의어 교체(ㄹ) 함정을 판별하는 능력 평가.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설 - ㄷ 불일치] ', bold=True, size=10)
add_mixed_run(p, '지문은 "even if you haven\'t spoken in a while"이라 하여 오래 연락 안 한 사람에게도 공유하라고 했다. '
    '선지의 "only with those you regularly keep in touch with"는 한정사 only를 삽입하여 의미를 반대로 전환.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설 - ㄹ 불일치] ', bold=True, size=10)
add_mixed_run(p, '지문의 "conscious effort(의식적 노력)"와 선지의 "naturally without much thought(자연스럽게 별 생각 없이)"는 정반대. 반의어 교체 함정.', size=10)

add_empty_line()

# 4번 해설
p = doc.add_paragraph()
add_mixed_run(p, '4번  정답: ④', bold=True, size=12)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
add_mixed_run(p, '지문 전체의 핵심 메시지를 종합하여 일반화된 원칙을 도출하는 추론 능력 평가.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설] ', bold=True, size=10)
add_mixed_run(p, '④ "conscious effort" + "even if you haven\'t spoken in a while"에서, 즉각적 이유가 없어도 지속적 관심이 필요하다는 추론이 가능하다. '
    '① stuck(정체)≠end permanently(영구 종료)로 정도 과장. '
    '②③ 지문이 부정한 내용(requests)을 긍정적으로 전환. '
    '⑤ "only when ~ both parties"라는 이중 조건은 지문에 없음.', size=10)

add_empty_line()

# 5번 해설
p = doc.add_paragraph()
add_mixed_run(p, '5번  정답: ④', bold=True, size=12)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[출제 의도] ', bold=True, size=10)
add_mixed_run(p, '대화문이라는 비전형적 형식에서 지문의 세부 조건을 정밀하게 판별하는 능력 평가. '
    '대화의 자연스러운 흐름 속에 무근거 조건을 삽입한 함정.', size=10)

p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.3)
add_mixed_run(p, '[해설] ', bold=True, size=10)
add_mixed_run(p, '지문은 "share it with them, even if you haven\'t spoken in a while"이라고 하여 사전 확인 없이도 공유를 권장한다. '
    '④의 "but only after you\'ve confirmed that the other person wants to receive it"은 지문에 없는 조건을 추가한 것이다. '
    '대화 맥락상 "배려하는 행동"으로 읽혀 자연스러워 보이지만, 지문은 이런 조건을 전혀 요구하지 않는다.', size=10)

# ─────────────────────────────────────
# 저장
# ─────────────────────────────────────
output_path = '/home/user/Paichai-1st/배재고_일치불일치_문항.docx'
doc.save(output_path)
print(f'파일 저장 완료: {output_path}')
