from pathlib import Path
from docx import Document

source = Path(r"C:\Unity Project\XR_Escape\output\documents\표류_우주선_탈출_XR콘텐츠_기획서_편집용.docx")
output = Path(r"C:\Unity Project\XR_Escape\output\documents\표류_우주선_탈출_XR콘텐츠_기획서_수정본.docx")
doc = Document(source)

replacements = {
    "기말 프로젝트에서는 현재 데모의 구조를 XR Interaction Toolkit 기반의 손 상호작용으로 확장한다. 에너지 코어 삽입, 양손 밸브 조작, 공간 음향 추적과 같은 퍼즐을 추가하고, 각 퍼즐의 결과가 우주선 전력 복구와 탈출 포드 개방으로 이어지도록 구성할 계획이다.": "기말 프로젝트에서는 현재 데모의 구조를 XR Interaction Toolkit 기반의 손 상호작용으로 확장한다. 에너지 코어 삽입, 양손 밸브 조작, 공간 음향 추적과 같은 퍼즐을 추가하고, 각 퍼즐의 결과가 우주선 전력 복구와 탈출 포드 개방으로 이어지도록 구성할 계획이다. 이후에는 우주선뿐 아니라 연구소, 폐광, 고대 유적 등 서로 다른 콘셉트의 방탈출 맵을 여러 개 제작하여 맵마다 공간 연출과 상호작용 방식, 퍼즐 규칙이 달라지는 시리즈형 콘텐츠로 확장한다.",
    "기말 프로젝트의 목표는 현재 데모의 조건형 퍼즐 구조를 실제 XR 조작으로 전환하고, 우주선의 전력 복구 과정이 여러 구역을 연결하도록 확장하는 것이다. 플레이어가 손으로 잡고 돌리고 삽입하는 행동을 수행하며, 조명과 사운드 변화로 진행 상태를 이해할 수 있도록 한다.": "기말 프로젝트의 목표는 현재 데모의 조건형 퍼즐 구조를 실제 XR 조작으로 전환하고, 우주선의 전력 복구 과정이 여러 구역을 연결하도록 확장하는 것이다. 플레이어가 손으로 잡고 돌리고 삽입하는 행동을 수행하며, 조명과 사운드 변화로 진행 상태를 이해할 수 있도록 한다. 장기적으로는 우주선, 연구소, 폐광, 고대 유적처럼 콘셉트가 다른 방탈출 맵을 복수 제작하고, 공통 상호작용 시스템을 재사용하면서 각 맵에 고유한 환경 장치와 퍼즐을 구성할 계획이다.",
    "기말 단계에서는 중간 데모와 XR 확장 기능을 별도 씬 또는 브랜치로 관리한다.": "기말 단계에서는 중간 데모와 XR 확장 기능을 별도 씬 또는 브랜치로 관리하고, 이후 추가되는 콘셉트별 방탈출 맵은 독립 씬으로 제작한다.",
}

replacement_hits = 0
for paragraph in doc.paragraphs:
    if paragraph.text in replacements:
        paragraph.text = replacements[paragraph.text]
        replacement_hits += 1

phrase = "XR콘텐츠응용 중간과제 기획서"
footer_hits = 0
for section in doc.sections:
    for paragraph in section.footer.paragraphs:
        for run in paragraph.runs:
            if phrase in run.text:
                run.text = run.text.replace(phrase, "")
                footer_hits += 1

doc.save(output)
print(f"saved={output}")
print(f"replacement_hits={replacement_hits} footer_hits={footer_hits}")
