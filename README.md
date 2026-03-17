# Word-Formatter

Word 문서를 MLA 형식으로 정리해주는 Python 앱입니다.

## 실행 방법

### Python까지 자동 설치해서 실행 (Linux/macOS)

Python 3가 없는 환경에서도 아래 스크립트가 Python 설치를 먼저 시도한 뒤 앱을 실행합니다.

```bash
bash run_with_python.sh
```

CLI 모드로 실행하려면:

```bash
bash run_with_python.sh --cli
```

### pip가 없는 환경에서 권장 실행

아래 명령으로 실행하면, `run_app.py`가 자동으로 다음을 시도합니다.

- `ensurepip`로 pip 설치
- 필수 패키지 설치 (`python-docx`, `requests`)
- Windows인 경우 `pywin32` 추가 설치
- 이후 앱 실행

```bash
python run_app.py
```

CLI 모드로 실행하려면:

```bash
python run_app.py --cli
```

### pip가 이미 있는 경우 수동 실행

1. 필요한 패키지 설치

```bash
pip install python-docx requests
```

2. GUI 앱 실행 (기본)

```bash
python wordFormatter.py
```

3. CLI 모드 실행

```bash
python wordFormatter.py --cli
```

## 의존성 목록

- [requirements.txt](requirements.txt)

## 앱에서 가능한 작업

- 스타일 선택 (MLA/Chicago/APA/Harvard/IEEE)
- 교수 호칭 선택 (doctor/professor) 후 이름만 입력
- 페이지 제한 설정 및 Works Cited 포함 여부 선택
- Works Cited 항목 여러 줄 입력
- 입력 .docx 선택 후 즉시 포맷팅 실행

출력 파일은 입력 파일명 뒤에 `_formatted.docx`가 붙어서 생성됩니다.