<div align="center">

# Chrome 다중 창 관리자

[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6.svg?style=flat&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Chrome](https://img.shields.io/badge/Chrome-Latest-4285F4.svg?style=flat&logo=google-chrome&logoColor=white)](https://www.google.com/chrome/)
[![License](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](LICENSE)



  <strong>제작자：Devilflasher</strong>：<span title="No Biggie Community Founder"></span>
  [![X](https://img.shields.io/badge/X-1DA1F2.svg?style=flat&logo=x&logoColor=white)](https://x.com/DevilflasherX)
[![WeChat](https://img.shields.io/badge/WeChat-7BB32A.svg?style=flat&logo=wechat&logoColor=white)](https://x.com/DevilflasherX/status/1781563666485448736 "Devilflasherx")
 [![Telegram](https://img.shields.io/badge/Telegram-0A74DA.svg?style=flat&logo=telegram&logoColor=white)](https://t.me/devilflasher0) (WeChat 그룹 참여 환영)
 

</div>

> [!IMPORTANT]
> ## ⚠️ 면책 조항
> 
> 1. **본 소프트웨어는 오픈소스 프로젝트로, 학습 및 교류 목적으로만 사용되어야 하며 비공개 상업적 용도로 사용할 수 없습니다**
> 2. **사용자는 현지 법률과 규정을 준수해야 하며, 불법적인 용도로의 사용을 금지합니다**
> 3. **개발자는 본 소프트웨어 사용으로 인한 직접/간접적 손실에 대해 어떠한 책임도 지지 않습니다**
> 4. **본 소프트웨어를 사용하는 것은 이 면책 조항을 읽고 동의했음을 의미합니다**

## 도구 소개
Chrome 다중 창 관리자는 `NoBiggie 커뮤니티`를 위해 특별히 제작된 Chrome 브라우저 다중 창 관리 도구입니다. 사용자가 여러 Chrome 창을 쉽게 관리하고, 창을 일괄적으로 열고 정렬하며 동기화 작업을 수행할 수 있어 작업 효율을 크게 향상시킵니다.

## 주요 기능

- `일괄 관리 기능`: 단일 또는 다중 Chrome 인스턴스를 원클릭으로 열기/닫기
- `스마트 레이아웃 시스템`: 자동 그리드 정렬 및 사용자 정의 좌표 레이아웃 지원
- `다중 창 동기화 제어`: 선택된 모든 창에 실시간 마우스/키보드 작업 동기화
- `일괄 웹페이지 열기`: 동일한 웹페이지 일괄 열기 지원
- `바로가기 아이콘 교체`: 여러 바로가기 아이콘 원클릭 교체 지원 (번호가 있는 아이콘은 icon 폴더에 준비되어 있음)
- `플러그인 창 동기화`: 팝업된 플러그인 창 내의 키보드 및 마우스 동기화 지원

## 시스템 요구사항

- Windows 10/11 (64-bit)
- Python 3.9+
- Chrome 브라우저 최신 버전

## 실행 가이드
### 방법 1: 독립 실행 파일로 패키징 (권장)

프로그램을 직접 패키징하려면 다음 단계를 따르세요:

1. **Python 및 의존성 설치**
   ```bash
   # Python 3.9 이상 버전 설치
   # https://www.python.org/downloads/ 에서 다운로드
   ```

2. **파일 준비**
   - 다음 파일들이 디렉토리에 있는지 확인:
     - chrome_manager.py (메인 프로그램)
     - build.py (패키징 스크립트)
     - app.manifest (관리자 권한 설정)
     - app.ico (프로그램 아이콘)

3. **패키징 스크립트 실행**
   ```bash
   # 프로그램 디렉토리에서 실행:
   python build.py
   ```

4. **생성된 파일 확인**
   - 패키징이 완료되면 `dist` 디렉토리에서 `chrome_manager.exe` 찾기
   - `chrome_manager.exe`를 더블클릭하여 프로그램 실행

### 방법 2: 소스 코드에서 실행

1. **Python 설치**
   ```bash
   # Python 3.9 이상 버전 다운로드 및 설치
   # https://www.python.org/downloads/ 에서 다운로드
   ```

2. **의존성 패키지 설치**
   ```bash
   # 명령 프롬프트(CMD)에서 실행:
   pip install tkinter pywin32 keyboard mouse sv-ttk typing-extensions
   ```

3. **프로그램 실행**
   ```bash
   # 프로그램 디렉토리에서 실행:
   python chrome_manager.py
   ```

## 사용 설명서

### 사전 준비


- Chrome 다중 실행 바로가기를 저장하는 폴더에서 바로가기 파일명을 `1.link`, `2.link`, `3.link`... 형식으로 지정해야 합니다.
- 같은 폴더에 `Data` 폴더를 만들고, `Data` 폴더 안에 각 브라우저의 독립적인 데이터 파일을 저장하는 폴더를 `1`, `2`, `3`... 형식으로 지정해야 합니다.

```디렉토리 구조 예시:
                                 다중 Chrome 디렉토리

                                ├── 1.link
                                ├── 2.link
                                ├── 3.link
                                └── Data
                                    ├── 1
                                    ├── 2
                                    └── 3
```
- 브라우저 바로가기의 대상 매개변수는 다음과 같습니다: (브라우저 설치 경로에 맞게 수정하세요)
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="D:\chrom duo\Data\번호"
```

### 기본 조작

1. **창 열기**
   - 소프트웨어 하단의 "창 열기" 탭에서 브라우저 바로가기가 있는 디렉토리 입력
   - "창 번호"에 열고자 하는 브라우저 번호 입력
   - "창 열기" 버튼을 클릭하여 해당 번호의 Chrome 창 열기

2. **창 가져오기**
   - "창 가져오기" 버튼을 클릭하여 현재 열려있는 Chrome 창 가져오기
   - 목록에서 작업할 창 선택

3. **창 정렬**
   - "자동 정렬"로 창 빠르게 정리
   - 또는 "사용자 정의 정렬"로 상세 정렬 매개변수 설정

4. **동기화 시작**
   - 주 제어 창 선택 ("주 제어" 열 클릭)
   - 동기화할 종속 창 선택
   - "동기화 시작" 클릭 또는 설정된 단축키 사용



## 주의사항

- 동기화 기능은 관리자 권한이 필요합니다
- 이론적으로 백신 프로그램의 오탐이나 방해를 받지 않지만, 오류 발생 시 백신 프로그램이 관련 기능을 차단하지 않았는지 확인하세요
- 일괄 작업 시 시스템 리소스 사용량에 주의하세요

## 자주 묻는 질문

1. **동기화를 시작할 수 없음**
   - 관리자 권한으로 실행했는지 확인
   - 주 제어 창이 선택되었는지 확인

2. **창이 제대로 가져와지지 않음**
   - "창 가져오기" 버튼을 다시 클릭해보세요
   - 
3. **스크롤바 동기화 폭이 다름**
   - 현재로서는 PageUp과 PageDown, 키보드 방향키로 동기화 폭을 조정하는 것이 해결책입니다
   
  

## 업데이트 로그

### v1.0
- 최초 출시
- 기본 창 관리 및 동기화 기능 구현


## 라이선스

본 프로젝트는 GPL-3.0 라이선스를 채택하며, 모든 권리를 보유합니다. 이 코드를 사용할 때는 출처를 명확히 표시해야 하며, 비공개 상업적 사용은 금지됩니다.

🔄 지속적으로 업데이트 중

