| 영어 컬럼명 | 한글 컬럼명 |
| --- | --- |
| dataTime | 측정시각 |
| pm10Value | PM10 |
| pm25Value | PM2.5 |
| pm10Grade | PM10등급 |
| pm25Grade | PM2.5등급 |
| so2Value | 아황산가스(SO2) |
| coValue | 일산화탄소(CO) |
| o3Value | 오존(O3) |
| no2Value | 이산화질소(NO2) |
| khaiValue | 통합대기환경지수 |
| khaiGrade | 통합등급 |

| 영어 컬럼명 | 한글 컬럼명 |
| --- | --- |
| stationName | 측정소명 |
| sidoName | 시도명 |
| pm10Value | PM10 |
| pm25Value | PM2.5 |
| pm10Grade | PM10등급 |
| pm25Grade | PM2.5등급 |
| so2Value | 아황산가스(SO2) |
| coValue | 일산화탄소(CO) |
| o3Value | 오존(O3) |
| no2Value | 이산화질소(NO2) |
| dataTime | 측정시간 |
| khaiValue | 통합대기환경지수 |
| khaiGrade | 통합등급 |

## 프로젝트 개요 :

공공데이터포털의 미세먼지 API를 활용, 지역별 or 측정소별 실시간 미세먼지 데이터 조회 및 리포트로 자동 변환 후 저장 또는 이메일로 발송을 할 수 있다.

## 구현 기능 :

지역별 / 측정소별 미세먼지 데이터 조회
시/도 및 측정소 선택 후 실시간/시간별

미세먼지 API 데이터를 수집, 가공, 리포트 생성, 이메일 발송 기능을 만들었다.
사용자의 인터페이스 제공으로 지역/측정소 선택, 리포트 생성 및 다운로드, 이메일 발송 등을 하였다.

## 사용 라이브러리 :

- Streamlit		: UI
- requests		: 외부 API 호출
- pandans		: 데이터프레임 처리 및 엑셀 저장
- openyxl		: 엑셀 파일 생성
- reportlab		: PDF 리포트 생성
- smtplib,email	: 이메일 발송
- plotly		: 데이터 시각화 ( 그래프)
- dotenv		: 환경변수 로드

## 사용 준비물 :

지역멸 미세먼지 농도 현황 API, 측정소 정보 API, Gmail 이메일 발송용 앱 비밀번호 .env에 양식에 맞게 입력
ex)

API_KEY=”여기에_발급받은_API_키_입력”

GMAIL_EMAIL=”본인_구글_이메일”

GMAIL_APP_PASSWORD=”앱_비밀번호”

## 주요 코드 설명 :

```python
from dotenv import load_dotenv
import os

load_dotenv()
SENDER_EMAIL = os.getenv("GMAIL_EMAIL")
SENDER_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
API_KEY = os.getenv("API_KEY")
```

API 키와 이메일 계정 정보 등 민감한 정보를 코드에 직접 쓰지 않고,
.env 파일에서 안전하게 불러옴

```python
def register_korean_font():
    for font_path in FONT_PATHS:
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
            return 'KoreanFont'
    return None
```

PDF 리포트에서 한글이 깨지지 않도록,
시스템 또는 프로젝트 내에 있는 한글 폰트를 등록

```python
class AirQualityReporter:
    def __init__(self, api_key):
        self.api_key = api_key
        # ... (API URL 등 초기화)

    def get_sido_list(self):
        # 시/도 목록 반환

    def fetch_station_list(self, sido_name):
        # 선택한 시/도의 측정소 목록을 API로부터 받아옴

    def fetch_station_dust_data(self, station_name):
        # 특정 측정소의 미세먼지 데이터를 API로부터 받아옴

    def fetch_sido_dust_data(self, sido_name):
        # 특정 시/도의 전체 미세먼지 데이터를 API로부터 받아옴

    def create_pdf_report(self, df, sido_name):
        # 데이터프레임을 PDF 리포트로 변환

    def create_excel_report(self, df, sido_name):
        # 데이터프레임을 Excel 파일로 변환

    def create_station_pdf_report(self, df, station_name):
        # 측정소별 PDF 리포트 생성

    def create_station_excel_report(self, df):
        # 측정소별 Excel 리포트 생성

    def send_email_report(self, to_email, subject, body, attachments, sender_email, sender_password):
        # 이메일로 리포트 파일 전송
```

미세먼지 데이터 수집, 리포트 생성, 이메일 발송 등
핵심 기능을 모두 담당하는 클래스

```python
def main():
    st.set_page_config(page_title="미세먼지 통합 리포트", layout="wide")
    st.title("🌫️ 미세먼지 통합 리포트 시스템")
    reporter = AirQualityReporter(API_KEY)

    # 사이드바: 모드, 지역, 측정소 선택, 이메일 입력 등
    # 메인: 데이터 조회, 표/그래프 출력, 리포트 다운로드, 이메일 발송
    # ... (상세 UI 코드)
```

Streamlit을 이용해 웹 대시보드를 구성, 사용자는 지역/측정소를 선택하고,
데이터 조회, 리포트 다운로드, 이메일 발송을 할 수 있다.

```python
def create_pdf_report(self, df, sido_name):
    # 한글 폰트 적용, 표와 제목, 생성일시 등 포함하여 PDF 생성

def create_excel_report(self, df, sido_name):
    # 컬럼명 한글 변환, 데이터프레임을 Excel로 저장
```

조회된 미세먼지 데이터를 PDF 또는 Excel 파일로 변환하여
다운로드 및 이메일 첨부가 가능하도록 한다.

```python
def send_email_report(self, to_email, subject, body, attachments, sender_email, sender_password):
    # 이메일 제목, 본문, 첨부파일 설정 후 Gmail SMTP로 발송
```

생성된 리포트 파일을 이메일로 자동 발송하는 기능

```python
fig = px.bar(df.head(10), x='stationName', y='pm10Value', title="...")
st.plotly_chart(fig)
```

Plotly를 이용해 미세먼지 농도를 막대그래프, 선그래프 등으로 시각화하여
사용자가 한눈에 데이터를 파악할 수 있도록 한다.
