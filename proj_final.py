import streamlit as st
import requests
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
from email.message import EmailMessage
import io
import plotly.express as px
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import os
from dotenv import load_dotenv

# 환경변수 로드
load_dotenv()
SENDER_EMAIL = os.getenv("GMAIL_EMAIL")
SENDER_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
API_KEY = os.getenv("API_KEY")

# =========================
# 1. API KEY (공통)
# =========================
# API_KEY = "SpBjtAWYaN2aNqknzNlYA4wmB0Amo1IcAM8cNfrU5NAk8nuKEtNGw5dNf6MtkVwliAuKWek+4YG8zjq+osj2og=="

# =========================
# 2. 데이터 수집/가공 클래스
# =========================
FONT_PATHS = [
    "C:/Windows/Fonts/malgun.ttf",
    "./fonts/NanumGothic.ttf",
    "./NanumGothic.ttf"
]

def register_korean_font():
    for font_path in FONT_PATHS:
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
            return 'KoreanFont'
    return None

class AirQualityReporter:
    def __init__(self, api_key):
        self.api_key = api_key
        self.base_url = "http://apis.data.go.kr/B552584/ArpltnInforInqireSvc"
        self.station_url = "http://apis.data.go.kr/B552584/MsrstnInfoInqireSvc"

    def get_sido_list(self):
        # 실제 API로 가져와도 되지만, 고정값 사용
        return [
            '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
            '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
        ]

    def fetch_station_list(self, sido_name):
        url = f"{self.station_url}/getMsrstnList"
        params = {
            "serviceKey": self.api_key,
            "returnType": "json",
            "numOfRows": 100,
            "pageNo": 1,
            "addr": sido_name
        }
        response = requests.get(url, params=params)
        data = response.json()
        return pd.DataFrame(data['response']['body']['items'])

    def fetch_station_dust_data(self, station_name):
        url = f"{self.base_url}/getMsrstnAcctoRltmMesureDnsty"
        params = {
            "serviceKey": self.api_key,
            "returnType": "json",
            "numOfRows": 24,
            "pageNo": 1,
            "stationName": station_name,
            "dataTerm": "DAILY",
            "ver": "1.3"
        }
        response = requests.get(url, params=params)
        data = response.json()
        return pd.DataFrame(data['response']['body']['items'])

    def fetch_sido_dust_data(self, sido_name):
        url = f"{self.base_url}/getCtprvnRltmMesureDnsty"
        params = {
            'serviceKey': self.api_key,
            'returnType': 'json',
            'numOfRows': '100',
            'pageNo': '1',
            'sidoName': sido_name,
            'ver': '1.0'
        }
        response = requests.get(url, params=params)
        data = response.json()
        return pd.DataFrame(data['response']['body']['items'])

    def get_air_quality_grade_text(self, grade):
        grade_map = {
            '1': '좋음',
            '2': '보통',
            '3': '나쁨',
            '4': '매우나쁨'
        }
        return grade_map.get(str(grade), '정보없음')

    def create_pdf_report(self, df, sido_name):
        font_name = register_korean_font()
        styles = getSampleStyleSheet()
        if font_name:
            styles.add(ParagraphStyle(name='KoreanTitle', fontName=font_name, fontSize=18, leading=22))
            styles.add(ParagraphStyle(name='KoreanNormal', fontName=font_name, fontSize=10, leading=14))
            title_style = styles['KoreanTitle']
            normal_style = styles['KoreanNormal']
        else:
            title_style = styles['Title']
            normal_style = styles['Normal']

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []

        # 제목
        title = Paragraph(f"{sido_name} 지역 미세먼지 현황 리포트", title_style)
        story.append(title)
        story.append(Spacer(1, 12))

        # 생성 시간
        now = datetime.now().strftime("%Y년 %m월 %d일 %H시 %M분")
        date_para = Paragraph(f"생성일시: {now}", normal_style)
        story.append(date_para)
        story.append(Spacer(1, 20))

        # 테이블 데이터 준비 (한글 컬럼)
        table_data = [['측정소명', 'PM10', 'PM2.5', 'PM10등급', 'PM2.5등급', '측정시간']]
        for _, row in df.iterrows():
            table_data.append([
                row.get('stationName', ''),
                row.get('pm10Value', ''),
                row.get('pm25Value', ''),
                row.get('pm10Grade', ''),
                row.get('pm25Grade', ''),
                row.get('dataTime', '')
            ])

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name if font_name else 'Helvetica'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), font_name if font_name else 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
        doc.build(story)
        buffer.seek(0)
        return buffer

    def create_excel_report(self, df, sido_name):
        col_map = {
            'stationName': '측정소명',
            'sidoName': '시도명',
            'pm10Value': 'PM10',
            'pm25Value': 'PM2.5',
            'pm10Grade': 'PM10등급',
            'pm25Grade': 'PM2.5등급',
            'so2Value': '아황산가스(SO2)',
            'coValue': '일산화탄소(CO)',
            'o3Value': '오존(O3)',
            'no2Value': '이산화질소(NO2)',
            'dataTime': '측정시간',
            'khaiValue': '통합대기환경지수',
            'khaiGrade': '통합등급'
        }
        cols = [col for col in col_map.keys() if col in df.columns]
        df_kor = df[cols].copy()
        # 통합등급만 한글로 변환
        if 'khaiGrade' in df_kor.columns:
            df_kor['khaiGrade'] = df_kor['khaiGrade'].apply(self.get_air_quality_grade_text)
        df_kor = df_kor.rename(columns=col_map)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_kor.to_excel(writer, index=False)
        output.seek(0)
        return output

    def create_station_pdf_report(self, df, station_name):
        # 한글 폰트 등록
        font_name = register_korean_font()
        styles = getSampleStyleSheet()
        if font_name:
            styles.add(ParagraphStyle(name='KoreanTitle', fontName=font_name, fontSize=18, leading=22))
            styles.add(ParagraphStyle(name='KoreanNormal', fontName=font_name, fontSize=10, leading=14))
            title_style = styles['KoreanTitle']
            normal_style = styles['KoreanNormal']
        else:
            title_style = styles['Title']
            normal_style = styles['Normal']

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []

        # 제목
        title = Paragraph(f"{station_name} 측정소 미세먼지 리포트", title_style)
        story.append(title)
        story.append(Spacer(1, 12))

        # 생성 시간
        now = datetime.now().strftime("%Y년 %m월 %d일 %H시 %M분")
        date_para = Paragraph(f"생성일시: {now}", normal_style)
        story.append(date_para)
        story.append(Spacer(1, 20))

        # 테이블 데이터 준비 (한글 컬럼)
        table_data = [['측정시각', 'PM10', 'PM2.5']]
        for _, row in df.iterrows():
            table_data.append([
                row.get('dataTime', ''),
                row.get('pm10Value', ''),
                row.get('pm25Value', '')
            ])

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name if font_name else 'Helvetica'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), font_name if font_name else 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
        doc.build(story)
        buffer.seek(0)
        return buffer

    def create_station_excel_report(self, df):
        # 컬럼명 한글로 매핑
        col_map = {
            'dataTime': '측정시각',
            'pm10Value': 'PM10',
            'pm25Value': 'PM2.5',
            'pm10Grade': 'PM10등급',
            'pm25Grade': 'PM2.5등급',
            'so2Value': '아황산가스(SO2)',
            'coValue': '일산화탄소(CO)',
            'o3Value': '오존(O3)',
            'no2Value': '이산화질소(NO2)',
            'khaiValue': '통합대기환경지수',
            'khaiGrade': '통합등급',
            'stationName': '측정소명',
            'sidoName': '시도명'
        }
        # 기존 컬럼 순서대로, 존재하는 컬럼만 한글로 변환
        cols = [col for col in col_map.keys() if col in df.columns]
        df_kor = df[cols].rename(columns=col_map)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_kor.to_excel(writer, index=False)
        output.seek(0)
        return output

    def send_email_report(self, to_email, subject, body, attachments, sender_email, sender_password):
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = to_email
        msg.set_content(body)
        for attachment, filename in attachments:
            attachment.seek(0)
            msg.add_attachment(
                attachment.read(),
                maintype='application',
                subtype='octet-stream',
                filename=filename
            )
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)

# =========================
# 3. Streamlit UI
# =========================
def main():
    st.set_page_config(page_title="미세먼지 통합 리포트", layout="wide")
    st.title("🌫️ 미세먼지 통합 리포트 시스템")
    reporter = AirQualityReporter(API_KEY)

    # ---- 사이드바 ----
    with st.sidebar:
        st.header("설정")
        mode = st.radio("모드 선택", ["지역별 미세먼지 현황", "측정소별 미세먼지 현황"])
        sido = st.selectbox("지역(시/도) 선택", reporter.get_sido_list())
        station = None
        if mode == "측정소별 미세먼지 현황" and sido:
            stations = reporter.fetch_station_list(sido)
            station_names = stations['stationName'].tolist()
            station = st.selectbox("측정소 선택", station_names)
        if st.button("데이터 조회 및 리포트 생성", key="fetch_btn"):
            if mode == "지역별 미세먼지 현황":
                df = reporter.fetch_sido_dust_data(sido)
                st.session_state['aq_df'] = df
                st.session_state['aq_mode'] = 'sido'
                st.session_state['aq_title'] = f"{sido} 미세먼지 리포트"
                st.session_state['aq_station'] = None
            else:
                if not station:
                    st.warning("측정소를 선택하세요.")
                else:
                    df = reporter.fetch_station_dust_data(station)
                    st.session_state['aq_df'] = df
                    st.session_state['aq_mode'] = 'station'
                    st.session_state['aq_title'] = f"{station} 미세먼지 리포트"
                    st.session_state['aq_station'] = station
        send_email = st.checkbox("이메일로 리포트 발송")
        recipient_email = None
        send_email_btn = False
        if send_email:
            recipient_email = st.text_input("받는 사람 이메일")
            send_email_btn = st.button("이메일 보내기", key="send_email_btn")

    # ---- 메인 화면 ----
    if 'aq_df' in st.session_state and st.session_state['aq_df'] is not None:
        df = st.session_state['aq_df']
        mode = st.session_state.get('aq_mode', 'sido')
        report_title = st.session_state.get('aq_title', '미세먼지 리포트')
        station = st.session_state.get('aq_station', None)

        if mode == "sido":
            st.subheader(report_title)
            st.dataframe(df[['stationName', 'pm10Value', 'pm25Value', 'pm10Grade', 'pm25Grade', 'dataTime']])
            fig = px.bar(df.head(10), x='stationName', y='pm10Value', title=f"{report_title} - PM10 상위 10개 측정소")
            st.plotly_chart(fig)
            # PDF
            pdf_buffer = reporter.create_pdf_report(df, sido)
            st.session_state['pdf_buffer'] = pdf_buffer # 세션에 저장
            st.session_state['excel_buffer'] = None # 초기화
            st.download_button("PDF 리포트 다운로드", data=pdf_buffer, file_name=f"{sido}_미세먼지_리포트.pdf")

            # Excel
            excel_buffer = reporter.create_excel_report(df, sido)
            st.session_state['excel_buffer'] = excel_buffer # 세션에 저장
            st.download_button("Excel 리포트 다운로드", data=excel_buffer, file_name=f"{sido}_미세먼지_리포트.xlsx")
        else:
            st.subheader(report_title)
            st.dataframe(df)
            fig = px.line(df, x='dataTime', y=['pm10Value', 'pm25Value'], title=f"{report_title} - 시간별 미세먼지 농도")
            st.plotly_chart(fig)
            pdf_buffer = reporter.create_station_pdf_report(df, station)
            excel_buffer = reporter.create_station_excel_report(df)
            st.session_state['pdf_buffer'] = pdf_buffer # 세션에 저장
            st.session_state['excel_buffer'] = excel_buffer # 세션에 저장
            st.download_button("PDF 리포트 다운로드", data=pdf_buffer, file_name=f"{station}_미세먼지_리포트.pdf")
            st.download_button("Excel 리포트 다운로드", data=excel_buffer, file_name=f"{station}_미세먼지_리포트.xlsx")

        # 메일 발송
        if send_email and send_email_btn and SENDER_EMAIL and SENDER_PASSWORD and recipient_email:
            try:
                attachments = [
                    (st.session_state['pdf_buffer'], f"{report_title}.pdf"),
                    (st.session_state['excel_buffer'], f"{report_title}.xlsx")
                ]
                reporter.send_email_report(
                    recipient_email, report_title, "첨부파일을 확인하세요.",
                    attachments, SENDER_EMAIL, SENDER_PASSWORD
                )
                st.success("메일 발송 완료!")
            except Exception as e:
                st.error(f"메일 발송 실패: {e}")

if __name__ == "__main__":
    main()
