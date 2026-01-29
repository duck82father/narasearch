from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5 import uic
import sys
import requests
import pandas as pd
from datetime import datetime
import os
import sqlite3
import webbrowser
import re 

# ==========================================
# [커스텀] 날짜/시간 선택 팝업 클래스
# ==========================================
class DateTimePopup(QFrame):
    dateTimeSelected = pyqtSignal(QDateTime)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.setStyleSheet("""
            QFrame {
                background: white;
                border: 1px solid #a0a0a0;
                border-radius: 5px;
            }
            QLabel {
                font-size: 13px;
                font-weight: bold;
                color: #333;
                border: none;
            }
            QSpinBox, QComboBox {
                min-height: 25px;
                font-size: 13px;
                font-weight: bold;
                padding: 2px;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        self.calendar = QCalendarWidget(self)
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setStyleSheet("""
            QCalendarWidget QWidget { alternate-background-color: #f0f0f0; }
            QAbstractItemView:enabled { 
                color: #333;  
                selection-background-color: #2563eb; 
                selection-color: white; 
            }
            QToolButton {
                icon-size: 20px;
                font-weight: bold;
                border: none;
            }
        """)
        main_layout.addWidget(self.calendar)

        time_frame = QFrame(self)
        time_frame.setStyleSheet("QFrame { border: none; }")
        time_layout = QHBoxLayout(time_frame)
        time_layout.setContentsMargins(5, 5, 5, 5)
        
        self.ampm = QComboBox(self)
        self.ampm.addItems(["오전", "오후"])
        self.ampm.setFixedWidth(60)

        self.lbl_hour = QLabel("시간")
        self.hour = QSpinBox(self)
        self.hour.setRange(1, 12)
        self.hour.setAlignment(Qt.AlignCenter)
        self.hour.setFixedWidth(50)

        self.lbl_minute = QLabel("분")
        self.minute = QSpinBox(self)
        self.minute.setRange(0, 50)
        self.minute.setSingleStep(10)
        self.minute.setAlignment(Qt.AlignCenter)
        self.minute.setFixedWidth(50)

        time_layout.addWidget(self.ampm)
        time_layout.addSpacing(10)
        time_layout.addWidget(self.lbl_hour)
        time_layout.addWidget(self.hour)
        time_layout.addSpacing(10)
        time_layout.addWidget(self.lbl_minute)
        time_layout.addWidget(self.minute)
        
        main_layout.addWidget(time_frame)

        btn_confirm = QPushButton("선택 완료", self)
        btn_confirm.setCursor(Qt.PointingHandCursor)
        btn_confirm.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 3px;
                border: none;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
        """)
        btn_confirm.clicked.connect(self.emit_datetime)
        main_layout.addWidget(btn_confirm)

    def set_initial_datetime(self, dt):
        self.calendar.setSelectedDate(dt.date())
        time = dt.time()
        h = time.hour()
        m = time.minute()
        
        if h < 12:
            self.ampm.setCurrentIndex(0)
            self.hour.setValue(h if h != 0 else 12)
        else:
            self.ampm.setCurrentIndex(1)
            self.hour.setValue(h - 12 if h != 12 else 12)
            
        m_rounded = (m // 10) * 10
        self.minute.setValue(m_rounded)

    def emit_datetime(self):
        date = self.calendar.selectedDate()
        h_val = self.hour.value()
        is_pm = (self.ampm.currentIndex() == 1)
        
        if is_pm and h_val != 12:
            h_val += 12
        elif not is_pm and h_val == 12:
            h_val = 0
            
        m_val = self.minute.value()
        final_time = QTime(h_val, m_val)
        dt = QDateTime(date, final_time)
        
        self.dateTimeSelected.emit(dt)
        self.close()

# 환경 설정
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
UI_PATH = "./ui/narasearchv1.ui"
DB_PATH = os.path.join(BASE_DIR, './ui/narasearchdata.db') 

# ==========================================
# Pandas 모델 클래스 (폰트 14pt)
# ==========================================
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        if role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        
        elif role == Qt.FontRole:
            font = QFont()
            font.setPointSize(14) 
            
            # 제목 컬럼들 볼드 처리
            col_name = self._data.columns[index.column()]
            if col_name in ['입찰공고명', '입찰개시일시', '품명(사업명)', '접수일시']:
                font.setBold(True)
            return font
            
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
            
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

# ==========================================
# 검색 워커 (스레드)
# ==========================================
class SearchWorker(QThread):
    result_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    
    def __init__(self, url_base, params_base, keywords, title_field='bidNtceNm'):
        super().__init__()
        self.url_base = url_base
        self.params_base = params_base
        self.keywords = keywords
        self.title_field = title_field 

    def run(self):
        all_items = []
        page_no = 1
        rows_per_page = 999 

        try:
            while True:
                current_params = f"{self.params_base}&pageNo={page_no}&numOfRows={rows_per_page}"
                full_url = self.url_base + current_params
                
                res = requests.get(full_url)
                
                if res.status_code != 200:
                    self.error_signal.emit(f"서버 접속 오류: {res.status_code}")
                    return

                try:
                    data = res.json()
                except:
                    self.error_signal.emit(f"데이터 파싱 실패: {res.text[:300]}")
                    return

                result_code = None
                if 'response' in data and 'header' in data['response']:
                    result_code = data['response']['header'].get('resultCode')
                elif 'nkoneps.com.response.ResponseError' in data:
                    result_code = data['nkoneps.com.response.ResponseError'].get('header', {}).get('resultCode')
                elif 'resultCode' in data:
                    result_code = data.get('resultCode')

                if str(result_code) == "07":
                    self.error_signal.emit("최대 검색기간을 초과하였습니다.\n31일 이내로 검색해주세요.")
                    return

                if 'response' not in data or 'body' not in data['response']:
                    break 

                body = data['response']['body']
                items = body.get('items')
                total_count = body.get('totalCount', 0)

                if not items:
                    break 

                if isinstance(items, dict):
                    items = [items]

                all_items.extend(items)

                if len(all_items) >= int(total_count):
                    break
                
                page_no += 1
                if page_no > 20: 
                    break

            if not all_items:
                self.error_signal.emit("검색 결과가 없습니다.")
                return

            final_items = []
            
            if self.keywords:
                for item in all_items:
                    title = item.get(self.title_field, '')
                    title_clean = title.replace(" ", "").lower()
                    
                    is_match = True
                    for k in self.keywords:
                        k_clean = k.replace(" ", "").lower()
                        if k_clean not in title_clean:
                            is_match = False
                            break
                    
                    if is_match:
                        final_items.append(item)
            else:
                final_items = all_items

            if not final_items:
                if len(self.keywords) > 1:
                    detail_msg = f"상세 조건('{', '.join(self.keywords[1:])}')이"
                else:
                    detail_msg = "조건이"
                self.error_signal.emit(f"'{self.keywords[0]}' 관련 데이터 {len(all_items)}개를 가져왔으나,\n{detail_msg} 포함된 공고는 없습니다.")
                return

            self.result_signal.emit({'items': final_items})

        except Exception as e:
            self.error_signal.emit(f"시스템 에러: {str(e)}")

# ==========================================
# 메인 위젯
# ==========================================
class MainWidget(QWidget):
    def __init__(self):
        QWidget.__init__(self, None)
        uic.loadUi(os.path.join(BASE_DIR, UI_PATH), self)
        
        self.init_db() 
        self.initUI()
        
        self.main_search_timer = QTimer(self)
        self.main_search_timer.setSingleShot(True) 
        self.main_search_timer.timeout.connect(self.save_settings_to_db) 

        self.shortcut_timers = {} 

        self.startButton.clicked.connect(self.search_start_main)
        self.resetButton.clicked.connect(self.search_reset)
        self.saveButton.clicked.connect(self.search_save)
        self.endButton.clicked.connect(self.search_end)

        self.search_keyword.returnPressed.connect(self.search_start_main)
        self.search_servicekey.returnPressed.connect(self.search_start_main)
        
        if hasattr(self.expiredkeydate, 'returnPressed'):
            self.expiredkeydate.returnPressed.connect(self.search_start_main)

        self.search_keyword.textChanged.connect(lambda: self.main_search_timer.start(2000))
        self.search_servicekey.textChanged.connect(self.save_settings_to_db)
        self.expiredkeydate.textChanged.connect(self.save_settings_to_db)

        self.threeweeksButton.clicked.connect(self.set_date_range_3weeks)
        
        if hasattr(self, 'noticeButton'):
            self.noticeButton.clicked.connect(self.show_notice)

        self.tableView.doubleClicked.connect(self.open_link)
        
        # [추가] 컨텍스트 메뉴(우클릭 메뉴) 정책 설정
        self.tableView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableView.customContextMenuRequested.connect(self.show_context_menu)

        for i in range(10):
            line_edit = getattr(self, f'Shortcut_{i}', None)
            btn = getattr(self, f'startShortcutButton_{i}', None)

            if line_edit and btn:
                line_edit.setAlignment(Qt.AlignCenter)
                timer = QTimer(self)
                timer.setSingleShot(True)
                timer.timeout.connect(lambda idx=i: self.save_shortcut_actual(idx))
                self.shortcut_timers[i] = timer

                line_edit.textChanged.connect(lambda text, idx=i: self.shortcut_timers[idx].start(2000))
                line_edit.textChanged.connect(lambda text, le=line_edit: self.update_shortcut_style(le))

                line_edit.returnPressed.connect(lambda idx=i: self.search_start_shortcut(idx))
                btn.clicked.connect(lambda checked, idx=i: self.search_start_shortcut(idx))

        self.load_settings_from_db()
        self.set_date_range_3weeks()
        
        self.df2 = None
        self.display_df = None 

    def open_datetime_popup(self, target):
        try:
            self.datetime_popup.dateTimeSelected.disconnect()
        except:
            pass
        self.datetime_popup.dateTimeSelected.connect(target.setDateTime)
        pos = target.mapToGlobal(QPoint(0, target.height()))
        self.datetime_popup.move(pos)
        self.datetime_popup.show()

    def show_notice(self):
        msg_text = """

■ 문화의 창 나라장터 통합 검색기 ] LJH ver.1 

- 나라장터 입찰공고, 사전규격 API를 이용한 검색 프로그램입니다.


■ 기능 안내 ■ ----------------------------------------------------------------------

- 검색창 오른쪽 노란색 버튼에서 [입찰공고] / [사전규격] 선택 후
  검색어를 입력하시면 원하는 검색이 가능합니다.

- 검색 기간이 반드시 지정되며, 최대 31일 간의 검색이 가능합니다.

- 저장단어 검색(총 10칸)에 원하는 검색어를 저장하여
  같은 단어를 손쉽게 자주 검색할 수 있습니다.

- 검색 결과는 엑셀 파일로 저장 가능합니다.


■ 검색 Tip ■ ----------------------------------------------------------------------

- 검색어 입력 시, [특징이 되는 단어] + [지역명] 순으로 입력합니다.
  ex) '의왕어린이철도축제' → 검색어 '어린이, 철도, 의왕'


■ 나라장터 API 인증키 ■ ----------------------------------------------------------------------

- 나라장터 API 인증키는 공공데이터포털(www.data.go.kr)에서 누구나
  가입 후 신청할 수 있으며, 다음 두 서비스를 모두 신청해야합니다.
  - '조달청_나라장터 사전규격정보서비스'
  - '조달청_나라장터 입찰공고정보서비스'
"""
        QMessageBox.information(self, "프로그램 정보", msg_text)

    def update_shortcut_style(self, line_edit):
        text = line_edit.text().strip()
        if text:
            line_edit.setStyleSheet("background-color: #d2ffd2;")
        else:
            line_edit.setStyleSheet("background-color: white;")

    def initUI(self):
        self.setWindowTitle('나라장터 통합 검색기')
        self.setWindowIcon(QIcon('./ui/icon.png'))

        # ----------------------------------------------------------------
        # 테이블 뷰 스타일 설정 (padding 10px)
        # ----------------------------------------------------------------
        self.tableView.setStyleSheet("""
            QTableView::item {
                padding: 18px;
            }
        """)

        # comboBox 초기화
        if hasattr(self, 'comboBox'):
            self.comboBox.clear()
            self.comboBox.addItems([" 입찰공고", " 사전규격"])
        else:
            print("경고: 'comboBox' 객체를 찾을 수 없습니다. UI 파일에 해당 객체가 있는지 확인해주세요.")

        date_style = """
        QDateTimeEdit {
            background-color: #ffffff;
        }
        QDateTimeEdit::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 30px;
            border: none;
            background-color: transparent;
        }
        """

        self.search_startdate.setStyleSheet(date_style)
        self.search_enddate.setStyleSheet(date_style)

        self.search_startdate.setCalendarPopup(False)
        self.search_enddate.setCalendarPopup(False)

        self.datetime_popup = DateTimePopup(self)

        self.search_startdate.mousePressEvent = lambda e: self.open_datetime_popup(self.search_startdate)
        self.search_enddate.mousePressEvent = lambda e: self.open_datetime_popup(self.search_enddate)

        self.show()

    def load_settings_from_db(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT api_key, expired_date, use_3weeks FROM settings WHERE id=1")
            row = cursor.fetchone()
            if row:
                api_key, expired_date, use_3weeks = row
                self.search_servicekey.blockSignals(True)
                self.expiredkeydate.blockSignals(True)
                self.search_servicekey.setText(api_key)
                self.expiredkeydate.setText(expired_date)
                self.search_servicekey.blockSignals(False)
                self.expiredkeydate.blockSignals(False)
            
            cursor.execute("SELECT idx, keyword FROM shortcuts")
            rows = cursor.fetchall()
            for idx, keyword in rows:
                line_edit = getattr(self, f'Shortcut_{idx}', None)
                if line_edit:
                    line_edit.blockSignals(True) 
                    line_edit.setText(keyword)
                    line_edit.blockSignals(False)
                    self.update_shortcut_style(line_edit)
            conn.close()
        except Exception as e:
            print(f"DB 로드 에러: {e}")

    def init_db(self):
        db_dir = os.path.dirname(DB_PATH)
        if not os.path.exists(db_dir):
            try:
                os.makedirs(db_dir)
            except Exception as e:
                print(f"폴더 생성 실패: {e}")
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                id INTEGER PRIMARY KEY,
                api_key TEXT,
                expired_date TEXT,
                use_3weeks INTEGER
            )
        ''')
        cursor.execute("SELECT count(*) FROM settings")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO settings (id, api_key, expired_date, use_3weeks) VALUES (1, '', '', 0)")
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS shortcuts (
                idx INTEGER PRIMARY KEY,
                keyword TEXT
            )
        ''')
        for i in range(10):
            cursor.execute("SELECT count(*) FROM shortcuts WHERE idx=?", (i,))
            if cursor.fetchone()[0] == 0:
                cursor.execute("INSERT INTO shortcuts (idx, keyword) VALUES (?, '')", (i,))
        conn.commit()
        conn.close()

    def save_settings_to_db(self):
        try:
            api_key = self.search_servicekey.text().strip()
            expired_date = self.expiredkeydate.text().strip()
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE settings 
                SET api_key=?, expired_date=?
                WHERE id=1
            ''', (api_key, expired_date))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"DB 저장 에러: {e}")

    def save_shortcut_actual(self, idx):
        line_edit = getattr(self, f'Shortcut_{idx}', None)
        if not line_edit: return
        text = line_edit.text() 
        self.save_shortcut_to_db(idx, text)

    def save_shortcut_to_db(self, idx, text):
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("UPDATE shortcuts SET keyword=? WHERE idx=?", (text, idx))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Shortcut 저장 에러: {e}")

    def set_date_range_3weeks(self):
        now = QDateTime.currentDateTime()
        three_weeks_ago = now.addDays(-21) 
        self.search_enddate.setDateTime(now)
        self.search_startdate.setDateTime(three_weeks_ago)

    def open_link(self, index):
        if self.display_df is None: return
        row = index.row()
        try:
            url = self.display_df.iloc[row]['상세링크']
            if url and str(url).startswith('http'):
                webbrowser.open(str(url))
            else:
                self.search_situation.setText("유효한 상세 링크가 없습니다.")
        except Exception as e:
            print(f"링크 열기 실패: {e}")

    # =========================================================================
    # [추가] 컨텍스트 메뉴 (우클릭 다운로드) 핸들러
    # =========================================================================
    def show_context_menu(self, pos):
        if self.display_df is None: return
        
        index = self.tableView.indexAt(pos)
        if not index.isValid(): return

        row = index.row()
        menu = QMenu(self)
        
        # 1. 파일 URL 컬럼 키 탐색
        if hasattr(self, 'comboBox'):
            category = self.comboBox.currentText().strip()
        else:
            category = "입찰공고"

        found_files = []
        
        if category == "입찰공고":
            # [cite_start]입찰공고 파일 컬럼: ntceSpecDocUrl1 ~ ntceSpecDocUrl10 [cite: 1]
            for i in range(1, 11):
                col_key = f'ntceSpecDocUrl{i}'
                if col_key in self.display_df.columns:
                    url = self.display_df.iloc[row][col_key]
                    if url and str(url).strip() != '' and str(url) != 'nan':
                        found_files.append((f"첨부파일 {i} 다운로드", str(url)))
        else:
            # [cite_start]사전규격 파일 컬럼: specDocFileUrl1 ~ specDocFileUrl5 [cite: 1]
            for i in range(1, 6):
                col_key = f'specDocFileUrl{i}'
                if col_key in self.display_df.columns:
                    url = self.display_df.iloc[row][col_key]
                    if url and str(url).strip() != '' and str(url) != 'nan':
                        found_files.append((f"규격문서 {i} 다운로드", str(url)))

        # 2. 메뉴 액션 추가
        if found_files:
            label_action = QAction("--- 다운로드 목록 ---", self)
            label_action.setEnabled(False)
            menu.addAction(label_action)
            
            for label, url in found_files:
                action = QAction(label, self)
                # 람다식에서 변수 바인딩(url=url) 주의
                action.triggered.connect(lambda checked, u=url: webbrowser.open(u))
                menu.addAction(action)
        else:
            no_action = QAction("다운로드 가능한 규격문서가 없습니다.", self)
            no_action.setEnabled(False)
            menu.addAction(no_action)

        menu.exec_(self.tableView.mapToGlobal(pos))

    def search_start_main(self):
        keyword = self.search_keyword.text().strip()
        self.execute_search(keyword)
        
    def search_start_shortcut(self, idx):
        line_edit = getattr(self, f'Shortcut_{idx}', None)
        if line_edit:
            keyword = line_edit.text().strip()
            self.execute_search(keyword)

    def execute_search(self, keyword_input):
        if not self.startButton.isEnabled():
            return

        if hasattr(self, 'comboBox'):
            category = self.comboBox.currentText().strip()
        else:
            category = "입찰공고"

        self.search_situation.setText(f"[{category}] '{keyword_input}' 검색 중입니다... (데이터량에 따라 시간이 걸릴 수 있습니다)")
        
        keywords_list = [k for k in re.split(r'[, ]+', keyword_input.strip()) if k]
        if not keywords_list:
            QMessageBox.warning(self, "알림", "검색어를 입력해주세요.")
            self.search_situation.setText("검색어를 입력해주세요.")
            return
        
        primary_keyword = keywords_list[0]
        self.startButton.setEnabled(False) 
        self.tableView.setModel(None)

        start_data_str = self.search_startdate.dateTime().toString('yyyyMMddHHmm')
        end_date_str = self.search_enddate.dateTime().toString('yyyyMMddHHmm')
        service_key = self.search_servicekey.text().strip()

        if not service_key:
            QMessageBox.warning(self, "오류", "API 인증키를 입력해주세요.")
            self.startButton.setEnabled(True)
            return
        
        # API URL 분기
        url_base = ""
        params_base = ""
        title_field = ""

        if category == "입찰공고":
            url_base = 'http://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch?'
            params_base = f'inqryDiv=1&inqryBgnDt={start_data_str}&inqryEndDt={end_date_str}&bidNtceNm={primary_keyword}&type=json&serviceKey={service_key}'
            title_field = 'bidNtceNm'
        
        elif category == "사전규격":
            url_base = 'http://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServcPPSSrch?'
            params_base = f'inqryDiv=1&inqryBgnDt={start_data_str}&inqryEndDt={end_date_str}&prdctClsfcNoNm={primary_keyword}&type=json&serviceKey={service_key}'
            title_field = 'prdctClsfcNoNm'
        
        else:
            QMessageBox.warning(self, "오류", "검색 유형(입찰공고/사전규격)을 선택해주세요.")
            self.startButton.setEnabled(True)
            return

        print(f"검색 시작: [{category}] 키워드='{primary_keyword}'")

        self.worker = SearchWorker(url_base, params_base, keywords_list, title_field)
        self.worker.result_signal.connect(self.handle_success)
        self.worker.error_signal.connect(self.handle_error)
        self.worker.start()

    def handle_success(self, data):
        bid_data = data['items']
        df = pd.DataFrame(bid_data)
        
        if hasattr(self, 'comboBox'):
            category = self.comboBox.currentText().strip()
        else:
            category = "입찰공고"

        if 'asignBdgtAmt' in df.columns:
            def format_money(x):
                try:
                    if not x or str(x).strip() == '': return '-'
                    amt = float(str(x).replace(',', ''))
                    eok = int(amt // 100000000)
                    man = int((amt % 100000000) // 10000)
                    result = ""
                    if eok > 0: result += f"{eok}억"
                    if man > 0:
                        if eok > 0: result += " "
                        result += f"{man:,}만원"
                    if result == "": result = f"{int(amt):,}원"
                    return result
                except:
                    return str(x)
            df['asignBdgtAmt'] = df['asignBdgtAmt'].apply(format_money)

        # ---------------------------------------------------------------------
        # 카테고리별 컬럼 처리
        # ---------------------------------------------------------------------
        # [수정] 파일 다운로드 URL 컬럼 추가 로직
        file_cols = []
        
        if category == "입찰공고":
            # [cite_start]파일 컬럼 후보군 정의 (1~10) [cite: 1]
            potential_file_cols = [f'ntceSpecDocUrl{i}' for i in range(1, 11)]
            
            target_cols = ['bidNtceNo', 'bidNtceDt', 'bidNtceNm', 'ntceInsttNm', 'cntrctCnclsMthdNm', 'bidBeginDt', 'bidClseDt', 'asignBdgtAmt']
            if 'bidNtceDtlUrl' in df.columns: target_cols.append('bidNtceDtlUrl')
            
            # 존재하는 파일 컬럼만 target_cols에 추가 (화면에는 숨길 예정)
            for c in potential_file_cols:
                if c in df.columns:
                    target_cols.append(c)
                    file_cols.append(c)
            
            target_cols = [c for c in target_cols if c in df.columns]
            df1 = df[target_cols].copy()
            
            rename_map = {
                'bidNtceNo':'입찰공고번호', 'bidNtceDt':'입찰공고일시', 'bidNtceNm':'입찰공고명', 
                'ntceInsttNm':'공고기관명',  'cntrctCnclsMthdNm':'계약체결방법명', 
                'bidBeginDt':'입찰개시일시', 'bidClseDt':'입찰마감일시', 'asignBdgtAmt':'배정예산금액',
                'bidNtceDtlUrl': '상세링크'
            }
            df1 = df1.rename(columns=rename_map)
            
            save_cols = ['bidNtceNo', 'ntceKindNm', 'bidNtceDt', 'bidNtceNm', 'ntceInsttNm', 'dminsttNm', 'bidMethdNm', 'cntrctCnclsMthdNm', 'bidBeginDt',
                         'bidClseDt', 'bidPrtcptLmtYn', 'asignBdgtAmt', 'sucsfbidLwltRate', 'sucsfbidMthdNm']
            save_cols = [c for c in save_cols if c in df.columns]
            df2 = df[save_cols].copy()
            df2 = df2.rename(columns={
                'bidNtceNo':'입찰공고번호', 'ntceKindNm':'공고종류명', 'bidNtceDt':'입찰공고일시', 
                'bidNtceNm':'입찰공고명', 'ntceInsttNm':'공고기관명', 'dminsttNm':'수요기관명',
                'bidMethdNm':'입찰방식명', 'cntrctCnclsMthdNm':'계약체결방법명', 
                'bidBeginDt':'입찰개시일시', 'bidClseDt':'입찰마감일시', 'bidPrtcptLmtYn':'입찰참가제한여부',
                'asignBdgtAmt':'배정예산금액', 'sucsfbidLwltRate':'낙찰하한율', 'sucsfbidMthdNm':'낙찰방법명'
            })

        else: # "사전규격"
            # [cite_start]파일 컬럼 후보군 정의 (1~5) [cite: 1]
            potential_file_cols = [f'specDocFileUrl{i}' for i in range(1, 6)]
            
            target_cols = ['bfSpecRgstNo', 'rcptDt', 'prdctClsfcNoNm', 'orderInsttNm', 'rlDminsttNm', 'opninRgstClseDt', 'asignBdgtAmt']
            
            # 존재하는 파일 컬럼만 target_cols에 추가
            for c in potential_file_cols:
                if c in df.columns:
                    target_cols.append(c)
                    file_cols.append(c)

            target_cols = [c for c in target_cols if c in df.columns]
            
            df1 = df[target_cols].copy()
            rename_map = {
                'bfSpecRgstNo': '사전규격등록번호', 'rcptDt': '접수일시', 'prdctClsfcNoNm': '품명(사업명)',
                'orderInsttNm': '발주기관명', 'rlDminsttNm': '실수요기관명', 
                'opninRgstClseDt': '의견등록마감일시', 'asignBdgtAmt': '배정예산금액'
            }
            df1 = df1.rename(columns=rename_map)

            save_cols = ['bfSpecRgstNo', 'refNo', 'rcptDt', 'prdctClsfcNoNm', 'orderInsttNm', 'rlDminsttNm', 
                         'opninRgstClseDt', 'asignBdgtAmt', 'ofclNm', 'ofclTelNo', 'dlvrTmlmtDt']
            save_cols = [c for c in save_cols if c in df.columns]
            df2 = df[save_cols].copy()
            df2 = df2.rename(columns={
                'bfSpecRgstNo': '사전규격등록번호', 'refNo': '참조번호', 'rcptDt': '접수일시',
                'prdctClsfcNoNm': '품명(사업명)', 'orderInsttNm': '발주기관명', 'rlDminsttNm': '실수요기관명',
                'opninRgstClseDt': '의견등록마감일시', 'asignBdgtAmt': '배정예산금액',
                'ofclNm': '담당자명', 'ofclTelNo': '담당자전화번호', 'dlvrTmlmtDt': '납품기한일시'
            })

        # 공통: 상세링크 컬럼이 없으면 빈 컬럼 생성 (에러 방지용)
        if '상세링크' not in df1.columns: df1['상세링크'] = ''

        self.display_df = df1 
        self.df2 = df2 
        
        model = PandasModel(df1)
        self.tableView.setModel(model)
        
        # 상세링크 숨기기
        if '상세링크' in df1.columns:
            link_col_index = df1.columns.get_loc('상세링크')
            self.tableView.setColumnHidden(link_col_index, True)

        # [추가] 파일 URL 컬럼들도 화면에서 숨기기 (데이터는 가지고 있음)
        for col in file_cols:
            if col in df1.columns:
                col_idx = df1.columns.get_loc(col)
                self.tableView.setColumnHidden(col_idx, True)

        self.tableView.resizeColumnsToContents()
        self.tableView.resizeRowsToContents()

        self.search_situation.setText(f"[{category}] 검색 완료: {len(df)}건이 검색되었습니다.")
        self.startButton.setEnabled(True)

    def handle_error(self, msg):
        if "검색 결과가 없습니다" in msg:
            self.search_situation.setText("검색 결과가 없습니다.")
        else:
            self.search_situation.setText("검색 결과가 없습니다.")
            QMessageBox.warning(self, "알림", msg)
        self.startButton.setEnabled(True)

    def search_reset(self):
        self.tableView.setModel(None)
        self.search_keyword.setText("")
        self.search_situation.setText("리셋 되었습니다.")
        self.df2 = None
        self.display_df = None

    def search_save(self):
        if self.df2 is None or self.df2.empty:
            QMessageBox.information(self, "알림", "저장할 데이터가 없습니다.")
            return
        
        keyword_part = self.search_keyword.text().strip()
        if not keyword_part: keyword_part = "통합"

        if hasattr(self, 'comboBox'):
            category = self.comboBox.currentText()
        else:
            category = "결과"

        default_name = f"{category}_{keyword_part}_검색결과_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "엑셀 파일 저장", default_name, "Excel Files (*.xlsx)")

        if save_path:
            try:
                self.df2.to_excel(save_path, index=False)
                self.search_situation.setText(f"저장 완료: {os.path.basename(save_path)}")
                QMessageBox.information(self, "성공", "파일이 성공적으로 저장되었습니다.\n확인을 누르면 엑셀 파일이 열립니다.")
                os.startfile(save_path)
            except Exception as e:
                self.search_situation.setText("저장 실패")
                QMessageBox.critical(self, "실패", f"파일 저장 중 오류가 발생했습니다.\n{str(e)}")

    def search_end(self):
        sys.exit()

if __name__ == '__main__':
    QApplication.setStyle("fusion")
    app = QApplication(sys.argv)
    main_Widget = MainWidget()
    main_Widget.showMaximized()
    sys.exit(app.exec_())