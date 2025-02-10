# -*- coding: utf-8 -*-

import pandas as pd
from xgboost import XGBRegressor
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from datetime import datetime, timedelta
import xlwings as xw
import matplotlib

# 한글 폰트 설정
matplotlib.rcParams['font.family'] = 'Malgun Gothic'

class GypsumApp:
    def __init__(self, root):
        self.root = root
        self.root.title("석고본드 지연제 예측 프로그램")
        self.root.geometry("1000x800")

        # 변수 초기화
        self.model_delay = None
        self.model_work = None
        self.data = None
        self.fig = None
        self.canvas = None
        self.r2_delay = None
        self.r2_work = None

        self.setup_ui()

    # ... (나머지 코드는 동일)
    def setup_ui(self):
        # 메인 프레임 설정
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 컨트롤 프레임
        control_frame = ttk.LabelFrame(main_frame, text="입력 설정")
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        # 엑셀 파일 로드 버튼
        self.load_excel_button = ttk.Button(
            control_frame, 
            text="엑셀 파일 로드", 
            command=self.load_data_from_excel
        )
        self.load_excel_button.pack(pady=5)

        # 원료 산지 선택
        origin_frame = ttk.Frame(control_frame)
        origin_frame.pack(fill=tk.X, pady=5)
        
        self.origin_label = ttk.Label(origin_frame, text="원료 산지:")
        self.origin_label.pack(side=tk.LEFT, padx=5)
        
        self.origin_var = tk.StringVar()
        self.origin_combo = ttk.Combobox(
            origin_frame, 
            textvariable=self.origin_var, 
            state="readonly"
        )
        self.origin_combo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.origin_combo['values'] = []
        self.origin_combo.bind("<<ComboboxSelected>>", self.update_model)

        # 소석고 초결시간 입력
        time_frame = ttk.Frame(control_frame)
        time_frame.pack(fill=tk.X, pady=5)
        
        self.setting_time_label = ttk.Label(time_frame, text="소석고 초결시간(min):")
        self.setting_time_label.pack(side=tk.LEFT, padx=5)
        
        self.setting_time_entry = ttk.Entry(time_frame)
        self.setting_time_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # 목표 가사시간 범위 입력
        target_frame = ttk.Frame(control_frame)
        target_frame.pack(fill=tk.X, pady=5)
        
        self.target_label = ttk.Label(target_frame, text="목표 가사시간 범위(min):")
        self.target_label.pack(side=tk.LEFT, padx=5)
        
        self.min_target_entry = ttk.Entry(target_frame, width=8)
        self.min_target_entry.pack(side=tk.LEFT, padx=5)
        self.min_target_entry.insert(0, "200")  # 기본값 설정
        
        self.target_separator = ttk.Label(target_frame, text="~")
        self.target_separator.pack(side=tk.LEFT)
        
        self.max_target_entry = ttk.Entry(target_frame, width=8)
        self.max_target_entry.pack(side=tk.LEFT, padx=5)
        self.max_target_entry.insert(0, "250")  # 기본값 설정

        # 예측 버튼
        self.predict_button = ttk.Button(
            control_frame, 
            text="예측", 
            command=self.predict
        )
        self.predict_button.pack(pady=5)

        # 결과 라벨
        self.result_label = ttk.Label(control_frame, text="", wraplength=400)
        self.result_label.pack(pady=5)

        # 그래프 프레임
        self.graph_frame = ttk.LabelFrame(main_frame, text="데이터 그래프")
        self.graph_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def load_data_from_excel(self):
        try:
            file_path = filedialog.askopenfilename(
                title="엑셀 파일 선택",
                filetypes=(("Excel files", "*.xlsx;*.xls"), ("all files", "*.*"))
            )
            
            if not file_path:
                return

            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            sheet = wb.sheets[0]
            
            # 데이터 로드
            excel_data = sheet.range('A1').expand('table').value
            headers = excel_data[0]
            data_rows = excel_data[1:]
            
            self.data = pd.DataFrame(data_rows, columns=headers)
            
            # 날짜 처리
            self.data['날짜'] = pd.to_datetime(self.data['날짜'], format='mixed')
            
            # 필수 컬럼 확인
            required_columns = ['원료 산지', '소석고 초결시간(min)', '지연제 투입량(kg)', '품질 가사시간(min)']
            missing_columns = [col for col in required_columns if col not in self.data.columns]
            
            if missing_columns:
                raise ValueError(f"다음 컬럼이 없습니다: {', '.join(missing_columns)}")

            # 콤보박스 값 업데이트
            self.origin_combo['values'] = sorted(self.data['원료 산지'].unique())
            
            messagebox.showinfo("성공", "엑셀 파일 데이터 로드 완료")

        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 로드 중 오류 발생:\n{str(e)}")
        
        finally:
            try:
                wb.close()
                app.quit()
            except:
                pass

    def update_model(self, event=None):
        selected_origin = self.origin_var.get()
        if not selected_origin:
            return

        if self.data is None:
            messagebox.showerror("오류", "데이터가 먼저 로드되어야 합니다")
            return

        try:
            # 선택된 산지의 데이터만 필터링
            origin_data = self.data[self.data['원료 산지'] == selected_origin].copy()
            
            if origin_data.empty:
                messagebox.showerror("오류", "선택된 산지의 데이터가 존재하지 않습니다.")
                return

            # 마지막 날짜 찾기
            last_date = origin_data['날짜'].max()
            seven_days_ago = last_date - timedelta(days=7)
            
            # 최근 7일 데이터 필터링
            filtered_data = origin_data[origin_data['날짜'] > seven_days_ago].copy()
            
            if filtered_data.empty:
                messagebox.showerror("오류", "선택된 산지의 최근 7일 데이터가 존재하지 않습니다.")
                return

            # 결측치 처리
            filtered_data = filtered_data.dropna(subset=['소석고 초결시간(min)', '지연제 투입량(kg)', '품질 가사시간(min)'])

            X = filtered_data[['소석고 초결시간(min)']]
            y_delay = filtered_data['지연제 투입량(kg)']
            y_work = filtered_data['품질 가사시간(min)']

            # XGBoost 모델 학습
            self.model_delay = XGBRegressor(
                n_estimators=100,
                learning_rate=0.1,
                max_depth=3,
                random_state=42
            )
            self.model_work = XGBRegressor(
                n_estimators=100,
                learning_rate=0.1,
                max_depth=3,
                random_state=42
            )
            
            self.model_delay.fit(X, y_delay)
            self.model_work.fit(X, y_work)

            # R² 점수 계산
            self.r2_delay = self.model_delay.score(X, y_delay)
            self.r2_work = self.model_work.score(X, y_work)

            messagebox.showinfo("성공", f"{selected_origin} 모델 학습 완료\n(데이터 기간: {seven_days_ago.strftime('%Y-%m-%d')} ~ {last_date.strftime('%Y-%m-%d')})")
            self.plot_recent_data(selected_origin)

        except Exception as e:
            messagebox.showerror("오류", f"모델 학습 중 오류가 발생했습니다\n{str(e)}")

    def predict(self):
        if not self.validate_inputs():
            return

        try:
            setting_time = float(self.setting_time_entry.get())
            min_target = float(self.min_target_entry.get())
            max_target = float(self.max_target_entry.get())
            
            predicted_delay = self.model_delay.predict([[setting_time]])[0]
            predicted_work = self.model_work.predict([[setting_time]])[0]
            
            result_text = (
                f"예측 결과\n"
                f"지연제 투입량: {predicted_delay:.2f} kg (R² = {self.r2_delay:.3f})\n"
                f"예상 가사시간: {predicted_work:.2f} 분 (R² = {self.r2_work:.3f})\n"
                f"목표 가사시간 범위: {min_target:.0f}~{max_target:.0f} 분"
            )
            
            # 가사시간이 목표 범위를 벗어났는지 체크
            if predicted_work < min_target:
                result_text += "\n\n※ 주의: 예상 가사시간이 목표 최소 시간보다 짧습니다"
            elif predicted_work > max_target:
                result_text += "\n\n※ 주의: 예상 가사시간이 목표 최대 시간을 초과합니다"
            
            self.result_label.config(text=result_text)
            self.plot_recent_data()

        except Exception as e:
            messagebox.showerror("오류", f"예측 중 오류가 발생했습니다\n{str(e)}")

    def validate_inputs(self):
        if not self.origin_var.get():
            messagebox.showerror("오류", "원료 산지을 선택해주세요.")
            return False
            
        if not self.setting_time_entry.get():
            messagebox.showerror("오류", "소석고 초결시간을 입력해주세요.")
            return False
            
        if not self.min_target_entry.get() or not self.max_target_entry.get():
            messagebox.showerror("오류", "목표 가사시간 범위를 입력해주세요.")
            return False
            
        try:
            setting_time = float(self.setting_time_entry.get())
            min_target = float(self.min_target_entry.get())
            max_target = float(self.max_target_entry.get())
            
            if min_target >= max_target:
                messagebox.showerror("오류", "최대 가사시간은 최소 가사시간보다 커야 합니다.")
                return False
                
        except ValueError:
            messagebox.showerror("오류", "모든 입력값은 숫자여야 합니다.")
            return False
            
        if self.model_delay is None or self.model_work is None:
            messagebox.showerror("오류", "모델이 학습되지 않았습니다")
            return False
            
        return True

    def plot_recent_data(self, selected_origin=None):
        if selected_origin is None:
            selected_origin = self.origin_var.get()

        if self.data is None or not selected_origin:
            return

        try:
            # 선택된 산지의 데이터만 필터링
            origin_data = self.data[self.data['원료 산지'] == selected_origin].copy()
            
            # 최근 30개 데이터 필터링
            filtered_data = origin_data.sort_values('날짜', ascending=False).head(30).sort_values('날짜')
            
            if filtered_data.empty:
                messagebox.showinfo("알림", "데이터가 존재하지 않습니다.")
                return

            self.create_plot(filtered_data, selected_origin)

        except Exception as e:
            messagebox.showerror("오류", f"그래프 생성 중 오류가 발생했습니다\n{str(e)}")

    def create_plot(self, filtered_data, selected_origin):
        if self.fig:
            plt.close(self.fig)

        self.fig = Figure(figsize=(10, 6), dpi=100)
        ax1 = self.fig.add_subplot(111)
        ax2 = ax1.twinx()

        # 날짜를 숫자 인덱스로 변환
        date_indices = range(len(filtered_data))
        dates = filtered_data['날짜'].values
        delay_data = filtered_data['지연제 투입량(kg)'].values
        work_data = filtered_data['품질 가사시간(min)'].values

        # 데이터 플로팅 (x축을 인덱스로 사용)
        line1, = ax1.plot(date_indices, delay_data, 'b-o', label='지연제 투입량(kg)')
        line2, = ax2.plot(date_indices, work_data, 'r-x', label='가사시간(분)')

        ax1.set_xlabel('날짜')
        ax1.set_ylabel('지연제 투입량(kg)', color='blue')
        ax2.set_ylabel('가사시간(분)', color='red')

        ax1.tick_params(axis='y', labelcolor='blue')
        ax2.tick_params(axis='y', labelcolor='red')

        # X축 눈금 설정 - datetime64를 datetime으로 변환
        ax1.set_xticks(date_indices)
        date_labels = [pd.Timestamp(d).strftime('%Y-%m-%d') for d in dates]
        ax1.set_xticklabels(date_labels, rotation=45)

        # 그래프 제목에 데이터 기간 표시
        start_date = pd.Timestamp(filtered_data['날짜'].min()).strftime('%Y-%m-%d')
        end_date = pd.Timestamp(filtered_data['날짜'].max()).strftime('%Y-%m-%d')
        ax1.set_title(f'{selected_origin} 지연제 투입량과 가사시간\n({start_date} ~ {end_date})')

        # 범례 설정
        lines = [line1, line2]
        labels = [l.get_label() for l in lines]
        ax1.legend(lines, labels, loc='upper left')

        # 레이아웃 조정
        self.fig.tight_layout()

        if self.canvas:
            self.canvas.get_tk_widget().destroy()

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = GypsumApp(root)
    root.mainloop()
