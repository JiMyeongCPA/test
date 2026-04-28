
import dart_fss as dart
import pandas as pd
import re
import tkinter as tk
from tkinter import messagebox

# 1. 통합 컬럼 세척 함수 (시점 및 기간 모두 대응)
def final_clean_columns(col):
    col = str(col)

    # [기간] 20240101-20241231 형식 우선 추출
    period_match = re.search(r'\d{8}-\d{8}', col)
    if period_match:
        return period_match.group()

    # [시점] 20241231 형식 추출
    date_match = re.search(r'\d{8}', col)
    if date_match:
        return date_match.group()

    # 핵심 키워드 보존
    keywords = ['concept_id', 'label_ko', 'label_en', 'class0', 'class1', 'class2', 'class3', 'class4']
    for word in keywords:
        if word in col:
            return word

    # 앞부분 잘라내고 마지막 유의미한 정보 추출
    if '_' in col:
        parts = col.split('_')
        if parts[-1] in ['id', 'ko', 'en'] and len(parts) >= 2:
            return "_".join(parts[-2:])
        return parts[-1]

    return col

# 2. 통합 추출 및 시트별 저장 로직
def run_total_extraction():
    corp_name = entry.get()
    api_key = api_entry.get()  # 사용자로부터 API 키를 직접 가져옵니다.

    if not corp_name:
        messagebox.showwarning("경고", "기업명을 입력해주세요.")
        return

    if not api_key:
        messagebox.showwarning("경고", "DART API 키를 입력해주세요.")
        return

    try:
        status_label.config(text=f"'{corp_name}' 전체 재무제표 통합 추출 중...")
        root.update()

        # 사용자로부터 입력받은 API 키 설정
        dart.set_api_key(api_key=api_key)

        corp_list = dart.get_corp_list()
        target_corp = corp_list.find_by_corp_name(corp_name, exactly=True)

        if not target_corp:
            messagebox.showerror("에러", f"'{corp_name}'을(를) 찾을 수 없습니다.")
            return

        corp_obj = target_corp[0]
        # 시작 날짜는 필요에 따라 수정하여 사용하세요
        fs = corp_obj.extract_fs(bgn_de='20230101')

        # 엑셀 파일 생성 (ExcelWriter 활용)
        filename = f"{corp_name}_재무제표_통합.xlsx"

        fs_map = {
            'bs': '재무상태표',
            'is': '손익계산서',
            'cis': '포괄손익계산서',
            'cf': '현금흐름표'
        }

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            found_any = False
            for key, sheet_name in fs_map.items():
                df = fs[key]

                if df is not None:
                    # 1단계: 평탄화
                    df.columns = ["_".join([str(i) for i in col if i]) if isinstance(col, tuple) else str(col) for col in df.columns]
                    # 2단계: 세척
                    df.columns = [final_clean_columns(c) for c in df.columns]

                    # 3단계: [선택] 숫자 가독성을 위해 조 단위 환산 (금액 컬럼만)
                    amount_cols = [c for c in df.columns if re.search(r'\d{8}', c)]
                    df[amount_cols] = df[amount_cols] / 10**12

                    # 4단계: 시트로 저장
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    found_any = True

        if found_any:
            messagebox.showinfo("성공", f"'{filename}' 파일에 모든 제표가 시트별로 저장되었습니다!")
        else:
            messagebox.showwarning("알림", "추출된 데이터가 없습니다.")

        status_label.config(text="대기 중...")

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))
        status_label.config(text="에러 발생")

# 3. GUI 구성
root = tk.Tk()
root.title("CPA용 재무제표 통합 마스터")
root.geometry("400x320") # 입력창 추가로 인해 세로 길이를 늘렸습니다.

# 기업명 입력란
tk.Label(root, text="기업명 입력:", font=("돋움", 10)).pack(pady=(20, 0))
entry = tk.Entry(root, width=30, font=("돋움", 12))
entry.pack(pady=5)
entry.insert(0, "삼성전자")

# API 키 입력란 (추가된 부분)
tk.Label(root, text="DART API 키 입력:", font=("돋움", 10)).pack(pady=(10, 0))
api_entry = tk.Entry(root, width=30, font=("돋움", 12), show="*") # 보안을 위해 별표 표시
api_entry.pack(pady=5)

# 실행 버튼
btn = tk.Button(root, text="통합 엑셀 파일 생성", command=run_total_extraction, 
                bg="#1A237E", fg="white", font=("돋움", 10, "bold"))
btn.pack(pady=25)

status_label = tk.Label(root, text="대기 중...", fg="gray")
status_label.pack()

root.mainloop()

