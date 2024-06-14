import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def process_orders(input_file, output_file):
    try:
        df = pd.read_excel(input_file, dtype={'주문번호': str})

        # 필요한 컬럼만 추출
        df_filtered = df[['주문번호', '상품명', '옵션정보', '클레임상태']]

        # '챌린지 신청'이 포함된 주문번호 필터링
        challenge_orders = df_filtered[df_filtered['상품명'].str.contains('챌린지 신청', na=False)].copy()

        # '신청 안함'이 포함된 주문번호 필터링
        non_challenge_orders = df_filtered[df_filtered['상품명'].str.contains('신청 안함', na=False)].copy()

        # 해당 주문번호에 대한 모든 주문 정보 조회
        challenge_order_numbers = challenge_orders['주문번호'].unique()
        non_challenge_order_numbers = non_challenge_orders['주문번호'].unique()

        all_challenge_orders = df_filtered[df_filtered['주문번호'].isin(challenge_order_numbers)].copy()
        all_non_challenge_orders = df_filtered[df_filtered['주문번호'].isin(non_challenge_order_numbers)].copy()

        # 가격 정보 사전 정의
        price_dict = {
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 1. 강남역 호랑이 삼겹' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 2. 수원 왕갈비통 닭목살' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 3. 기사식당 최강 제육' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 4. 춘천 들깨 닭갈비' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 5. 수랏간 삼치 솥밥' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 시그니처 6. 항아리 차돌 된장' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 1. 수비드 통삼겹 된장덮밥' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 2. 수비드 통삼겹 들기름 막국수' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 3. 훈제오리 들깨 크림 리조또' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 4. 우삼겹 규동' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 5. 우삼겹 오일 파스타' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 6. B.T.S 치킨치즈 리조또' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 7. 소고기 버섯 들깨 덮밥' : 8500,
        '도시락 14개 자유선택  (14개x1회 배송): 오리지널 8. 저당 두부면 라자냐' : 8500,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 1. 강남역 호랑이 삼겹' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 2. 수원 왕갈비통 닭목살' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 3. 기사식당 최강 제육' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 4. 춘천 들깨 닭갈비' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 5. 수랏간 삼치 솥밥' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 시그니처 6. 항아리 차돌 된장' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 1. 수비드 통삼겹 된장덮밥' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 2. 수비드 통삼겹 들기름 막국수' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 3. 훈제오리 들깨 크림 리조또' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 4. 우삼겹 규동' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 5. 우삼겹 오일 파스타' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 6. B.T.S 치킨치즈 리조또' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 7. 소고기 버섯 들깨 덮밥' : 17000,
        '도시락 14개 자유선택  (14개x2회 배송): 오리지널 8. 저당 두부면 라자냐' : 17000,
        '시그니처: 1. 강남역 호랑이 삼겹' : 8500,
        '시그니처: 2. 수원 왕갈비통 닭목살' : 8500,
        '시그니처: 3. 기사식당 최강 제육' : 8500,
        '시그니처: 4. 춘천 들깨 닭갈비' : 8500,
        '시그니처: 5. 수랏간 삼치 솥밥' : 8500,
        '시그니처: 6. 항아리 차돌 된장' : 8500,
        '시그니처: 도시락 6종 x  1팩' : 52000,
        '시그니처: 도시락 6종 x 1팩 (조합 선택 불가)' : 52000,
        '오리지널: 1. 수비드 통삼겹 된장덮밥' : 8500,
        '오리지널: 2. 수비드 통삼겹 들기름 막국수' : 8500,
        '오리지널: 3. 훈제오리 들깨 크림 리조또' : 8500,
        '오리지널: 4. 우삼겹 규동' : 8500,
        '오리지널: 5. 우삼겹 오일 파스타' : 8500,
        '오리지널: 6. B.T.S 치킨치즈 리조또' : 8500,
        '오리지널: 7. 소고기 버섯 들깨 덮밥' : 8500,
        '오리지널: 8. 저당 두부면 라자냐' : 8500,
        '오리지널: 도시락 8종 x 1팩 (조합 선택 불가)' : 68000,
        '패키지 선택: [구독 패키지 1] 황금비율 1:2:7' : 238000,
        '패키지 선택: [구독 패키지 2] 실패없는 베스트셀러' : 240000,
        '패키지 선택: [구독 패키지 3] 김윤겸이 매일 먹는 도시락' : 244000
        }

        # 금액 컬럼 추가 (옵션정보 기준 매핑)
        all_challenge_orders['금액'] = all_challenge_orders['옵션정보'].map(price_dict)
        all_non_challenge_orders['금액'] = all_non_challenge_orders['옵션정보'].map(price_dict)

        # 금액이 없는 경우를 확인하고, NaN 값을 0으로 대체
        all_challenge_orders['금액'] = all_challenge_orders['금액'].fillna(0)
        all_non_challenge_orders['금액'] = all_non_challenge_orders['금액'].fillna(0)

        # 각 주문번호별 총 구매 금액 계산 (상품명 및 클레임상태 포함)
        challenge_total = all_challenge_orders.groupby('주문번호').agg({'상품명': 'first', '금액': 'sum', '클레임상태': 'first'}).reset_index()
        non_challenge_total = all_non_challenge_orders.groupby('주문번호').agg({'상품명': 'first', '금액': 'sum', '클레임상태': 'first'}).reset_index()

        # 결과 결합
        result = pd.concat([challenge_total, non_challenge_total], ignore_index=True)

        # 결과를 엑셀 파일로 저장
        with pd.ExcelWriter(output_file) as writer:
            result.to_excel(writer, sheet_name='Orders Total', index=False)

        messagebox.showinfo("성공", f"파일이 성공적으로 생성되었습니다: {output_file}")

    except Exception as e:
        messagebox.showerror("오류", f"오류가 발생했습니다: {e}")

def select_input_file():
    input_file = filedialog.askopenfilename(title="입력 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_file)

def select_output_file():
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="출력 파일 저장", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_file)

def run_process():
    input_file = input_entry.get()
    output_file = output_entry.get()
    if not input_file or not output_file:
        messagebox.showwarning("경고", "입력 파일과 출력 파일을 모두 선택해야 합니다.")
        return
    process_orders(input_file, output_file)

# GUI 설정
root = tk.Tk()
root.title("Process Orders")

tk.Label(root, text="입력 파일:").grid(row=0, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="파일 선택", command=select_input_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="출력 파일:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="파일 저장", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="실행", command=run_process).grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
