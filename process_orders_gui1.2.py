{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 5,
   "id": "2746f78d-c8f8-4601-8101-b88b7a1172d5",
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "import pandas as pd\n",
    "import tkinter as tk\n",
    "from tkinter import filedialog, messagebox\n",
    "\n",
    "def process_orders(input_file, output_file):\n",
    "    try:\n",
    "        df = pd.read_excel(input_file, dtype={'주문번호': str})\n",
    "\n",
    "        # 필요한 컬럼만 추출\n",
    "        df_filtered = df[['주문번호', '상품명', '옵션정보', '클레임상태', '수량']]\n",
    "\n",
    "        # '챌린지 신청'이 포함된 주문번호 필터링\n",
    "        challenge_orders = df_filtered[df_filtered['상품명'].str.contains('챌린지 신청', na=False)].copy()\n",
    "\n",
    "        # '신청 안함'이 포함된 주문번호 필터링\n",
    "        non_challenge_orders = df_filtered[df_filtered['상품명'].str.contains('신청 안함', na=False)].copy()\n",
    "\n",
    "        # '부트캠프'가 포함된 주문번호 필터링\n",
    "        bootcamp_orders = df_filtered[df_filtered['상품명'].str.contains('부트캠프', na=False)].copy()\n",
    "\n",
    "        # 해당 주문번호에 대한 모든 주문 정보 조회\n",
    "        challenge_order_numbers = challenge_orders['주문번호'].unique()\n",
    "        non_challenge_order_numbers = non_challenge_orders['주문번호'].unique()\n",
    "        bootcamp_order_numbers = bootcamp_orders['주문번호'].unique()\n",
    "\n",
    "        all_challenge_orders = df_filtered[df_filtered['주문번호'].isin(challenge_order_numbers)].copy()\n",
    "        all_non_challenge_orders = df_filtered[df_filtered['주문번호'].isin(non_challenge_order_numbers)].copy()\n",
    "        all_bootcamp_orders = df_filtered[df_filtered['주문번호'].isin(bootcamp_order_numbers)].copy()\n",
    "\n",
    "        # 가격 정보 사전 정의 (한 줄로 요약)\n",
    "        price_dict = {\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 1. 강남역 호랑이 삼겹' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 2. 수원 왕갈비통 닭목살' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 3. 기사식당 최강 제육' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 4. 춘천 들깨 닭갈비' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 5. 수랏간 삼치 솥밥' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 시그니처 6. 항아리 차돌 된장' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 1. 수비드 통삼겹 된장덮밥' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 2. 수비드 통삼겹 들기름 막국수' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 3. 훈제오리 들깨 크림 리조또' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 4. 우삼겹 규동' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 5. 우삼겹 오일 파스타' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 6. B.T.S 치킨치즈 리조또' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 7. 소고기 버섯 들깨 덮밥' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x1회 배송): 오리지널 8. 저당 두부면 라자냐' : 8500,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 1. 강남역 호랑이 삼겹' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 2. 수원 왕갈비통 닭목살' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 3. 기사식당 최강 제육' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 4. 춘천 들깨 닭갈비' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 5. 수랏간 삼치 솥밥' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 시그니처 6. 항아리 차돌 된장' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 1. 수비드 통삼겹 된장덮밥' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 2. 수비드 통삼겹 들기름 막국수' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 3. 훈제오리 들깨 크림 리조또' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 4. 우삼겹 규동' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 5. 우삼겹 오일 파스타' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 6. B.T.S 치킨치즈 리조또' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 7. 소고기 버섯 들깨 덮밥' : 17000,\n",
    "            '도시락 14개 자유선택  (14개x2회 배송): 오리지널 8. 저당 두부면 라자냐' : 17000,\n",
    "            '시그니처: 1. 강남역 호랑이 삼겹' : 8500,\n",
    "            '시그니처: 2. 수원 왕갈비통 닭목살' : 8500,\n",
    "            '시그니처: 3. 기사식당 최강 제육' : 8500,\n",
    "            '시그니처: 4. 춘천 들깨 닭갈비' : 8500,\n",
    "            '시그니처: 5. 수랏간 삼치 솥밥' : 8500,\n",
    "            '시그니처: 6. 항아리 차돌 된장' : 8500,\n",
    "            '시그니처: 도시락 6종 x  1팩' : 52000,\n",
    "            '시그니처: 도시락 6종 x 1팩 (조합 선택 불가)' : 52000,\n",
    "            '오리지널: 1. 수비드 통삼겹 된장덮밥' : 8500,\n",
    "            '오리지널: 2. 수비드 통삼겹 들기름 막국수' : 8500,\n",
    "            '오리지널: 3. 훈제오리 들깨 크림 리조또' : 8500,\n",
    "            '오리지널: 4. 우삼겹 규동' : 8500,\n",
    "            '오리지널: 5. 우삼겹 오일 파스타' : 8500,\n",
    "            '오리지널: 6. B.T.S 치킨치즈 리조또' : 8500,\n",
    "            '오리지널: 7. 소고기 버섯 들깨 덮밥' : 8500,\n",
    "            '오리지널: 8. 저당 두부면 라자냐' : 8500,\n",
    "            '오리지널: 도시락 8종 x 1팩 (조합 선택 불가)' : 68000,\n",
    "            '패키지 선택: [구독 패키지 1] 황금비율 1:2:7' : 238000,\n",
    "            '패키지 선택: [구독 패키지 2] 실패없는 베스트셀러' : 240000,\n",
    "            '패키지 선택: [구독 패키지 3] 김윤겸이 매일 먹는 도시락' : 244000\n",
    "        }\n",
    "\n",
    "        # 금액 컬럼 추가 (옵션정보 기준 매핑 및 수량 곱하기)\n",
    "        all_challenge_orders['단가'] = all_challenge_orders['옵션정보'].map(price_dict)\n",
    "        all_non_challenge_orders['단가'] = all_non_challenge_orders['옵션정보'].map(price_dict)\n",
    "        all_bootcamp_orders['단가'] = all_bootcamp_orders['옵션정보'].map(price_dict)\n",
    "\n",
    "        all_challenge_orders['금액'] = all_challenge_orders['단가'] * all_challenge_orders['수량']\n",
    "        all_non_challenge_orders['금액'] = all_non_challenge_orders['단가'] * all_non_challenge_orders['수량']\n",
    "        all_bootcamp_orders['금액'] = all_bootcamp_orders['단가'] * all_bootcamp_orders['수량']\n",
    "\n",
    "        # 금액이 없는 경우를 확인하고, NaN 값을 0으로 대체\n",
    "        all_challenge_orders['금액'] = all_challenge_orders['금액'].fillna(0)\n",
    "        all_non_challenge_orders['금액'] = all_non_challenge_orders['금액'].fillna(0)\n",
    "        all_bootcamp_orders['금액'] = all_bootcamp_orders['금액'].fillna(0)\n",
    "\n",
    "        # 각 주문번호별 총 구매 금액 계산 (상품명 및 클레임상태 포함)\n",
    "        challenge_total = all_challenge_orders.groupby('주문번호').agg({'상품명': 'first', '금액': 'sum', '클레임상태': 'first'}).reset_index()\n",
    "        non_challenge_total = all_non_challenge_orders.groupby('주문번호').agg({'상품명': 'first', '금액': 'sum', '클레임상태': 'first'}).reset_index()\n",
    "        bootcamp_total = all_bootcamp_orders.groupby('주문번호').agg({'상품명': 'first', '금액': 'sum', '클레임상태': 'first'}).reset_index()\n",
    "\n",
    "        # 결과 결합\n",
    "        result = pd.concat([challenge_total, non_challenge_total, bootcamp_total], ignore_index=True)\n",
    "\n",
    "        # 첫 번째 시트: 매출 요약 테이블\n",
    "        일반_매출 = result[~result['상품명'].str.contains('챌린지 신청|신청 안함|부트캠프', na=False)]['금액'].sum()\n",
    "        챌린지_매출 = result[result['상품명'].str.contains('챌린지 신청', na=False)]['금액'].sum()\n",
    "        신청_안함_매출 = result[result['상품명'].str.contains('신청 안함', na=False)]['금액'].sum()\n",
    "        부트캠프_매출 = result[result['상품명'].str.contains('부트캠프', na=False)]['금액'].sum()\n",
    "        전체_도시락_매출 = 일반_매출 + 챌린지_매출 + 신청_안함_매출 + 부트캠프_매출\n",
    "\n",
    "        summary_data = {\n",
    "            '일반 매출': [일반_매출, 0, 0, 일반_매출],\n",
    "            '챌린지 매출': [0, 챌린지_매출, 신청_안함_매출, 챌린지_매출 + 신청_안함_매출],\n",
    "            '부트캠프 매출': [0, 0, 부트캠프_매출, 부트캠프_매출],\n",
    "            '전체 도시락 매출': [일반_매출, 챌린지_매출, 부트캠프_매출, 전체_도시락_매출]\n",
    "        }\n",
    "\n",
    "        summary_df = pd.DataFrame(summary_data, index=['일반 매출', '챌린지 매출', '부트캠프 매출', '전체 도시락 매출'])\n",
    "\n",
    "        summary_df.loc['비율'] = summary_df.loc['전체 도시락 매출'] / 전체_도시락_매출 * 100\n",
    "        summary_df.loc['비율'] = summary_df.loc['비율'].map('{:.0f}%'.format)\n",
    "\n",
    "        # 결과를 엑셀 파일로 저장\n",
    "        with pd.ExcelWriter(output_file) as writer:\n",
    "            summary_df.to_excel(writer, sheet_name='매출 요약')\n",
    "            result.to_excel(writer, sheet_name='기존 매출', index=False)\n",
    "            bootcamp_total.to_excel(writer, sheet_name='부트캠프 매출', index=False)\n",
    "\n",
    "        messagebox.showinfo(\"성공\", f\"파일이 성공적으로 생성되었습니다: {output_file}\")\n",
    "\n",
    "    except Exception as e:\n",
    "        messagebox.showerror(\"오류\", f\"오류가 발생했습니다: {e}\")\n",
    "\n",
    "def select_input_file():\n",
    "    input_file = filedialog.askopenfilename(title=\"입력 파일 선택\", filetypes=[(\"Excel files\", \"*.xlsx *.xls\")])\n",
    "    input_entry.delete(0, tk.END)\n",
    "    input_entry.insert(0, input_file)\n",
    "\n",
    "def select_output_file():\n",
    "    output_file = filedialog.asksaveasfilename(defaultextension=\".xlsx\", title=\"출력 파일 저장\", filetypes=[(\"Excel files\", \"*.xlsx\")])\n",
    "    output_entry.delete(0, tk.END)\n",
    "    output_entry.insert(0, output_file)\n",
    "\n",
    "def run_process():\n",
    "    input_file = input_entry.get()\n",
    "    output_file = output_entry.get()\n",
    "    if not input_file or not output_file:\n",
    "        messagebox.showwarning(\"경고\", \"입력 파일과 출력 파일을 모두 선택해야 합니다.\")\n",
    "        return\n",
    "    process_orders(input_file, output_file)\n",
    "\n",
    "# GUI 설정\n",
    "root = tk.Tk()\n",
    "root.title(\"Process Orders\")\n",
    "\n",
    "tk.Label(root, text=\"입력 파일:\").grid(row=0, column=0, padx=10, pady=10)\n",
    "input_entry = tk.Entry(root, width=50)\n",
    "input_entry.grid(row=0, column=1, padx=10, pady=10)\n",
    "tk.Button(root, text=\"파일 선택\", command=select_input_file).grid(row=0, column=2, padx=10, pady=10)\n",
    "\n",
    "tk.Label(root, text=\"출력 파일:\").grid(row=1, column=0, padx=10, pady=10)\n",
    "output_entry = tk.Entry(root, width=50)\n",
    "output_entry.grid(row=1, column=1, padx=10, pady=10)\n",
    "tk.Button(root, text=\"파일 저장\", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)\n",
    "\n",
    "tk.Button(root, text=\"실행\", command=run_process).grid(row=2, column=0, columnspan=3, pady=20)\n",
    "\n",
    "root.mainloop()\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "6ae06077-998a-4697-9e31-372f0ae6604e",
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.11.7"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
