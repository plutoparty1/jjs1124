import tkinter as tk
from tkinter import messagebox

# 면적별 등급/금액 기준 (min_area, max_area, grade, amount)
RULES = [
    (0,    50,    0, 0),
    (50,   100,   1, 4400),
    (100,  200,   2, 7920),
    (200,  300,   3, 10780),
    (300,  500,   4, 13640),
    (500,  1000,  5, 17160),
    (1000, float('inf'), 6, 22000),
]

def get_base_grade_amount(area):
    """면적에 따른 기본 등급/금액 반환"""
    for min_a, max_a, g, amt in RULES:
        if min_a <= area < max_a:
            return g, amt
    return 0, 0

def calc_fee(event=None):
    # 면적 숫자 확인
    try:
        area = float(entry_area.get())
        if area < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("입력 오류", "면적을 0 이상 숫자로 입력해 주세요.")
        return

    is_rest = var_restaurant.get() == 1   # 휴게음식점 여부
    is_eupmyeon = var_eupmyeon.get() == 1 # 읍·면·동 여부

    # 일반음식점이면 납부 안함
    if not is_rest:
        label_grade.config(text="등급: 해당 없음 (일반음식점)")
        label_result.config(text="납부금액: 0 원")
        return

    # 기본 등급 계산
    grade, _ = get_base_grade_amount(area)

    # 읍·면·동이면 등급 1단계 하향 (최하 1등급, 0등급으로는 안 내려감)
    if is_eupmyeon and grade > 1:
        grade = max(1, grade - 1)

    # 조정된 등급으로 금액 다시 찾기
    amount = 0
    for _, _, g, amt in RULES:
        if g == grade:
            amount = amt
            break

    label_grade.config(text=f"등급: {grade} 등급")
    label_result.config(text=f"납부금액: {amount:,} 원")

# ----- Tkinter GUI -----
root = tk.Tk()
root.title("공연권료 납부액 계산기")

# 창 크기 & 위치 (가로 320 x 세로 180)
root.geometry("320x180")
root.resizable(False, False)

# 면적 입력
frame_top = tk.Frame(root, padx=10, pady=10)
frame_top.pack(fill="x")

tk.Label(frame_top, text="매장 면적 (㎡):").grid(row=0, column=0, sticky="w")
entry_area = tk.Entry(frame_top, width=10)
entry_area.grid(row=0, column=1, padx=5)
entry_area.insert(0, "0")

# 체크박스들
frame_mid = tk.Frame(root, padx=10)
frame_mid.pack(fill="x")

var_restaurant = tk.IntVar(value=1)  # 기본값: 휴게음식점 체크
var_eupmyeon = tk.IntVar(value=0)

chk_rest = tk.Checkbutton(frame_mid, text="휴게음식점", variable=var_restaurant)
chk_eupmyeon = tk.Checkbutton(frame_mid, text="읍·면·동 소재지", variable=var_eupmyeon)

chk_rest.grid(row=0, column=0, sticky="w")
chk_eupmyeon.grid(row=0, column=1, sticky="w", padx=10)

# 계산 버튼
frame_btn = tk.Frame(root, padx=10, pady=5)
frame_btn.pack(fill="x")

btn_calc = tk.Button(frame_btn, text="계산", width=10, command=calc_fee)
btn_calc.pack()

# 결과 표시
frame_bottom = tk.Frame(root, padx=10, pady=5)
frame_bottom.pack(fill="x")

label_grade = tk.Label(frame_bottom, text="등급: -")
label_grade.pack(anchor="w")

label_result = tk.Label(frame_bottom, text="납부금액: - 원", font=("맑은 고딕", 11, "bold"))
label_result.pack(anchor="w")

# 단축키: Ctrl + S, Enter
root.bind("<Control-s>", calc_fee)
root.bind("<Return>", calc_fee)

root.mainloop()
