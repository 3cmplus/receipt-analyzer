# -*- coding: utf-8 -*-
import os, json, re, mimetypes, threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import customtkinter as ctk
from PIL import Image
from google import genai
from google.genai import types
from config import GEMINI_API_KEY, GEMINI_MODEL
from excel_export import append_row, EXCEL_PATH

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

client = genai.Client(api_key=GEMINI_API_KEY)

PROMPT = """이 영수증 이미지를 꼼꼼히 분석해서 아래 JSON 형식으로만 응답해줘.
JSON 외에 다른 텍스트, 설명, 마크다운 코드블록은 절대 포함하지 마.

{
  "가맹점명": "상호명",
  "주소": "가맹점 주소 (없으면 확인불가)",
  "결제일시": "YYYY-MM-DD HH:MM",
  "총금액": "숫자만 (예: 15000)",
  "항목": [
    {"품목명": "상품명", "수량": "1", "단가": "숫자만"}
  ]
}

주의사항:
- 한글 상호명·주소·품목명은 그대로 유지해줘
- 읽을 수 없는 항목은 "확인불가"로 표시
- 항목이 없으면 빈 배열 [] 반환
- 금액은 쉼표 없이 숫자만 반환 (예: 12000)
"""

_Base = TkinterDnD.Tk if HAS_DND else tk.Tk


class ReceiptApp(_Base):
    def __init__(self):
        super().__init__()
        self.selected_file = None
        self._ctk_img_ref  = None
        self._last_result  = None

        self.title("영수증분석기")
        self.geometry("680x900")
        self.resizable(False, False)
        self.configure(bg="#1c1c1e")

        self._build_ui()

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    # ── UI 구성 ────────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 20}

        # 헤더
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", pady=(22, 0), **pad)
        ctk.CTkLabel(hdr, text="🧾  영수증분석기",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(hdr, text="이미지를 업로드하면 Gemini AI가 항목을 자동 추출합니다",
                     font=ctk.CTkFont(size=12),
                     text_color="#8e8e93").pack(anchor="w", pady=(3, 0))

        # 업로드 존
        self.upload_frame = ctk.CTkFrame(self, corner_radius=12,
                                          fg_color="#2c2c2e",
                                          border_color="#48484a",
                                          border_width=2,
                                          cursor="hand2")
        self.upload_frame.pack(fill="x", pady=14, **pad)
        self.upload_frame.bind("<Button-1>", lambda e: self._browse())

        ctk.CTkLabel(self.upload_frame, text="📂",
                     font=ctk.CTkFont(size=42)).pack(pady=(26, 6))
        lbl = ctk.CTkLabel(self.upload_frame,
                            text="클릭하거나 이미지를 끌어다 놓으세요",
                            font=ctk.CTkFont(size=14, weight="bold"))
        lbl.pack()
        lbl.bind("<Button-1>", lambda e: self._browse())
        ctk.CTkLabel(self.upload_frame, text="JPG · PNG · WEBP 지원",
                     font=ctk.CTkFont(size=11), text_color="#8e8e93").pack(pady=(4, 26))

        # 미리보기 (초기 숨김)
        self.preview_frame = ctk.CTkFrame(self, corner_radius=12, fg_color="#2c2c2e")
        self.preview_img_lbl  = ctk.CTkLabel(self.preview_frame, text="")
        self.preview_img_lbl.pack(padx=14, pady=(14, 6))
        self.preview_name_lbl = ctk.CTkLabel(self.preview_frame, text="",
                                              font=ctk.CTkFont(size=11),
                                              text_color="#8e8e93")
        self.preview_name_lbl.pack(pady=(0, 12))

        # 분석 버튼
        self.analyze_btn = ctk.CTkButton(self, text="분석 시작",
                                          font=ctk.CTkFont(size=15, weight="bold"),
                                          height=46, corner_radius=10,
                                          state="disabled",
                                          command=self._analyze)
        self.analyze_btn.pack(fill="x", pady=(0, 6), **pad)

        # 상태 레이블
        self.status_lbl = ctk.CTkLabel(self, text="",
                                        font=ctk.CTkFont(size=12),
                                        text_color="#8e8e93")
        self.status_lbl.pack(pady=(0, 6))

        # ── 결과 카드 (초기 숨김) ──────────────────────────
        self.result_frame = ctk.CTkFrame(self, corner_radius=12, fg_color="#2c2c2e")

        # 요약 행 1: 결제일시 / 가맹점명 / 총금액
        row1 = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        row1.pack(fill="x", padx=14, pady=(14, 6))
        row1.columnconfigure((0, 1, 2), weight=1)
        self.c_date  = self._card(row1, "결제일시", 0)
        self.c_store = self._card(row1, "가맹점명", 1)
        self.c_total = self._card(row1, "총 금액",  2, hi=True)

        # 요약 행 2: 주소 (전체 너비)
        row2 = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=(0, 8))
        row2.columnconfigure(0, weight=1)
        addr_f = ctk.CTkFrame(row2, fg_color="#3a3a3c", corner_radius=8)
        addr_f.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(addr_f, text="주소", font=ctk.CTkFont(size=10),
                     text_color="#8e8e93").pack(anchor="w", padx=12, pady=(8, 2))
        self.c_addr = ctk.CTkLabel(addr_f, text="-",
                                    font=ctk.CTkFont(size=12),
                                    text_color="#ffffff",
                                    anchor="w", wraplength=580)
        self.c_addr.pack(anchor="w", padx=12, pady=(0, 10))

        ctk.CTkFrame(self.result_frame, height=1,
                     fg_color="#48484a").pack(fill="x", padx=14, pady=4)

        ctk.CTkLabel(self.result_frame, text="항목 상세",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color="#8e8e93").pack(anchor="w", padx=14, pady=(4, 4))

        # Treeview
        tbl_frame = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        tbl_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("R.Treeview",
                         background="#2c2c2e", foreground="#ffffff",
                         fieldbackground="#2c2c2e", rowheight=28,
                         font=("Malgun Gothic", 11))
        style.configure("R.Treeview.Heading",
                         background="#3a3a3c", foreground="#8e8e93",
                         font=("Malgun Gothic", 10, "bold"), relief="flat")
        style.map("R.Treeview", background=[("selected", "#0a84ff")])

        self.tree = ttk.Treeview(tbl_frame,
                                   columns=("품목명", "수량", "단가", "금액"),
                                   show="headings", style="R.Treeview", height=6)
        for col, w, anc in [("품목명", 230, "w"), ("수량", 55, "center"),
                              ("단가", 110, "e"),  ("금액", 110, "e")]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=anc)

        sb = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 엑셀 저장 버튼 (result_frame 안에 위치)
        self.excel_btn = ctk.CTkButton(self.result_frame, text="📥  엑셀로 저장",
                                        font=ctk.CTkFont(size=13, weight="bold"),
                                        height=40, corner_radius=8,
                                        fg_color="#16a34a",
                                        hover_color="#15803d",
                                        command=self._save_excel)
        self.excel_btn.pack(fill="x", padx=14, pady=(4, 6))

        ctk.CTkButton(self.result_frame, text="다른 영수증 분석하기",
                       fg_color="transparent", border_color="#48484a",
                       border_width=1, text_color="#8e8e93",
                       hover_color="#3a3a3c", corner_radius=8, height=34,
                       command=self._reset).pack(fill="x", padx=14, pady=(0, 14))

    def _card(self, parent, label, col, hi=False):
        f = ctk.CTkFrame(parent, fg_color="#3a3a3c", corner_radius=8)
        f.grid(row=0, column=col, padx=4, sticky="ew")
        ctk.CTkLabel(f, text=label, font=ctk.CTkFont(size=10),
                     text_color="#8e8e93").pack(pady=(10, 2))
        v = ctk.CTkLabel(f, text="-",
                          font=ctk.CTkFont(size=13, weight="bold"),
                          text_color="#0a84ff" if hi else "#ffffff",
                          wraplength=160)
        v.pack(pady=(0, 10), padx=6)
        return v

    # ── 파일 선택 ──────────────────────────────────────────────
    def _browse(self):
        path = filedialog.askopenfilename(
            title="영수증 이미지 선택",
            filetypes=[("이미지", "*.jpg *.jpeg *.png *.webp *.bmp"),
                       ("모든 파일", "*.*")])
        if path:
            self._load(path)

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path:
            self._load(path)

    def _load(self, path):
        self.selected_file = path
        self._last_result  = None

        img = Image.open(path)
        img.thumbnail((420, 240))
        ctk_img = ctk.CTkImage(light_image=img, dark_image=img,
                                 size=(img.width, img.height))
        self._ctk_img_ref = ctk_img
        self.preview_img_lbl.configure(image=ctk_img)
        self.preview_name_lbl.configure(text=os.path.basename(path))

        self.upload_frame.pack_forget()
        self.preview_frame.pack(fill="x", padx=20, pady=8,
                                  before=self.analyze_btn)
        self.result_frame.pack_forget()
        self.analyze_btn.configure(state="normal", text="분석 시작")
        self.status_lbl.configure(text="")

    # ── 분석 ───────────────────────────────────────────────────
    def _analyze(self):
        self.analyze_btn.configure(state="disabled", text="분석 중...")
        self.status_lbl.configure(text="Gemini AI가 분석 중입니다...",
                                   text_color="#0a84ff")
        threading.Thread(target=self._call_api, daemon=True).start()

    def _call_api(self):
        try:
            with open(self.selected_file, "rb") as f:
                data = f.read()
            mime = mimetypes.guess_type(self.selected_file)[0] or "image/jpeg"

            res = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=[types.Part.from_bytes(data=data, mime_type=mime), PROMPT])

            text = re.sub(r"^```(?:json)?\s*", "", res.text.strip())
            text = re.sub(r"\s*```$", "", text)
            m = re.search(r"\{.*\}", text, re.DOTALL)
            result = json.loads(m.group() if m else text)
            self.after(0, self._show, result)
        except Exception as e:
            self.after(0, self._err, str(e))

    def _show(self, d):
        self._last_result = d
        self.c_date.configure(text=d.get("결제일시", "-"))
        self.c_store.configure(text=d.get("가맹점명", "-"))
        self.c_addr.configure(text=d.get("주소", "-"))

        total = d.get("총금액", "-")
        try:
            total = f"{int(total):,}원"
        except Exception:
            total = f"{total}원" if total not in ("-", "") else "-"
        self.c_total.configure(text=total)

        for row in self.tree.get_children():
            self.tree.delete(row)
        for item in d.get("항목", []):
            name  = item.get("품목명", "-")
            qty   = item.get("수량",   "-")
            price = item.get("단가",   "-")
            try:
                amt = f"{int(qty) * int(price):,}원"
            except Exception:
                amt = "-"
            try:
                price = f"{int(price):,}원"
            except Exception:
                pass
            self.tree.insert("", "end", values=(name, qty, price, amt))

        self.result_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.excel_btn.configure(text="📥  엑셀로 저장",
                                  fg_color="#16a34a", state="normal")
        self.analyze_btn.configure(state="normal", text="분석 시작")
        self.status_lbl.configure(text="✅  분석 완료!", text_color="#30d158")

    def _err(self, msg):
        self.analyze_btn.configure(state="normal", text="분석 시작")
        self.status_lbl.configure(text=f"오류: {msg}", text_color="#ff453a")

    # ── 엑셀 저장 ──────────────────────────────────────────────
    def _save_excel(self):
        if not self._last_result:
            return
        try:
            path = append_row(self._last_result)
            self.excel_btn.configure(text="✅  저장 완료",
                                      fg_color="#0f766e", state="disabled")
            self.status_lbl.configure(
                text=f"저장: {os.path.basename(path)}", text_color="#30d158")
        except Exception as e:
            messagebox.showerror("저장 오류", str(e))

    # ── 초기화 ────────────────────────────────────────────────
    def _reset(self):
        self.selected_file = None
        self._last_result  = None
        self.preview_frame.pack_forget()
        self.result_frame.pack_forget()
        self.upload_frame.pack(fill="x", padx=20, pady=14,
                                 before=self.analyze_btn)
        self.analyze_btn.configure(state="disabled")
        self.status_lbl.configure(text="")


if __name__ == "__main__":
    app = ReceiptApp()
    app.mainloop()
