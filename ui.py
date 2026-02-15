# ui.py
from __future__ import annotations

import os
import json
import sys
import tkinter as tk
import traceback
from tkinter import ttk, filedialog, messagebox, simpledialog

# Try to import tkinterdnd2 for drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False

from __version__ import __version__
from diagnostics import collect_summary, format_summary
from excel_io import read_header_file, get_sheet_names
from open_excel import (
    list_open_books,
    list_sheets,
    read_header_open,
    xlwings_available,
)
from matcher import match_universal
from commercial_config import (
    ADMIN_PASSWORD,
    CREATOR_NAME,
    SALES_INFO,
    BANK_INFO,
    PRICE_INFO,
    CONTACT_INFO,
    TRIAL_EMAIL,
    TRIAL_SUBJECT,
)
from admin_panel import AdminPanel

APP_TITLE = "Easy Match(이지 매치)"
APP_DESCRIPTION = "엑셀과 CSV를 하나로, 클릭 한 번으로 끝나는 데이터 매칭"
OUT_DIR = os.path.join(os.getcwd(), "outputs")
from config import PRESET_FILE, REPLACE_FILE, get_system_font
os.makedirs(OUT_DIR, exist_ok=True)


# -------------------------
# ToolTip Class
# -------------------------
class ToolTip:
    """
    Create a tooltip for a given widget with improved styling and positioning.
    """
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.id_after = None
        
        widget.bind("<Enter>", self.on_enter)
        widget.bind("<Leave>", self.on_leave)
        widget.bind("<Button>", self.on_leave)  # Hide on click
    
    def on_enter(self, event=None):
        # Schedule tooltip display after delay
        self.id_after = self.widget.after(self.delay, self.show_tooltip)
    
    def on_leave(self, event=None):
        # Cancel scheduled tooltip
        if self.id_after:
            self.widget.after_cancel(self.id_after)
            self.id_after = None
        # Hide tooltip if visible
        self.hide_tooltip()
    
    def show_tooltip(self):
        if self.tooltip_window:
            return
        
        # Get widget position
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        # Create tooltip window
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        # Create tooltip label with modern styling
        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            background="#2c3e50",
            foreground="#ecf0f1",
            relief="solid",
            borderwidth=1,
            font=(get_system_font()[0], 10),
            padx=8,
            pady=5,
            justify="left"
        )
        label.pack()
    
    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


# -------------------------
# Replace rules persistence
# -------------------------
def _load_replace_file():
    try:
        if os.path.exists(REPLACE_FILE):
            with open(REPLACE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            # 구버전 호환: {"presetName": {...}} 형태였던 경우
            if isinstance(data, dict) and "presets" not in data:
                return {"presets": data, "active": {}, "active_name": ""}

            return {
                "presets": data.get("presets", {}) or {},
                "active": data.get("active", {}) or {},
                "active_name": data.get("active_name", "") or "",
            }
    except Exception:
        pass

    return {"presets": {}, "active": {}, "active_name": ""}


def _save_replace_file(presets: dict, active: dict, active_name: str = ""):
    with open(REPLACE_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {
                "presets": presets or {},
                "active": active or {},
                "active_name": active_name or "",
            },
            f,
            ensure_ascii=False,
            indent=4,
        )


class ReplacementEditor(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("데이터 치환 규칙 설정")
        self.geometry("650x520")

        data = _load_replace_file()
        self.presets = data["presets"]
        self.rules = data["active"] or {}
        self.active_name = data.get("active_name", "") or ""

        top_frame = ttk.LabelFrame(self, text="규칙 프리셋 관리", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        self.preset_var = tk.StringVar()
        self.cb_preset = ttk.Combobox(
            top_frame, textvariable=self.preset_var, state="readonly", width=25
        )
        self.cb_preset.pack(side="left", padx=5)
        self.cb_preset.bind("<<ComboboxSelected>>", self.on_preset_selected)

        ttk.Button(top_frame, text="저장 (Save)", command=self.save_preset).pack(
            side="left", padx=2
        )
        ttk.Button(top_frame, text="삭제 (Del)", command=self.delete_preset).pack(
            side="left", padx=2
        )
        ttk.Button(top_frame, text="초기화 (Reset)", command=self.clear_all).pack(
            side="right", padx=2
        )

        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)

        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(form_frame, text="대상 컬럼명:").grid(
            row=0, column=0, padx=5, sticky="w"
        )
        self.ent_col = ttk.Entry(form_frame, width=20)
        self.ent_col.grid(row=0, column=1, padx=5)

        ttk.Label(form_frame, text="변경 전 (Old):").grid(
            row=0, column=2, padx=5, sticky="w"
        )
        self.ent_old = ttk.Entry(form_frame, width=15)
        self.ent_old.grid(row=0, column=3, padx=5)

        ttk.Label(form_frame, text="변경 후 (New):").grid(
            row=0, column=4, padx=5, sticky="w"
        )
        self.ent_new = ttk.Entry(form_frame, width=15)
        self.ent_new.grid(row=0, column=5, padx=5)

        ttk.Button(form_frame, text="추가/수정", command=self.add_rule).grid(
            row=0, column=6, padx=10
        )

        columns = ("col", "old", "new")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=15)
        self.tree.heading("col", text="대상 컬럼")
        self.tree.heading("old", text="변경 전 값 (Old)")
        self.tree.heading("new", text="변경 후 값 (New)")
        self.tree.column("col", width=180)
        self.tree.column("old", width=180)
        self.tree.column("new", width=180)

        scroll = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.on_double_click)

        bot_frame = ttk.Frame(self, padding=10)
        bot_frame.pack(fill="x")
        ttk.Label(
            bot_frame, text="* 목록을 더블클릭하면 삭제됩니다.", foreground="gray"
        ).pack(side="left")
        ttk.Button(bot_frame, text="닫기 (적용)", command=self._close).pack(side="right")

        self.update_preset_cb()
        self.refresh_tree()

    def _persist(self):
        name = self.cb_preset.get() if self.cb_preset.get() in self.presets else self.active_name
        _save_replace_file(self.presets, self.rules, name)

    def _close(self):
        self._persist()
        self.destroy()

    def add_rule(self):
        col = self.ent_col.get().strip()
        old = self.ent_old.get().strip()
        new = self.ent_new.get().strip()
        if not col or not old:
            messagebox.showwarning("입력 오류", "컬럼명과 변경 전 값은 필수입니다.")
            return
        self.rules.setdefault(col, {})[old] = new
        self.refresh_tree()
        self.ent_old.delete(0, tk.END)
        self.ent_new.delete(0, tk.END)
        self.ent_old.focus()
        self._persist()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for col, mapping in (self.rules or {}).items():
            if not isinstance(mapping, dict):
                continue
            for old, new in mapping.items():
                self.tree.insert("", "end", values=(col, old, new))

    def on_double_click(self, event=None):
        item = self.tree.selection()
        if not item:
            return
        col, old, new = self.tree.item(item, "values")
        if messagebox.askyesno("삭제", f"'{col}' 컬럼의 '{old}' -> '{new}' 규칙을 삭제할까요?"):
            if col in self.rules and old in self.rules[col]:
                del self.rules[col][old]
                if not self.rules[col]:
                    del self.rules[col]
            self.refresh_tree()
            self._persist()

    def clear_all(self):
        if messagebox.askyesno("초기화", "모든 규칙을 지우시겠습니까?"):
            self.rules = {}
            self.refresh_tree()
            self._persist()

    def update_preset_cb(self):
        names = list((self.presets or {}).keys())
        self.cb_preset["values"] = names
        self.cb_preset.set("설정을 선택하세요" if names else "저장된 설정 없음")

    def save_preset(self):
        if not self.rules:
            messagebox.showwarning("경고", "저장할 규칙이 없습니다.")
            return
        name = simpledialog.askstring("설정 저장", "치환 규칙 세트 이름 입력:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        import copy

        self.presets[name] = copy.deepcopy(self.rules)
        self.active_name = name
        self.cb_preset.set(name)
        self.update_preset_cb()
        self._persist()
        messagebox.showinfo("저장", f"[{name}] 규칙이 저장되었습니다.")

    def delete_preset(self):
        name = self.cb_preset.get()
        if name in self.presets and messagebox.askyesno(
            "삭제", f"[{name}] 규칙 세트를 삭제하시겠습니까?"
        ):
            del self.presets[name]
            if self.active_name == name:
                self.active_name = ""
            self.update_preset_cb()
            self._persist()

    def on_preset_selected(self, event=None):
        name = self.cb_preset.get()
        if name in self.presets:
            import copy

            self.rules = copy.deepcopy(self.presets[name])
            self.active_name = name
            self.refresh_tree()
            self._persist()

    def get_rules(self):
        return self.rules


class MultiSelectListBox(ttk.Frame):
    def __init__(self, master, height=5):
        super().__init__(master)
        self.listbox = tk.Listbox(
            self, selectmode="multiple", height=height, exportselection=False
        )
        sb = ttk.Scrollbar(self, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ttk.Label(
            self,
            text="(Ctrl/Shift+클릭으로 다중 선택)",
            font=(get_system_font()[0], 10),
            foreground="gray",
        ).pack(side="bottom", anchor="w")

    def set_items(self, items):
        self.listbox.delete(0, tk.END)
        for it in items or []:
            self.listbox.insert(tk.END, it)

    def get_selected(self):
        return [self.listbox.get(i) for i in self.listbox.curselection()]

    def select_item(self, item):
        self.listbox.selection_clear(0, tk.END)
        try:
            idx = self.listbox.get(0, tk.END).index(item)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        except ValueError:
            pass


class SourceFrame(ttk.LabelFrame):
    def __init__(self, master, title, mode_var, on_change=None, is_base=False):
        super().__init__(master, text=title, padding=10)
        self.mode = mode_var
        self.on_change = on_change
        self.is_base = is_base

        self.path = tk.StringVar()
        self.book = tk.StringVar()
        self.sheet = tk.StringVar()
        self.header = tk.IntVar(value=1)

        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 5))
        
        file_radio = ttk.Radiobutton(
            top, text="파일 불러오기", value="file", variable=self.mode, command=self._refresh_ui
        )
        file_radio.pack(side="left", padx=(0, 10))
        ToolTip(file_radio, "Excel 또는 CSV 파일을 선택합니다")
        
        open_radio = ttk.Radiobutton(
            top, text="열려있는 엑셀", value="open", variable=self.mode, command=self._refresh_ui
        )
        open_radio.pack(side="left")
        ToolTip(open_radio, "현재 Excel에서 열려있는 파일을 사용합니다\n(xlwings 필요)")
        
        refresh_btn = ttk.Button(top, text="새로고침 (Refresh)", command=self.refresh_open)
        refresh_btn.pack(side="right")
        ToolTip(refresh_btn, "열려있는 Excel 파일 목록을 새로고침합니다")

        self.f_frame = ttk.Frame(self)
        path_entry = ttk.Entry(self.f_frame, textvariable=self.path)
        path_entry.pack(side="left", fill="x", expand=True)
        ToolTip(path_entry, "선택한 파일 경로가 표시됩니다\n파일을 여기에 드래그하여 선택할 수도 있습니다")
        
        # Add drag and drop support if available
        if DRAG_DROP_AVAILABLE:
            try:
                path_entry.drop_target_register(DND_FILES)
                path_entry.dnd_bind('<<Drop>>', self._on_drop)
                
                # Visual feedback for drag over
                def on_drag_enter(event):
                    path_entry.config(background="#e8f5e9")  # Light green
                    return event.action
                
                def on_drag_leave(event):
                    path_entry.config(background="white")
                    return event.action
                
                path_entry.dnd_bind('<<DragEnter>>', on_drag_enter)
                path_entry.dnd_bind('<<DragLeave>>', on_drag_leave)
            except Exception:
                pass  # Silently fail if drag and drop setup fails
        
        browse_btn = ttk.Button(self.f_frame, text="찾기 (Browse)", command=self._pick_file)
        browse_btn.pack(side="left", padx=5)
        ToolTip(browse_btn, "Excel 또는 CSV 파일을 선택합니다\n(xlsx, xls, csv 지원)")

        self.o_frame = ttk.Frame(self)
        self.cb_book = ttk.Combobox(self.o_frame, textvariable=self.book, state="readonly")
        self.cb_book.pack(side="left", fill="x", expand=True)
        self.cb_book.bind("<<ComboboxSelected>>", self._on_book_select)

        self.c_frame = ttk.Frame(self)
        self.c_frame.pack(fill="x", pady=(5, 5))
        ttk.Label(self.c_frame, text="시트:").pack(side="left")
        self.cb_sheet = ttk.Combobox(self.c_frame, textvariable=self.sheet, state="readonly", width=18)
        self.cb_sheet.pack(side="left", padx=5)
        self.cb_sheet.bind("<<ComboboxSelected>>", self._notify_change)
        ToolTip(self.cb_sheet, "데이터가 있는 시트를 선택합니다")

        ttk.Label(self.c_frame, text="헤더:").pack(side="left", padx=(10, 0))
        header_spin = ttk.Spinbox(
            self.c_frame, from_=1, to=99, textvariable=self.header, width=5, command=self._notify_change
        )
        header_spin.pack(side="left", padx=5)
        ToolTip(header_spin, "컬럼명이 있는 행 번호를 지정합니다\n(보통 1행)")

        if self.is_base:
            key_frame = ttk.LabelFrame(self, text="매칭 키 (Key) 선택", padding=5)
            key_frame.pack(fill="x", pady=(5, 0))
            self.key_listbox = MultiSelectListBox(key_frame, height=4)
            self.key_listbox.pack(fill="x")

        self._refresh_ui()
    
    def _on_drop(self, event):
        """Handle file drop event"""
        try:
            # Get dropped file path (remove curly braces if present)
            file_path = event.data.strip('{}')
            
            # Validate file extension
            valid_extensions = ('.xlsx', '.xls', '.csv')
            if not file_path.lower().endswith(valid_extensions):
                messagebox.showwarning(
                    "파일 형식 오류",
                    f"지원하지 않는 파일 형식입니다.\n\n지원 형식: {', '.join(valid_extensions)}"
                )
                return
            
            # Set the path
            self.path.set(file_path)
            
            # Load sheets
            try:
                self.cb_sheet["values"] = get_sheet_names(file_path)
                if self.cb_sheet["values"]:
                    self.cb_sheet.current(0)
            except Exception:
                pass
            
            # Reset background color
            event.widget.config(background="white")
            
            # Notify change
            self._notify_change()
            
        except Exception as e:
            messagebox.showerror("오류", f"파일을 불러올 수 없습니다:\n{str(e)}")

    def _refresh_ui(self):
        if self.mode.get() == "file":
            self.o_frame.pack_forget()
            self.f_frame.pack(fill="x", before=self.c_frame)
        else:
            self.f_frame.pack_forget()
            self.o_frame.pack(fill="x", before=self.c_frame)
            self.refresh_open()
        self._notify_change()

    def _pick_file(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls *.csv"), ("All", "*.*")]
        )
        if not p:
            return
        self.path.set(p)
        try:
            self.cb_sheet["values"] = get_sheet_names(p)
            if self.cb_sheet["values"]:
                self.cb_sheet.current(0)
        except Exception:
            pass
        self._notify_change()

    def refresh_open(self):
        if self.mode.get() != "open":
            return

        if not xlwings_available():
            messagebox.showwarning(
                "열려있는 엑셀 모드 불가",
                "xlwings/Excel 연동이 불가능합니다.\n\n"
                "- Excel 설치 확인\n"
                "- xlwings 설치(pip)\n"
                "- macOS: Automation 권한 허용\n\n"
                "파일 모드로 사용하거나 설치 후 다시 시도하세요.",
            )
            self.cb_book["values"] = []
            self.book.set("")
            self.cb_sheet["values"] = []
            self.sheet.set("")
            return

        books = list_open_books()
        self.cb_book["values"] = books
        if books:
            if not self.book.get() or self.book.get() not in books:
                self.book.set(books[0])
        else:
            self.book.set("")
            self.cb_sheet["values"] = []
            self.sheet.set("")
            return

        self._on_book_select()

    def _on_book_select(self, event=None):
        if not self.book.get():
            return
        try:
            sheets = list_sheets(self.book.get())
            self.cb_sheet["values"] = sheets
            if sheets:
                self.cb_sheet.current(0)
        except Exception as e:
            messagebox.showerror("시트 조회 실패", str(e))
        self._notify_change()

    def _notify_change(self, event=None):
        if self.on_change:
            self.on_change()

    def get_config(self):
        return {
            "type": self.mode.get(),
            "path": self.path.get(),
            "book": self.book.get(),
            "sheet": self.sheet.get(),
            "header": int(self.header.get()),
        }

    def get_selected_keys(self):
        if self.is_base and hasattr(self, "key_listbox"):
            return self.key_listbox.get_selected()
        return []


class GridCheckList(ttk.Frame):
    def __init__(self, master, columns=4, height=300):
        super().__init__(master)
        self.columns = columns
        self.vars: dict[str, tk.BooleanVar] = {}

        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 4))
        self.q = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.q)
        ent.pack(side="left", fill="x", expand=True)
        ent.bind("<KeyRelease>", lambda e: self._filter())
        ttk.Button(top, text="지우기", width=6, command=self._clear).pack(side="left", padx=4)

        self.canvas = tk.Canvas(self, height=height, bg="white", highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.inner_id, width=e.width))

        self.all_items: list[str] = []

    def set_items(self, items):
        self.all_items = list(items or [])
        self.vars.clear()
        for w in self.inner.winfo_children():
            w.destroy()
        if not self.all_items:
            ttk.Label(self.inner, text="(데이터 없음)").pack()
            return
        for it in self.all_items:
            self.vars[it] = tk.BooleanVar(value=False)
        self._render(self.all_items)

    def _render(self, items):
        for w in self.inner.winfo_children():
            w.destroy()
        for idx, it in enumerate(items):
            cb = ttk.Checkbutton(self.inner, text=it, variable=self.vars[it])
            cb.grid(row=idx // self.columns, column=idx % self.columns, sticky="w", padx=4, pady=2)
        for i in range(self.columns):
            self.inner.columnconfigure(i, weight=1)

    def checked(self):
        return [k for k, v in self.vars.items() if v.get()]

    def check_all(self):
        for v in self.vars.values():
            v.set(True)

    def uncheck_all(self):
        for v in self.vars.values():
            v.set(False)

    def set_checked_items(self, items):
        self.uncheck_all()
        cnt = 0
        for it in items or []:
            if it in self.vars:
                self.vars[it].set(True)
                cnt += 1
        return cnt

    def _clear(self):
        self.q.set("")
        self._render(self.all_items)

    def _filter(self):
        q = (self.q.get() or "").strip().lower()
        self._render([it for it in self.all_items if q in it.lower()] if q else self.all_items)


class App(tk.Tk):
    def __init__(self, license_info=None):
        super().__init__()
        self.license_info = license_info or {"type": "personal", "expiry": "-"}
        
        # Restore missing initializations
        default_base = "file" if sys.platform == "darwin" else "open"
        default_tgt = "file"
        self.base_mode = tk.StringVar(value=default_base)
        self.tgt_mode = tk.StringVar(value=default_tgt)

        self.presets: dict[str, list[str]] = {}
        self.opt_fuzzy = tk.BooleanVar(value=False)
        self.opt_color = tk.BooleanVar(value=True)
        self.replacer_win = None

        self.title(APP_TITLE)
        self.geometry("1050x850")  # Further reduced to ensure bottom content is visible


        # --- HEADER SECTION ---
        header = tk.Frame(self, bg="#2c3e50", height=110)
        header.pack(fill="x", side="top")
        
        # Left side: Logo + Title + Slogan
        left_header = tk.Frame(header, bg="#2c3e50")
        left_header.pack(side="left", padx=20, pady=15)
        
        # Logo icon (text-based to avoid emoji crash on macOS)
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "assets", "logo.png")
            if os.path.exists(logo_path):
                from PIL import Image, ImageTk
                logo_img = Image.open(logo_path).resize((48, 48), Image.Resampling.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = tk.Label(left_header, image=logo_photo, bg="#2c3e50", cursor="hand2")
                logo_label.image = logo_photo  # Keep reference
                logo_label.pack(side="left", padx=(0, 15))
                
                # Add subtle hover effect to logo
                def on_logo_enter(e):
                    logo_label.config(relief="raised", borderwidth=2)
                def on_logo_leave(e):
                    logo_label.config(relief="flat", borderwidth=0)
                logo_label.bind("<Enter>", on_logo_enter)
                logo_label.bind("<Leave>", on_logo_leave)
                logo_label.bind("<Button-1>", lambda e: show_feature_info())
        except:
            # Fallback: Use text icon (NO EMOJI - causes crash on macOS)
            logo_label = tk.Label(left_header, text="EM", font=(get_system_font()[0], 24, "bold"), 
                                 bg="#27ae60", fg="white", padx=8, pady=4, relief="raised", borderwidth=2)
            logo_label.pack(side="left", padx=(0, 15))
        
        # Title and slogan container
        title_container = tk.Frame(left_header, bg="#2c3e50")
        title_container.pack(side="left")
        
        lbl_title = tk.Label(title_container, text=APP_TITLE, font=(get_system_font()[0], 18, "bold"), bg="#2c3e50", fg="#ecf0f1")
        lbl_title.pack(anchor="w")
        
        lbl_desc = tk.Label(title_container, text=APP_DESCRIPTION, font=(get_system_font()[0], 11), bg="#2c3e50", fg="#95a5a6", justify="left")
        lbl_desc.pack(anchor="w", pady=(3, 0))

        # Feature Info Popup
        def show_feature_info():
            top = tk.Toplevel(self)
            top.title("Easy Match란?")
            top.geometry("600x500")
            top.configure(bg="white")
            
            # Center the popup
            root_x = self.winfo_rootx()
            root_y = self.winfo_rooty()
            root_w = self.winfo_width()
            root_h = self.winfo_height()
            x = root_x + (root_w // 2) - 300
            y = root_y + (root_h // 2) - 250
            top.geometry(f"600x500+{x}+{y}")

            # Title
            tk.Label(top, text="Easy Match란?", font=(get_system_font()[0], 20, "bold"), bg="white", fg="#2c3e50").pack(pady=(25, 20))
            
            # Content frame
            content = tk.Frame(top, bg="white")
            content.pack(padx=30, fill="both", expand=True)
            
            # Easy section
            easy_frame = tk.Frame(content, bg="#e8f5e9", relief="solid", borderwidth=2, highlightbackground="#27ae60", highlightthickness=2)
            easy_frame.pack(fill="x", pady=(0, 15))
            
            tk.Label(easy_frame, text="[Easy] 이지", font=(get_system_font()[0], 15, "bold"), bg="#e8f5e9", fg="#27ae60").pack(pady=(15, 8), anchor="w", padx=20)
            tk.Label(easy_frame, text="복잡한 엑셀 수식이나 코딩 없이, 누구나 쉽고 간편하게 사용할 수 있습니다.", 
                    font=(get_system_font()[0], 11), bg="#e8f5e9", fg="#2c3e50", anchor="w", wraplength=520, justify="left").pack(padx=20, pady=(0, 5), fill="x")
            tk.Label(easy_frame, text="• 숫자/문자 서식 자동 보정 기능 포함", 
                    font=(get_system_font()[0], 10), bg="#e8f5e9", fg="#34495e", anchor="w").pack(padx=40, pady=2, fill="x")
            tk.Label(easy_frame, text="• 직관적인 UI로 클릭 몇 번이면 완료", 
                    font=(get_system_font()[0], 10), bg="#e8f5e9", fg="#34495e", anchor="w").pack(padx=40, pady=(0, 15), fill="x")
            
            # Match section
            match_frame = tk.Frame(content, bg="#e3f2fd", relief="solid", borderwidth=2, highlightbackground="#2B579A", highlightthickness=2)
            match_frame.pack(fill="x", pady=(0, 15))
            
            tk.Label(match_frame, text="[Match] 매치", font=(get_system_font()[0], 15, "bold"), bg="#e3f2fd", fg="#2B579A").pack(pady=(15, 8), anchor="w", padx=20)
            tk.Label(match_frame, text="엑셀(Excel)뿐만 아니라 CSV 파일까지 복잡한 설정 없이 쉽게 매칭해줍니다.", 
                    font=(get_system_font()[0], 11), bg="#e3f2fd", fg="#2c3e50", anchor="w", wraplength=520, justify="left").pack(padx=20, pady=(0, 5), fill="x")
            tk.Label(match_frame, text="• 다양한 파일 형식 지원 (xlsx, xls, csv)", 
                    font=(get_system_font()[0], 10), bg="#e3f2fd", fg="#34495e", anchor="w").pack(padx=40, pady=2, fill="x")
            tk.Label(match_frame, text="• 자동 데이터 타입 감지", 
                    font=(get_system_font()[0], 10), bg="#e3f2fd", fg="#34495e", anchor="w").pack(padx=40, pady=(0, 15), fill="x")
            
            # Features section
            features_frame = tk.Frame(content, bg="#fff3e0", relief="solid", borderwidth=2, highlightbackground="#f39c12", highlightthickness=2)
            features_frame.pack(fill="x", pady=(0, 15))
            
            tk.Label(features_frame, text="[주요 기능]", font=(get_system_font()[0], 15, "bold"), bg="#fff3e0", fg="#e67e22").pack(pady=(15, 8), anchor="w", padx=20)
            tk.Label(features_frame, text="• 자주 쓰는 컬럼 저장 기능", 
                    font=(get_system_font()[0], 10), bg="#fff3e0", fg="#34495e", anchor="w").pack(padx=40, pady=2, fill="x")
            tk.Label(features_frame, text="• 오타 자동 보정 (Fuzzy Match)", 
                    font=(get_system_font()[0], 10), bg="#fff3e0", fg="#34495e", anchor="w").pack(padx=40, pady=2, fill="x")
            tk.Label(features_frame, text="• 치환 설정으로 데이터 전처리", 
                    font=(get_system_font()[0], 10), bg="#fff3e0", fg="#34495e", anchor="w").pack(padx=40, pady=2, fill="x")
            tk.Label(features_frame, text="• 색상 강조로 매칭 결과 확인", 
                    font=(get_system_font()[0], 10), bg="#fff3e0", fg="#34495e", anchor="w").pack(padx=40, pady=(0, 15), fill="x")
            
            # Close button
            tk.Button(top, text="닫기", command=top.destroy, bg="#95a5a6", fg="white", 
                     font=(get_system_font()[0], 11, "bold"), padx=25, pady=8, cursor="hand2").pack(pady=(0, 20))

        # Right side: Buttons
        right_header = tk.Frame(header, bg="#2c3e50")
        right_header.pack(side="right", padx=20, pady=15)
        
        # Feature Info Button
        btn_feature_info = tk.Label(
            right_header,
            text="[i] 기능 자세히 보기",
            font=(get_system_font()[0], 10, "bold"),
            bg="#3498db",
            fg="white",
            padx=12,
            pady=6,
            cursor="hand2",
            relief="raised"
        )
        btn_feature_info.pack(side="left", padx=(0, 10))
        btn_feature_info.bind("<Button-1>", lambda e: show_feature_info())
        
        # Smooth hover effect with animation
        def on_enter_info(e):
            btn_feature_info.config(bg="#2980b9", relief="sunken")
        def on_leave_info(e):
            btn_feature_info.config(bg="#3498db", relief="raised")
        btn_feature_info.bind("<Enter>", on_enter_info)
        btn_feature_info.bind("<Leave>", on_leave_info)
        
        # User Guide Button
        btn_guide = tk.Label(
            right_header,
            text="사용방법 (Guide)",
            font=(get_system_font()[0], 10, "bold"),
            bg="#16a085",
            fg="white",
            padx=12,
            pady=6,
            cursor="hand2",
            relief="raised"
        )
        btn_guide.pack(side="left", padx=(0, 10))
        
        # Guide button click handler
        def open_usage_guide_click(e=None):
            import webbrowser
            import os
            try:
                guide_path = os.path.abspath("usage_guide.html")
                webbrowser.open(f"file://{guide_path}")
            except Exception as ex:
                messagebox.showinfo("안내", "사용 가이드 파일을 찾을 수 없습니다.")
        
        btn_guide.bind("<Button-1>", open_usage_guide_click)
        
        # Smooth hover effect for guide button
        def on_enter_guide(e):
            btn_guide.config(bg="#138d75", relief="sunken")
        def on_leave_guide(e):
            btn_guide.config(bg="#16a085", relief="raised")
        btn_guide.bind("<Enter>", on_enter_guide)
        btn_guide.bind("<Leave>", on_leave_guide)
        
        # License Button
        btn_license = tk.Label(
            right_header,
            text="라이선스/관리",
            font=(get_system_font()[0], 10, "bold"),
            bg="#7f8c8d",
            fg="white",
            padx=12,
            pady=6,
            cursor="hand2",
            relief="raised"
        )
        btn_license.pack(side="left")

        def open_license_menu(e=None):
            # Create a popup menu
            menu = tk.Menu(self, tearoff=0)
            
            def open_admin():
                pw = simpledialog.askstring("관리자", "관리자 비밀번호:", show="*")
                if pw != ADMIN_PASSWORD:
                    messagebox.showerror("오류", "비밀번호가 올바르지 않습니다.")
                    return
                AdminPanel(self)

            menu.add_command(label="라이선스 등록 (제품 키 입력)", command=self.register_license)
            menu.add_command(label="관리자 패널", command=open_admin)
            menu.add_separator()
            menu.add_command(
                label=f"라이선스: {self.license_info.get('type','?')} / 만료: {self.license_info.get('expiry','?')}",
                state="disabled",
            )
            
            try:
                # Show menu at button position
                x = btn_license.winfo_rootx()
                y = btn_license.winfo_rooty() + btn_license.winfo_height()
                menu.tk_popup(x, y)
            finally:
                menu.grab_release()
        
        btn_license.bind("<Button-1>", open_license_menu)
        
        # Smooth hover effect for license button
        def on_enter_license(e):
            btn_license.config(bg="#5d6d6e", relief="sunken")
        def on_leave_license(e):
            btn_license.config(bg="#7f8c8d", relief="raised")
        btn_license.bind("<Enter>", on_enter_license)
        btn_license.bind("<Leave>", on_leave_license)


        self._init_ui()
        
        # Apply system font
        default_font = get_system_font()
        self.option_add("*Font", default_font)
        
        self.load_presets()

        # Diagnostics will be shown on demand via toggle button

    def _init_ui(self):
        try:
            style = ttk.Style(self)
            style.theme_use("clam")
            # Modern Button Style
            style.configure("TButton", justify="center", anchor="center", font=(get_system_font()[0], 11), padding=5)
            style.configure("Header.TLabel", background="#2c3e50", foreground="white", font=(get_system_font()[0], 13, "bold"))
        except Exception:
            pass

        main = ttk.Frame(self, padding=10)  # Reduced from 20 to 10
        
        footer = ttk.Frame(main)
        footer.pack(side="bottom", fill="x", pady=(10, 0))

        # Collapsible Diagnostics Section
        diag_container = ttk.Frame(footer)
        diag_container.pack(side="bottom", fill="x", pady=(0, 5))
        
        # Toggle button for diagnostics
        self.diag_expanded = tk.BooleanVar(value=False)
        
        def toggle_diagnostics():
            if self.diag_expanded.get():
                diag_frame.pack_forget()
                diag_toggle_btn.config(text="▶ 환경 진단 (펼치기)")
                self.diag_expanded.set(False)
            else:
                # Show diagnostics
                diag_frame.pack(side="bottom", fill="x", pady=(5, 0))
                diag_toggle_btn.config(text="▼ 환경 진단 (접기)")
                self.diag_expanded.set(True)
                
                # Load diagnostics if not already loaded
                if self.diag_txt.get("1.0", "end-1c") == "":
                    self.diag_txt.config(state="normal")
                    self.diag_txt.insert("1.0", "환경 진단:\n")
                    self.diag_txt.insert("end", format_summary(collect_summary()))
                    self.diag_txt.config(state="disabled")
        
        diag_toggle_btn = ttk.Button(diag_container, text="▶ 환경 진단 (펼치기)", command=toggle_diagnostics)
        diag_toggle_btn.pack(side="top", fill="x")
        ToolTip(diag_toggle_btn, "시스템 환경 정보를 표시/숨김합니다\n(Python 버전, xlwings 설치 여부 등)")
        
        # Diagnostics frame (initially hidden)
        diag_frame = ttk.Frame(diag_container)
        # Don't pack initially - will be shown on toggle
        
        self.diag_txt = tk.Text(diag_frame, height=2, state="disabled", bg="#f9f9f9", fg="#555555")  # Reduced to 2
        self.diag_txt.pack(side="bottom", fill="x")

        # Log text area (always visible) - reduced height
        self.log_txt = tk.Text(footer, height=2, state="disabled", bg="#f0f0f0")  # Reduced to 2
        self.log_txt.pack(side="bottom", fill="x")


        # Run button with simple flat blue styling
        run_btn = tk.Button(
            footer, 
            text="매칭 실행 (RUN)", 
            command=self.run,
            bg="#1e3a8a",  # Deep blue background
            fg="#ffffff",  # Bright white text
            font=(get_system_font()[0], 14, "bold"),  # Increased font size for better readability
            padx=20,
            pady=15,
            cursor="hand2",
            relief="flat",
            borderwidth=0,
            activebackground="#1e40af",  # Slightly lighter blue when clicked
            activeforeground="#ffffff"
        )
        run_btn.pack(side="bottom", fill="x", pady=10)
        ToolTip(run_btn, "설정한 조건으로 데이터 매칭을 시작합니다\n(단축키: Ctrl+M)")
        
        # Hover effect for run button
        def on_run_enter(e):
            run_btn.config(bg="#d35400", relief="sunken")
        
        def on_run_leave(e):
            run_btn.config(bg="#e67e22", relief="raised")
        
        run_btn.bind("<Enter>", on_run_enter)
        run_btn.bind("<Leave>", on_run_leave)

        # Collapsible Advanced Settings
        opt_container = ttk.Frame(footer)
        opt_container.pack(side="bottom", fill="x", pady=(10, 0))
        
        # Toggle button for advanced settings
        self.opt_expanded = tk.BooleanVar(value=False)
        
        def toggle_advanced():
            if self.opt_expanded.get():
                opt_frame.pack_forget()
                toggle_btn.config(text="▶ 고급 설정 (펼치기)")
                self.opt_expanded.set(False)
            else:
                opt_frame.pack(side="bottom", fill="x", pady=(5, 0))
                toggle_btn.config(text="▼ 고급 설정 (접기)")
                self.opt_expanded.set(True)
        
        toggle_btn = ttk.Button(opt_container, text="▶ 고급 설정 (펼치기)", command=toggle_advanced)
        toggle_btn.pack(side="top", fill="x")
        ToolTip(toggle_btn, "오타 보정, 치환 설정 등 고급 기능을 표시/숨김합니다")
        
        # Advanced settings frame (initially hidden)
        opt_frame = ttk.Frame(opt_container, padding=10)
        # Don't pack initially - will be shown on toggle
        
        replace_btn = ttk.Button(opt_frame, text="치환 설정 (Replace)", command=self.open_replacer)
        replace_btn.pack(side="left", padx=(0, 20))
        ToolTip(replace_btn, "데이터를 변환할 규칙을 설정합니다\n예: '남' → 'M', '여' → 'F'")
        
        fuzzy_check = ttk.Checkbutton(opt_frame, text="오타 보정 (Fuzzy Match)", variable=self.opt_fuzzy)
        fuzzy_check.pack(side="left", padx=(0, 10))
        ToolTip(fuzzy_check, "오타를 자동으로 보정합니다\n예: '홍길동' ≈ '홍길둥' (유사도 90% 이상)")
        
        color_check = ttk.Checkbutton(opt_frame, text="색상 강조 (Highlight)", variable=self.opt_color)
        color_check.pack(side="left")
        ToolTip(color_check, "매칭된 행에 색상을 추가하여 결과를 쉽게 확인할 수 있습니다")

        # --- 2. Header Content (Top) ---
        # Pack top content with side="top"
        top_content = ttk.Frame(main)
        top_content.pack(side="top", fill="x")

        self.base_ui = SourceFrame(top_content, "1. 기준 데이터 (Key 보유)", self.base_mode, is_base=True)
        self.base_ui.pack(fill="x", pady=(0, 15))

        self.tgt_ui = SourceFrame(top_content, "2. 대상 데이터 (데이터 가져올 곳)", self.tgt_mode)
        self.tgt_ui.pack(fill="x", pady=(0, 15))

        self.base_ui.on_change = self._load_base_cols
        self.tgt_ui.on_change = self._load_tgt_cols

        preset_frame = ttk.Frame(top_content)
        preset_frame.pack(fill="x", pady=(5, 5))
        ttk.Label(preset_frame, text="자주 쓰는 컬럼:", font=get_system_font()).pack(side="left")

        self.cb_preset = ttk.Combobox(preset_frame, state="readonly", width=25)
        self.cb_preset.pack(side="left", padx=5)
        self.cb_preset.bind("<<ComboboxSelected>>", self.apply_preset)
        ToolTip(self.cb_preset, "저장된 컬럼 설정을 불러옵니다\n자주 사용하는 매칭 조합을 빠르게 적용할 수 있습니다")
        
        save_preset_btn = ttk.Button(preset_frame, text="선택 저장 (Save)", command=self.save_preset)
        save_preset_btn.pack(side="left", padx=2)
        ToolTip(save_preset_btn, "현재 선택한 컬럼 조합을 저장합니다\n다음에 빠르게 불러올 수 있습니다")
        
        del_preset_btn = ttk.Button(preset_frame, text="삭제 (Del)", command=self.delete_preset)
        del_preset_btn.pack(side="left", padx=2)
        ToolTip(del_preset_btn, "선택한 프리셋을 삭제합니다")

        # --- 3. Middle Content (Remaining Space) ---
        mid_content = ttk.Frame(main)
        mid_content.pack(side="top", fill="both", expand=True)

        ttk.Label(
            mid_content, 
            text="가져올 컬럼 선택 (4열 보기):", 
            font=(get_system_font()[0], 11, "bold")
        ).pack(anchor="w", pady=(5, 5))
        
        self.col_list = GridCheckList(mid_content, height=4)  # Further reduced to 4 rows
        self.col_list.pack(fill="both", expand=True, pady=5)

        btns = ttk.Frame(mid_content)
        btns.pack(fill="x", pady=5)
        
        # Style improvement for buttons
        btn_all = ttk.Button(btns, text="전체 선택 (All)", command=self.col_list.check_all)
        btn_all.pack(side="left", fill="x", expand=True, padx=(0, 2))
        
        btn_none = ttk.Button(btns, text="선택 해제 (None)", command=self.col_list.uncheck_all)
        btn_none.pack(side="left", fill="x", expand=True, padx=(2, 0))
        
        # Force menu update
        def force_menu():
            try:
                if hasattr(self, 'menubar'):
                    self.config(menu=self.menubar)
                self.update_idletasks()
            except:
                pass
        self.after(500, force_menu)

        # --- COMMERCIAL FOOTER (pack BEFORE main to stay at bottom) ---
        from commercial_config import CREATOR_NAME, SALES_INFO, PRICE_INFO, CONTACT_INFO, TRIAL_EMAIL, TRIAL_SUBJECT, SUPPORT_MESSAGE, DONATION_CTA
        c_footer = tk.Frame(self, bg="#2c3e50")
        c_footer.pack(side="bottom", fill="x", pady=0)
        c_footer.config(height=60)
        
        # Left: Creator name
        tk.Label(
            c_footer,
            text=f"Made by {CREATOR_NAME}",
            bg="#2c3e50",
            fg="#95a5a6",
            font=(get_system_font()[0], 10)
        ).pack(side="left", padx=20, pady=15)
        
        # Comprehensive inquiry popup
        def show_inquiry_popup(e=None):
            top = tk.Toplevel(self)
            top.title("문의")
            top.geometry("450x600")
            top.configure(bg="white")
            top.resizable(False, False)
            
            # Center
            top.update_idletasks()
            x = (top.winfo_screenwidth() // 2) - 225
            y = (top.winfo_screenheight() // 2) - 300
            top.geometry(f"450x600+{x}+{y}")
            
            # Main container
            container = tk.Frame(top, bg="white")
            container.pack(fill="both", expand=True, padx=25, pady=20)
            
            # Title (centered)
            tk.Label(
                container,
                text="문의",
                font=(get_system_font()[0], 16, "bold"),
                bg="white",
                fg="#2c3e50"
            ).pack(pady=(0, 20), anchor="center")
            
            # Section 1: Pricing
            price_frame = tk.Frame(container, bg="#ecf0f1", relief="solid", borderwidth=1)
            price_frame.pack(fill="x", pady=(0, 12))
            
            tk.Label(price_frame, text="▶ 가격 안내", font=(get_system_font()[0], 12, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(anchor="w", padx=12, pady=(10, 4))
            tk.Label(price_frame, text="개인: 1년 33,000원 / 평생 88,000원 (최대 50,000행)", font=(get_system_font()[0], 9), bg="#ecf0f1", fg="#34495e", anchor="w").pack(anchor="w", padx=15, pady=1)
            tk.Label(price_frame, text="기업: 영구 180,000원 (무제한)", font=(get_system_font()[0], 9, "bold"), bg="#ecf0f1", fg="#c0392b", anchor="w").pack(anchor="w", padx=15, pady=(1, 10))
            
            # Section 2: Payment/Donation Account
            payment_frame = tk.Frame(container, bg="#e8f5e9", relief="solid", borderwidth=1)
            payment_frame.pack(fill="x", pady=(0, 12))
            
            tk.Label(payment_frame, text="▶ 후원 계좌", font=(get_system_font()[0], 12, "bold"), bg="#e8f5e9", fg="#2c3e50").pack(anchor="w", padx=12, pady=(10, 4))
            tk.Label(payment_frame, text="대구은행 508-14-202118-7 (이현주)", font=(get_system_font()[0], 10, "bold"), bg="#e8f5e9", fg="#c0392b", anchor="w").pack(anchor="w", padx=15, pady=2)
            
            def copy_account():
                top.clipboard_clear()
                top.clipboard_append("508-14-202118-7")
                top.update()
                messagebox.showinfo("완료", "계좌번호가 복사되었습니다!", parent=top)
            
            tk.Button(payment_frame, text="계좌번호 복사", command=copy_account, bg="#1e3a8a", fg="#ffffff", font=(get_system_font()[0], 9, "bold"), padx=12, pady=4, cursor="hand2", relief="raised", borderwidth=2, activebackground="#1e40af", activeforeground="#ffffff").pack(anchor="w", padx=15, pady=(4, 10))
            
            # Section 3: Customization Contact
            contact_frame = tk.Frame(container, bg="#fff3e0", relief="solid", borderwidth=1)
            contact_frame.pack(fill="x", pady=(0, 20))
            
            tk.Label(contact_frame, text="▶ 커스터마이징 문의", font=(get_system_font()[0], 12, "bold"), bg="#fff3e0", fg="#2c3e50").pack(anchor="w", padx=12, pady=(10, 4))
            tk.Label(contact_frame, text="bough38@gmail.com", font=(get_system_font()[0], 10, "bold"), bg="#fff3e0", fg="#c0392b", anchor="w").pack(anchor="w", padx=15, pady=2)
            
            def copy_email():
                top.clipboard_clear()
                top.clipboard_append("bough38@gmail.com")
                top.update()
                messagebox.showinfo("완료", "이메일 주소가 복사되었습니다!", parent=top)
            
            tk.Button(contact_frame, text="이메일 복사", command=copy_email, bg="#1e3a8a", fg="#ffffff", font=(get_system_font()[0], 9, "bold"), padx=12, pady=4, cursor="hand2", relief="raised", borderwidth=2, activebackground="#1e40af", activeforeground="#ffffff").pack(anchor="w", padx=15, pady=(4, 10))
            
            # Close button (centered, larger)
            tk.Button(container, text="닫기", command=top.destroy, bg="#95a5a6", fg="white", font=(get_system_font()[0], 11, "bold"), padx=30, pady=8, cursor="hand2", relief="raised", borderwidth=1).pack(pady=(10, 0), anchor="center")
        
        # Right: Inquiry button
        inquiry_btn = tk.Button(
            c_footer,
            text="문의",
            command=show_inquiry_popup,
            bg="#3498db",
            fg="white",
            font=(get_system_font()[0], 11, "bold"),
            padx=20,
            pady=8,
            cursor="hand2",
            relief="raised",
            borderwidth=2
        )
        inquiry_btn.pack(side="right", padx=20, pady=10)
        
        def on_enter_inquiry(e):
            inquiry_btn.config(bg="#2980b9", relief="sunken")
        def on_leave_inquiry(e):
            inquiry_btn.config(bg="#3498db", relief="raised")
        
        inquiry_btn.bind("<Enter>", on_enter_inquiry)
        inquiry_btn.bind("<Leave>", on_leave_inquiry)
        
        # Pack main frame AFTER footer
        main.pack(fill="both", expand=True)

        # 초기 로드
        self._load_base_cols()
        self._load_tgt_cols()

    def _log(self, msg: str):
        self.log_txt.config(state="normal")
        self.log_txt.insert("end", f"- {msg}\n")
        self.log_txt.see("end")
        self.log_txt.config(state="disabled")
        self.update_idletasks()

    def open_replacer(self):
        if self.replacer_win is None or not self.replacer_win.winfo_exists():
            self.replacer_win = ReplacementEditor(self)
        else:
            self.replacer_win.lift()

    def register_license(self):
        from license_manager import validate_key, save_license_key
        
        # 1. Input Key
        key = simpledialog.askstring("라이선스 등록", "제품 키를 입력하세요:\n(예: EM-XXXX-XXXX...)", parent=self)
        if not key:
            return
            
        key = key.strip()
        if not key:
            return

        # 2. Validate
        valid, info = validate_key(key)
        if not valid:
            messagebox.showerror("등록 실패", "유효하지 않은 라이선스 키입니다.")
            return

        # 3. Save
        try:
            save_license_key(key)
            self.license_info = info
            
            # 4. Update UI (Menu)
            # Re-create menu or just show success message (Menu update is tricky dynamically without refactor)
            # For now, just show success and tell to restart for full effect if needed, 
            # though current runtime specific limits might not update until restart depending on logic.
            # But we can update the license_info dict which is used in run().
            
            messagebox.showinfo("등록 성공", f"라이선스가 등록되었습니다.\n타입: {info.get('type')}\n만료: {info.get('expiry')}")
            
        except Exception as e:
            messagebox.showerror("오류", f"라이선스 저장 중 오류가 발생했습니다.\n{e}")

    # ----------------
    # Preset (columns)
    # ----------------
    def load_presets(self):
        try:
            if os.path.exists(PRESET_FILE):
                with open(PRESET_FILE, "r", encoding="utf-8") as f:
                    self.presets = json.load(f) or {}
            else:
                self.presets = {}
        except Exception:
            self.presets = {}

        self.cb_preset["values"] = list(self.presets.keys())
        self.cb_preset.set("설정을 선택하세요" if self.presets else "저장된 설정 없음")

    def _save_presets(self):
        with open(PRESET_FILE, "w", encoding="utf-8") as f:
            json.dump(self.presets, f, ensure_ascii=False, indent=4)

    def save_preset(self):
        checked = self.col_list.checked()
        if not checked:
            messagebox.showwarning("경고", "저장할 컬럼을 하나 이상 체크해주세요.")
            return
        name = simpledialog.askstring("설정 저장", "이 설정의 이름을 입력하세요:\n(예: 급여대장용)")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        self.presets[name] = checked
        self._save_presets()
        self.load_presets()
        self.cb_preset.set(name)
        messagebox.showinfo("저장 완료", f"[{name}] 설정이 저장되었습니다.")

    def delete_preset(self):
        name = self.cb_preset.get()
        if name in self.presets and messagebox.askyesno("삭제 확인", f"정말 [{name}] 설정을 삭제하시겠습니까?"):
            del self.presets[name]
            self._save_presets()
            self.load_presets()

    def apply_preset(self, event=None):
        name = self.cb_preset.get()
        if name in self.presets:
            cnt = self.col_list.set_checked_items(self.presets[name])
            self._log(f"프리셋 [{name}] 적용됨 ({cnt}개 항목 선택)")

    # -------------
    # Header loaders
    # -------------
    def _load_base_cols(self):
        cfg = self.base_ui.get_config()
        if not cfg.get("sheet"):
            return
        try:
            if cfg["type"] == "file":
                if not cfg["path"] or not os.path.exists(cfg["path"]):
                    return
                cols = read_header_file(cfg["path"], cfg["sheet"], cfg["header"])
            else:
                if not cfg["book"]:
                    return
                cols = read_header_open(cfg["book"], cfg["sheet"], cfg["header"])

            self.base_ui.key_listbox.set_items(cols)
            if cols:
                self.base_ui.key_listbox.select_item(cols[0])
        except Exception as e:
            self._log(f"기준 헤더 로드 실패: {e}")

    def _load_tgt_cols(self):
        cfg = self.tgt_ui.get_config()
        if not cfg.get("sheet"):
            return
        try:
            if cfg["type"] == "file":
                if not cfg["path"] or not os.path.exists(cfg["path"]):
                    return
                cols = read_header_file(cfg["path"], cfg["sheet"], cfg["header"])
            else:
                if not cfg["book"]:
                    return
                cols = read_header_open(cfg["book"], cfg["sheet"], cfg["header"])

            self.col_list.set_items(cols)
            self._log(f"대상 컬럼 로드됨 ({len(cols)}개)")
        except Exception as e:
            self._log(f"대상 헤더 로드 실패: {e}")

    # ----
    # Run
    # ----
    def run(self):
        """Execute matching with progress dialog"""
        try:
            # Validate inputs
            b_cfg = self.base_ui.get_config()
            t_cfg = self.tgt_ui.get_config()
            keys = self.base_ui.get_selected_keys()
            take = self.col_list.checked()

            if not keys:
                messagebox.showwarning("경고", "매칭할 키(Key)를 하나 이상 선택하세요.")
                return
            if not take:
                messagebox.showwarning("경고", "가져올 컬럼을 선택하세요.")
                return

            options = {"fuzzy": self.opt_fuzzy.get(), "color": self.opt_color.get()}

            replace_rules = (
                self.replacer_win.get_rules()
                if self.replacer_win and self.replacer_win.winfo_exists()
                else (_load_replace_file().get("active", {}) or {})
            )

            # Create progress dialog
            self._show_progress_dialog(b_cfg, t_cfg, keys, take, options, replace_rules)

        except Exception as e:
            traceback.print_exc()
            msg = str(e)

            if "xlwings" in msg.lower():
                msg += (
                    "\n\n[힌트]\n"
                    "- Excel 설치 확인\n"
                    "- macOS: 개인정보 보호 및 보안 > 자동화에서 Terminal/iTerm2의 Excel 제어 허용\n"
                    "- 파일 모드로 사용 가능"
                )

            messagebox.showerror("오류", f"실행 중 오류:\n{msg}")
            self._log(f"Error: {msg}")

    def _show_progress_dialog(self, b_cfg, t_cfg, keys, take, options, replace_rules):
        """Show progress dialog and run matching in background thread"""
        import threading
        
        # Create progress window
        progress_win = tk.Toplevel(self)
        progress_win.title("매칭 진행 중...")
        progress_win.geometry("500x200")
        progress_win.resizable(False, False)
        progress_win.transient(self)
        progress_win.grab_set()
        
        # Center the window
        progress_win.update_idletasks()
        x = (progress_win.winfo_screenwidth() // 2) - (500 // 2)
        y = (progress_win.winfo_screenheight() // 2) - (200 // 2)
        progress_win.geometry(f"500x200+{x}+{y}")
        
        # Progress frame
        frame = tk.Frame(progress_win, bg="white", padx=30, pady=30)
        frame.pack(fill="both", expand=True)
        
        # Title
        tk.Label(
            frame,
            text="데이터 매칭 중...",
            font=(get_system_font()[0], 14, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(pady=(0, 15))
        
        # Progress bar
        progress_bar = ttk.Progressbar(
            frame,
            mode='determinate',
            length=400,
            maximum=100
        )
        progress_bar.pack(pady=(0, 10))
        
        # Status label
        status_label = tk.Label(
            frame,
            text="준비 중...",
            font=(get_system_font()[0], 11),
            bg="white",
            fg="#7f8c8d"
        )
        status_label.pack(pady=(0, 15))
        
        # Cancel button
        cancel_flag = {"cancelled": False}
        
        def cancel_matching():
            cancel_flag["cancelled"] = True
            cancel_btn.config(state="disabled", text="취소 중...")
        
        cancel_btn = tk.Button(
            frame,
            text="취소",
            command=cancel_matching,
            bg="#e74c3c",
            fg="white",
            font=(get_system_font()[0], 11, "bold"),
            padx=20,
            pady=5,
            cursor="hand2"
        )
        cancel_btn.pack()
        
        # Result storage
        result = {"out_path": None, "summary": None, "error": None}
        
        def update_progress(value, message):
            """Update progress from worker thread (thread-safe)"""
            def _update():
                if not progress_win.winfo_exists():
                    return
                progress_bar['value'] = value
                status_label.config(text=message)
                progress_win.update_idletasks()
            
            # Schedule update in main thread
            self.after(0, _update)
        
        def worker_thread():
            """Background worker thread for matching"""
            try:
                if cancel_flag["cancelled"]:
                    return
                
                update_progress(10, "파일 읽는 중...")
                
                # Perform matching with progress updates
                out_path, summary = match_universal(
                    b_cfg, t_cfg, keys, take, OUT_DIR, options, replace_rules, 
                    lambda msg: update_progress(
                        min(90, progress_bar['value'] + 10), 
                        msg if len(msg) < 50 else msg[:47] + "..."
                    )
                )
                
                if cancel_flag["cancelled"]:
                    return
                
                update_progress(95, "결과 저장 중...")
                result["out_path"] = out_path
                result["summary"] = summary
                
                update_progress(100, "완료!")
                
                # Close progress window after short delay
                def close_and_show_result():
                    if progress_win.winfo_exists():
                        progress_win.destroy()
                    
                    if not cancel_flag["cancelled"]:
                        msg = (
                            "작업 완료!\n\n"
                            "[결과 리포트]\n"
                            f"{summary}\n\n"
                            "저장 위치:\n"
                            f"{os.path.basename(out_path)}"
                        )
                        messagebox.showinfo("성공", msg)
                
                self.after(500, close_and_show_result)
                
            except Exception as e:
                result["error"] = str(e)
                
                def show_error():
                    if progress_win.winfo_exists():
                        progress_win.destroy()
                    
                    msg = str(e)
                    if "xlwings" in msg.lower():
                        msg += (
                            "\n\n[힌트]\n"
                            "- Excel 설치 확인\n"
                            "- macOS: 개인정보 보호 및 보안 > 자동화에서 Terminal/iTerm2의 Excel 제어 허용\n"
                            "- 파일 모드로 사용 가능"
                        )
                    
                    messagebox.showerror("오류", f"실행 중 오류:\n{msg}")
                    self._log(f"Error: {msg}")
                
                self.after(0, show_error)
        
        # Start worker thread
        thread = threading.Thread(target=worker_thread, daemon=True)
        thread.start()
        
        # Handle window close
        def on_closing():
            cancel_flag["cancelled"] = True
            progress_win.destroy()
        
        progress_win.protocol("WM_DELETE_WINDOW", on_closing)