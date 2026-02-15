import tkinter as tk
from tkinter import ttk, messagebox
from license_manager import save_license, load_license

class AdminPanel(tk.Toplevel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title("관리자 - 라이선스 설정")
        self.geometry("360x240")
        self.resizable(False, False)

        # Use Notebook for tabs
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Current License Info (Read-only view)
        tab_info = ttk.Frame(notebook, padding=10)
        notebook.add(tab_info, text="현재 라이선스")
        
        ttk.Label(tab_info, text="만료일 (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        self.ent_expiry = ttk.Entry(tab_info, width=20)
        self.ent_expiry.grid(row=0, column=1, pady=6, sticky="w")
        self.ent_expiry.insert(0, "-")
        self.ent_expiry.config(state="readonly") # Managed via Key now

        ttk.Label(tab_info, text="라이선스 타입").grid(row=1, column=0, sticky="w")
        self.type_var = tk.StringVar(value="personal")
        self.cb_type = ttk.Combobox(tab_info, textvariable=self.type_var, values=["personal","enterprise"], state="disabled", width=17)
        self.cb_type.grid(row=1, column=1, pady=6, sticky="w")

        ttk.Label(tab_info, text="* 라이선스 변경은 '제품 키 입력'을 통해 가능합니다.").grid(row=2, column=0, columnspan=2, pady=10)

        # Load current info
        try:
            from license_manager import validate_license
            ok, msg, info = validate_license()
            if ok and info:
                self.ent_expiry.config(state="normal")
                self.ent_expiry.delete(0, "end")
                self.ent_expiry.insert(0, info.get("expiry", "-"))
                self.ent_expiry.config(state="readonly")
                self.type_var.set(info.get("type", "personal"))
        except:
            pass

        # Tab 2: Key Generator (For Admin/Distributor)
        tab_gen = ttk.Frame(notebook, padding=10)
        notebook.add(tab_gen, text="키 생성기 (관리자용)")
        
        ttk.Label(tab_gen, text="만료일 (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        self.gen_expiry = ttk.Entry(tab_gen, width=20)
        self.gen_expiry.grid(row=0, column=1, pady=6, sticky="w")
        import datetime
        next_year = datetime.date.today().replace(year=datetime.date.today().year + 1)
        self.gen_expiry.insert(0, next_year.strftime("%Y-%m-%d"))

        ttk.Label(tab_gen, text="라이선스 타입").grid(row=1, column=0, sticky="w")
        self.gen_type_var = tk.StringVar(value="personal")
        ttk.Combobox(tab_gen, textvariable=self.gen_type_var, values=["personal","enterprise"], state="readonly", width=17).grid(row=1, column=1, pady=6, sticky="w")

        ttk.Button(tab_gen, text="키 생성 및 복사", command=self.generate_key_action).grid(row=2, column=0, columnspan=2, pady=15, sticky="ew")
        
        self.txt_key_out = tk.Text(tab_gen, height=4, width=35)
        self.txt_key_out.grid(row=3, column=0, columnspan=2)

    def generate_key_action(self):
        try:
            from license_key import generate_key
        except ImportError:
            messagebox.showerror("오류", "license_key 모듈을 찾을 수 없습니다.")
            return

        expiry = self.gen_expiry.get().strip()
        l_type = self.gen_type_var.get().strip()
        
        # Validation
        try:
            import datetime
            datetime.datetime.strptime(expiry, "%Y-%m-%d")
        except:
            messagebox.showerror("오류", "날짜 형식이 올바르지 않습니다.")
            return
            
        new_key = generate_key(expiry, l_type)
        self.txt_key_out.delete("1.0", "end")
        self.txt_key_out.insert("end", new_key)
        
        # Copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(new_key)
        messagebox.showinfo("성공", "키가 생성되고 클립보드에 복사되었습니다.")


def main():
    root = tk.Tk()
    root.withdraw()
    win = AdminPanel(root)
    win.mainloop()

if __name__ == "__main__":
    main()
