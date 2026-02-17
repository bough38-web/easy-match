import tkinter as tk
from tkinter import ttk
import sys

def get_system_font():
    return ("Arial", 11)

class ToolTip:
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.id_after = None
        widget.bind("<Enter>", self.on_enter)
        widget.bind("<Leave>", self.on_leave)
    
    def on_enter(self, event=None):
        self.id_after = self.widget.after(self.delay, self.show_tooltip)
    
    def on_leave(self, event=None):
        if self.id_after:
            self.widget.after_cancel(self.id_after)
            self.id_after = None
        self.hide_tooltip()
    
    def show_tooltip(self):
        if self.tooltip_window: return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, background="#2c3e50", foreground="#ecf0f1", relief="solid", borderwidth=1, font=(get_system_font()[0], 10))
        label.pack()
    
    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def main():
    root = tk.Tk()
    root.title("Test UI Crash")
    root.geometry("400x300")
    
    btn1 = ttk.Button(root, text="Test Tooltip")
    btn1.pack(pady=20)
    ToolTip(btn1, "This is a test tooltip. (예: 테스트)")
    
    lbl = tk.Label(root, text="Test Label with Colors", bg="#16a085", fg="white", font=("Arial", 15, "bold"))
    lbl.pack(fill="x", pady=10)
    
    print("UI Launching...")
    root.after(1000, lambda: print("UI still alive..."))
    root.after(2000, lambda: root.destroy())
    root.mainloop()
    print("UI Closed safely.")

if __name__ == "__main__":
    main()
