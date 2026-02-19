import tkinter as tk
from tkinter import ttk
import sys

class GuideBubble:
    def __init__(self, master, target_widget, text, step_num, total_steps, on_next=None, on_skip=None):
        self.master = master
        self.target = target_widget
        self.on_next = on_next
        self.on_skip = on_skip
        
        self.top = tk.Toplevel(master)
        self.top.overrideredirect(True)  # No window decorations
        self.top.attributes('-topmost', True)
        self.top.lift()
        
        # Premium Colors (Dark Theme with Gold/Orange Accent)
        self.bg_color = "#2c3e50"     # Dark Blue-Grey
        self.fg_color = "#ecf0f1"     # Off-White
        self.accent_color = "#f39c12" # Orange-Gold
        self.border_color = "#f39c12" # Gold Border
        
        # Outer Border Frame (for "Box Border" effect)
        self.border_frame = tk.Frame(self.top, bg=self.border_color, padx=2, pady=2)
        self.border_frame.pack(fill="both", expand=True)
        
        # Main Content Frame
        self.frame = tk.Frame(self.border_frame, bg=self.bg_color)
        self.frame.pack(fill="both", expand=True)
        
        # Inner padding frame
        inner = tk.Frame(self.frame, bg=self.bg_color)
        inner.pack(padx=20, pady=20)
        
        # Header: Step indicator + Title styling
        header_frame = tk.Frame(inner, bg=self.bg_color)
        header_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(header_frame, text=f"STEP {step_num}", 
                 bg=self.bg_color, fg=self.accent_color, font=("Segoe UI", 8, "bold")).pack(side="left")
        
        tk.Label(header_frame, text=f" / {total_steps}", 
                 bg=self.bg_color, fg="#95a5a6", font=("Segoe UI", 8)).pack(side="left")

        # Main Text with improved typography
        tk.Label(inner, text=text, bg=self.bg_color, fg=self.fg_color, 
                 font=("Segoe UI", 11), justify="left", wraplength=280).pack(pady=(0, 20), anchor="w")
        
        # Buttons Frame
        btn_frame = tk.Frame(inner, bg=self.bg_color)
        btn_frame.pack(fill="x", pady=(5, 0))
        
        # Skip Button (Subtle)
        lbl_skip = tk.Label(btn_frame, text="건너뛰기", bg=self.bg_color, fg="#7f8c8d", 
                            cursor="hand2", font=("Segoe UI", 9, "underline"))
        lbl_skip.pack(side="left", anchor="s")
        lbl_skip.bind("<Button-1>", lambda e: self.close(skip=True))
        
        # Next Button (Premium Look)
        next_text = "다음 >" if step_num < total_steps else "완료 (Finish)"
        btn_next = tk.Label(btn_frame, text=next_text, 
                            bg=self.accent_color, fg="white", 
                            padx=15, pady=6, cursor="hand2", font=("Segoe UI", 10, "bold"),
                            relief="flat")
        btn_next.pack(side="right")
        btn_next.bind("<Button-1>", lambda e: self.next_step())
        
        # Button Hover Effects
        def on_next_enter(e): btn_next.config(bg="#e67e22") # Darker orange
        def on_next_leave(e): btn_next.config(bg=self.accent_color)
        btn_next.bind("<Enter>", on_next_enter)
        btn_next.bind("<Leave>", on_next_leave)
        
        def on_skip_enter(e): lbl_skip.config(fg="#bdc3c7") # Lighter grey
        def on_skip_leave(e): lbl_skip.config(fg="#7f8c8d")
        lbl_skip.bind("<Enter>", on_skip_enter)
        lbl_skip.bind("<Leave>", on_skip_leave)
        
        # Position the bubble
        self.update_position()
        
        # Bind to logic
        self.top.bind("<Escape>", lambda e: self.close(skip=True))

    def update_position(self):
        try:
            # Get target coordinates
            self.target.update_idletasks()
            x = self.target.winfo_rootx()
            y = self.target.winfo_rooty()
            w = self.target.winfo_width()
            h = self.target.winfo_height()
            
            # Bubble size (Get actual size)
            self.top.update_idletasks()
            bw = self.top.winfo_reqwidth()
            bh = self.top.winfo_reqheight()
            
            # Calculate position (try to place below, then above, then right)
            # Default: Below center
            bx = x + (w // 2) - (bw // 2)
            by = y + h + 15
            
            # Screen bounds
            sw = self.target.winfo_screenwidth()
            sh = self.target.winfo_screenheight()
            
            # Adjust X
            if bx < 10: bx = 10
            if bx + bw > sw - 10: bx = sw - bw - 10
            
            # Adjust Y (flip to top if cutting off bottom)
            # Add a small buffer (10px) to ensure it doesn't touch the edge
            if by + bh > sh - 10:
                by = y - bh - 15
            
            self.top.geometry(f"+{int(bx)}+{int(by)}")
        except:
            pass

    def next_step(self):
        self.top.destroy()
        if self.on_next:
            self.on_next()

    def close(self, skip=False):
        self.top.destroy()
        if skip and self.on_skip:
            self.on_skip()

class GuideManager:
    def __init__(self, master):
        self.master = master
        self.steps = []
        self.current_step_index = 0
        self.current_bubble = None
        self.is_running = False

    def set_steps(self, steps):
        """
        steps: list of dicts with 'widget' (string attr name or widget obj) and 'text' keys
        """
        self.steps = steps

    def start_guide(self):
        if not self.steps: return
        self.is_running = True
        self.current_step_index = 0
        self.show_current_step()

    def show_current_step(self):
        if not self.is_running: return
        if self.current_step_index >= len(self.steps):
            self.stop_guide()
            return
            
        step_data = self.steps[self.current_step_index]
        target = step_data['widget']
        
        # Resolve target if it's a string attribute name of master
        if isinstance(target, str):
            if hasattr(self.master, target):
                target = getattr(self.master, target)
            else:
                print(f"Guide Target not found: {target}")
                self.stop_guide()
                return

        # Ensure target is visible
        try:
            # If target is in a hidden tab or frame, we might need logic here to show it
            pass
        except:
            pass
            
        text = step_data['text']
        
        self.current_bubble = GuideBubble(
            self.master, 
            target, 
            text, 
            self.current_step_index + 1, 
            len(self.steps),
            on_next=self.advance,
            on_skip=self.stop_guide
        )

    def advance(self):
        self.current_step_index += 1
        self.show_current_step()

    def stop_guide(self):
        self.is_running = False
        if self.current_bubble:
            try:
                self.current_bubble.top.destroy()
            except:
                pass
        self.current_bubble = None
