import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
import time
import threading
import os
import pyperclip
import keyboard  # Import thư viện keyboard
from tkinter import font  # Import thư viện font

class PasteTool:
    def __init__(self, master):
        self.master = master
        master.title("Paste Auto")
        master.geometry("320x495")
        master.resizable(False, False)

        self.paste_thread = None
        self.stop_event = threading.Event()
        self.is_running = False
        self.radiobutton_original_colors = {}
        self.is_ready = False  # Thêm trạng thái sẵn sàng
        self.is_completed = False # Thêm trạng thái hoàn thành
        self.is_locked_after_completion = False # Thêm trạng thái khóa sau khi hoàn thành
        self.is_stopped_by_esc = False # Thêm trạng thái dừng bằng ESC
        self.paused_lines = [] # Thêm danh sách lưu trữ các dòng đã tạm dừng
        #self.current_line_index = 0 # Thêm biến lưu trữ chỉ số dòng hiện tại # Removed this line

        # --- Khu vực nhập liệu ---
        input_frame = tk.Frame(master, padx=10, pady=10)
        input_frame.grid(row=0, column=0, sticky="nsew")

        tk.Label(input_frame, text="Nhập dữ liệu").pack(anchor="w")

        # Lấy font Arial mặc định
        arial_font = font.Font(family="Calibri", size=11)
        self.text_area = tk.Text(input_frame, wrap=tk.NONE, width=30, height=15, font=arial_font) # Thay đổi wrap=tk.NONE, thêm font=arial_font
        self.text_area.pack(fill="both", expand=True)
        self.text_area.bind("<<Modified>>", self.on_text_modified)
        self.text_modified_flag = False
        self.text_area_original_bg = self.text_area.cget("bg") # Lưu lại màu nền ban đầu

        # --- Khu vực điều khiển ---
        control_frame = tk.Frame(master, padx=10, pady=3)  # Giảm padding theo chiều dọc
        control_frame.grid(row=0, column=1, sticky="ns", rowspan=2)

        # Tốc độ
        tk.Label(control_frame, text="Tốc độ (ms):").pack(anchor="w")
        self.speed_var = tk.IntVar(value=250)
        speeds = [150, 200, 250, 350, 400, 450, 500]
        self.radiobuttons = []
        for speed in speeds:
            rb = tk.Radiobutton(control_frame, text=f"{speed} ms", variable=self.speed_var, value=speed, command=self.update_current_speed_label)
            rb.pack(anchor="w")
            self.radiobuttons.append(rb)
            self.radiobutton_original_colors[rb] = rb.cget("fg")

        self.current_speed_label = tk.Label(control_frame, text=f"Hiện tại: {self.speed_var.get()} ms")
        self.current_speed_label.pack(anchor="w", pady=(0, 3)) # Giảm padding dưới

        # Các nút điều khiển
        button_width = 12  # Tăng chiều rộng nút
        button_height = 1 # Tăng chiều cao nút

        self.start_button = tk.Button(control_frame, text="START", command=self.prepare_pasting, bg="#32CD32", fg="#000000", width=10, height=2)
        self.start_button.pack(pady=(0, 4)) # Giảm padding trên và dưới

        self.stop_button = tk.Button(control_frame, text="STOP", command=self.stop_pasting, bg="#FA8072", fg="#000000", width=10, height=2, state=tk.DISABLED)
        self.stop_button.pack(pady=(0, 4)) # Giảm padding trên và dưới
        self.stop_button_disabled_fg = "#FFF" # Màu chữ khi disable cho nút STOP

        self.reset_button = tk.Button(control_frame, text="RESET", command=self.reset_all, bg="#FFCC99", fg="#000000", width=10, height=2)
        self.reset_button.pack(pady=(0, 4)) # Giảm padding trên và dưới

        self.upload_button = tk.Button(control_frame, text="LOAD FILE\n.csv .txt", command=self.upload_file, bg="#CDCDB4", fg="#000", width=10, height=2)
        self.upload_button.pack(pady=(0, 1)) # Giảm padding trên và dưới

        # Trạng thái
        tk.Label(control_frame, text="Trạng thái:").pack(anchor="w", pady=(5, 0)) # Giảm padding trên
        self.status_var = tk.StringVar(value="Đang chờ")
        self.status_label = tk.Label(control_frame, textvariable=self.status_var, fg="blue")
        self.status_label.pack(anchor="w")

        # --- Khu vực thông tin ---
        info_frame = tk.Frame(master, padx=10, pady=5)
        info_frame.grid(row=1, column=0, sticky="ew")

        self.line_count_var = tk.StringVar(value="Line: 0")
        self.line_count_label = tk.Label(info_frame, textvariable=self.line_count_var)
        self.line_count_label.pack(side="left")

        # Nhãn bản quyền
        self.copyright_label = tk.Label(master, text="2025 ©Nông Văn Phấn", fg="#CCC", cursor="hand2") # Add cursor="hand2"
        self.copyright_label.place(relx=1.0, rely=1.0, x=-5, y=-5, anchor="se")
        self.copyright_label.bind("<Button-1>", self.show_about_dialog) # Add bind

        # Cấu hình grid
        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)
        master.grid_columnconfigure(1, weight=0)

        # Cập nhật số dòng ban đầu
        self.update_line_count()

        # Phím tắt
        keyboard.add_hotkey("f1", self.start_pasting, suppress=True)  # F1 để bắt đầu
        keyboard.add_hotkey("esc", self.stop_pasting_esc, suppress=True)  # ESC để dừng

        self.locked = False

    def on_text_modified(self, event=None):
        """
        Xử lý khi nội dung trong text_area thay đổi.
        """
        if not self.text_modified_flag:
            self.text_modified_flag = True
            self.master.after(100, self.update_line_count)

    def update_line_count(self):
        """
        Cập nhật số dòng trong text_area.
        """
        content = self.text_area.get("1.0", tk.END).strip()
        lines = content.splitlines()
        num_lines = len(lines) if content else 0
        self.line_count_var.set(f"Line: {num_lines}")
        self.text_area.edit_modified(False)
        self.text_modified_flag = False

    def update_current_speed_label(self):
        """
        Cập nhật nhãn tốc độ hiện tại.
        """
        self.current_speed_label.config(text=f"Hiện tại: {self.speed_var.get()} ms")

    def upload_file(self):
        """
        Tải file dữ liệu vào text_area.
        """
        if self.is_running or self.is_ready:
            messagebox.showwarning("Không thể tải file", "Không thể tải file khi đang thực hiện dán hoặc đang ở trạng thái sẵn sàng.")
            return

        filetypes = (("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*"))
        filepath = filedialog.askopenfilename(title="Chọn file dữ liệu", filetypes=filetypes)
        if filepath:
            try:
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                except UnicodeDecodeError:
                    try:
                        with open(filepath, 'r', encoding='latin-1') as f:
                            content = f.read()
                    except UnicodeDecodeError:
                        with open(filepath, 'r', encoding='cp1252') as f:
                            content = f.read()

                self.text_area.delete("1.0", tk.END)
                self.text_area.insert("1.0", content)
                self.update_line_count()
                self.status_var.set("Đã tải file.")
            except FileNotFoundError:
                messagebox.showerror("Lỗi", f"Không tìm thấy file: {filepath}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file:\n{e}")

    def prepare_pasting(self):
        """
        Chuẩn bị ứng dụng cho việc dán.
        """
        if self.locked or self.is_ready:
            return
        self.is_ready = True
        self.locked = True
        self.text_area.config(state=tk.DISABLED)
        self.text_area.config(bg="#DCDCDC") # Thay đổi màu nền khi khóa
        self.status_var.set("Sẵn sàng\nNhấn F1 để dán")
        self.update_button_states()

    def start_pasting(self):
        """
        Bắt đầu quá trình dán.
        """
        if not self.is_ready:
            if self.is_completed and self.is_locked_after_completion:
                self.reset_all()
            elif self.is_stopped_by_esc:
                self.is_stopped_by_esc = False
            else:
                return
        
        if self.is_running:
            return

        text_content = self.text_area.get("1.0", tk.END).strip()
        lines = text_content.splitlines()

        if not lines:
            messagebox.showwarning("Dữ liệu trống", "Vui lòng nhập dữ liệu vào ô trống hoặc tải file.")
            return
        
        # Khóa text_area
        self.text_area.config(state=tk.DISABLED)
        self.text_area.config(bg="#D3D3D3") # Thay đổi màu nền khi khóa

        self.is_running = True
        self.stop_event.clear()
        self.status_var.set("Đang chạy...")
        self.update_button_states()

        delay_ms = self.speed_var.get()
        delay_s = delay_ms / 1000.0

        self._run_paste_thread(lines, delay_s)

    def _run_paste_thread(self, lines, delay_s, start_index=0):
        """
        Khởi chạy luồng dán.
        """
        if not self.is_running:
            self.status_var.set("Đã hủy")
            self.update_button_states()
            return

        self.paste_thread = threading.Thread(target=self._paste_loop, args=(lines, delay_s), daemon=True)
        self.paste_thread.start()

    def _paste_loop(self, lines, delay_s):
        """
        Vòng lặp dán.
        """
        try:
            for i, line in enumerate(lines):
                if self.stop_event.is_set():
                    self.master.after(0, lambda: self.status_var.set("Đã dừng bằng\nphím ESC"))
                    self.paused_lines = lines # Lưu lại các dòng
                    #self.current_line_index = i # Lưu lại chỉ số dòng hiện tại # Removed this line
                    break

                self.master.after(0, lambda i=i, total=len(lines): self.status_var.set(f"Tiến độ: {i+1}/{total}"))

                pyperclip.copy(line)
                time.sleep(0.1)
                pyautogui.hotkey("ctrl", "v")
                pyautogui.press('enter')
                time.sleep(delay_s)

            if not self.stop_event.is_set():
                self.master.after(0, lambda: self.status_var.set("Hoàn thành"))
                self.is_completed = True # Đánh dấu đã hoàn thành
                self.is_ready = False
                self.is_locked_after_completion = True # Khóa sau khi hoàn thành
                self.master.after(0, self.update_button_states)
                self.master.after(0, self.show_completion_message) # Gọi hàm hiển thị thông báo

        except Exception as e:
            self.master.after(0, lambda e=e: messagebox.showerror("Lỗi khi đang chạy", f"Đã xảy ra lỗi:\n{e}"))
            self.master.after(0, lambda: self.status_var.set("Lỗi!"))
        finally:
            self.master.after(0, self._finalize_pasting)

    def stop_pasting(self):
        """
        Dừng quá trình dán (dùng cho nút STOP).
        """
        if self.is_running and self.paste_thread and self.paste_thread.is_alive():
            self.stop_event.set()
            self.status_var.set("Đã mở khóa\ncó thể chỉnh sửa") # Thay đổi trạng thái ở đây
            self.is_stopped_by_esc = False # Đặt lại trạng thái dừng bằng ESC
        self.is_ready = False  # Thoát trạng thái sẵn sàng
        self.locked = False
        self.is_locked_after_completion = False # Mở khóa sau khi dừng
        self.text_area.config(state=tk.NORMAL)
        self.text_area.config(bg=self.text_area_original_bg) # Trả lại màu nền ban đầu
        self.update_button_states()

    def stop_pasting_esc(self):
        """
        Dừng quá trình dán (dùng cho phím ESC).
        """
        if self.is_running and self.paste_thread and self.paste_thread.is_alive():
            self.stop_event.set()
            self.status_var.set("Đã dừng bằng\nphím ESC")
            self.is_stopped_by_esc = True # Đánh dấu dừng bằng ESC
        self.update_button_states()

    def _finalize_pasting(self):
        """
        Hoàn tất quá trình dán.
        """
        self.is_running = False
        self.stop_event.clear()
        self.paste_thread = None
        if not self.is_completed and not self.is_stopped_by_esc:
            self.update_button_states()
        self.locked = False
        # Mở khóa text_area
        if not self.is_locked_after_completion and not self.is_stopped_by_esc:
            self.text_area.config(state=tk.NORMAL)
            self.text_area.config(bg=self.text_area_original_bg) # Trả lại màu nền ban đầu
        if self.status_var.get() == "Đang dừng...":
            self.status_var.set("Đã mở khóa\ncó thể chỉnh sửa")
        self.paused_lines = [] # Xóa danh sách các dòng đã tạm dừng

    def reset_all(self):
        """
        Reset ứng dụng về trạng thái ban đầu.
        """
        if self.is_running:
            self.stop_pasting()
        elif self.is_ready:
            self.stop_pasting()

        self.text_area.delete("1.0", tk.END)
        self.speed_var.set(250)
        self.update_current_speed_label()
        self.status_var.set("Đang chờ")
        self.update_line_count()
        self.is_running = False
        self.update_button_states()
        self.is_ready = False # Thoát trạng thái sẵn sàng
        self.locked = False
        self.is_locked_after_completion = False # Mở khóa sau khi reset
        self.text_area.config(state=tk.NORMAL)
        self.text_area.config(bg=self.text_area_original_bg) # Trả lại màu nền ban đầu
        self.is_completed = False
        self.is_stopped_by_esc = False # Đặt lại trạng thái dừng bằng ESC
        self.paused_lines = [] # Xóa danh sách các dòng đã tạm dừng
        #self.current_line_index = 0 # Đặt lại chỉ số dòng hiện tại # Removed this line

    def update_button_states(self):
        """
        Cập nhật trạng thái các nút bấm.
        """
        if self.is_running:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.reset_button.config(state=tk.DISABLED)
            self.upload_button.config(state=tk.DISABLED)
            for rb in self.radiobuttons:
                rb.config(state=tk.DISABLED, fg=self.radiobutton_original_colors[rb]) # Không disable radiobutton khi đang chạy
        elif self.is_ready: # Nếu đang ở trạng thái sẵn sàng
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.reset_button.config(state=tk.NORMAL)
            self.upload_button.config(state=tk.DISABLED)
            for rb in self.radiobuttons:
                rb.config(state=tk.NORMAL, fg="#878383")
        elif self.is_completed and self.is_locked_after_completion:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED, fg=self.stop_button_disabled_fg) # Thay đổi màu chữ khi disable
            self.reset_button.config(state=tk.NORMAL)
            self.upload_button.config(state=tk.DISABLED)
            for rb in self.radiobuttons:
                rb.config(state=tk.DISABLED, fg="#878383")
        elif self.is_completed:
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED, fg=self.stop_button_disabled_fg) # Thay đổi màu chữ khi disable
            self.reset_button.config(state=tk.NORMAL)
            self.upload_button.config(state=tk.NORMAL)
            for rb in self.radiobuttons:
                rb.config(state=tk.NORMAL, fg=self.radiobutton_original_colors[rb])
        else:
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED, fg=self.stop_button_disabled_fg) # Thay đổi màu chữ khi disable
            self.reset_button.config(state=tk.NORMAL)
            self.upload_button.config(state=tk.NORMAL)
            for rb in self.radiobuttons:
                rb.config(state=tk.NORMAL, fg=self.radiobutton_original_colors[rb])

    def show_completion_message(self): #Hiển thị hộp thoại thông báo hoàn thành.
        messagebox.showinfo("Hoàn thành", "Đã hoàn thành!")

    def show_about_dialog(self, event=None):
        """
        Hiển thị hộp thoại thông tin phần mềm.
        """
        about_text = """
        Paste Tool - Công cụ dán dữ liệu tự động

        Phiên bản: 1.0
        Tác giả: Nông Văn Phấn
        Năm: 2025
        Mô tả: 
        Công cụ này giúp bạn dán dữ liệu tự động vào các ứng dụng khác.
        Bạn có thể nhập dữ liệu trực tiếp hoặc tải từ file .txt, .csv.
        Nhấn F1 để bắt đầu dán.
        Nhấn ESC để dừng dán.
        Nhấn STOP để dừng và mở khóa dữ liệu.
        Nhấn RESET để xóa dữ liệu và đặt lại.
        Source: https://github.com/phandepzai/PasteTool
        """
        messagebox.showinfo("Thông tin phần mềm", about_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = PasteTool(root)
    root.mainloop()
