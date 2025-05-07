import os
import sys
import re
import queue
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from pdf2image import convert_from_path
from pdf2image.exceptions import PDFInfoNotInstalledError
from PIL import Image, ImageTk

class PDFToImageConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Image Converter")
        self.root.geometry("500x690")  # Tăng kích thước cửa sổ
        self.root.resizable(False, False)
        
        # Thiết lập icon
        try:
            self.root.iconbitmap("icon.ico")  # Thay bằng đường dẫn icon của bạn
        except:
            pass
        
        # Queue để giao tiếp giữa các thread
        self.queue = queue.Queue()
        
        # Thiết lập đường dẫn poppler
        self.poppler_path = self.find_poppler_path()
        self.poppler_available = self.check_poppler()
        
        # Debug poppler path
        # self.debug_poppler_path()
        
        # Biến lưu trữ
        self.pdf_path = StringVar()
        self.output_folder = StringVar()
        self.pages_to_convert = StringVar()
        self.convert_all = BooleanVar(value=True)
        
        # Tạo giao diện
        self.create_widgets()
        
        # Thiết lập kiểm tra queue định kỳ
        self.root.after(100, self.process_queue)
    
    def find_poppler_path(self):
        """Tìm đường dẫn poppler theo nhiều cách khác nhau"""
        # Cách 1: Kiểm tra trong PATH hệ thống
        try:
            convert_from_path("test.pdf", first_page=1, last_page=1)
            print("Poppler found in system PATH")
            return None  # Đã có trong PATH
        except:
            pass
        
        # Cách 2: Kiểm tra thư mục poppler trong cùng thư mục với tool
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        poppler_path = os.path.join(base_dir, "poppler", "Library", "bin")
        
        if os.path.exists(poppler_path):
            print(f"Found poppler in tool directory: {poppler_path}")
            # Thêm vào PATH tạm thời nếu chưa có
            os.environ["PATH"] = os.pathsep.join([poppler_path, os.environ.get("PATH", "")])
            return poppler_path
        
        # Cách 3: Kiểm tra các vị trí thông thường khác
        common_paths = [
            os.path.join(os.environ.get("ProgramFiles", ""), "poppler", "Library", "bin"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "poppler", "Library", "bin"),
            os.path.join(os.path.expanduser("~"), "poppler", "Library", "bin")
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                print(f"Found poppler in common path: {path}")
                os.environ["PATH"] = os.pathsep.join([path, os.environ.get("PATH", "")])
                return path
        
        print("Could not find poppler in any standard location")
        return None

    def debug_poppler_path(self):
        """Hiển thị thông tin debug về poppler path"""
        print("\n=== DEBUG POPPLER PATH ===")
        print(f"Popper path being used: {self.poppler_path}")
        if self.poppler_path:
            print("Contents of poppler bin directory:")
            try:
                print(os.listdir(self.poppler_path))
            except Exception as e:
                print(f"Error listing directory: {e}")
        else:
            print("Using system PATH for poppler")
        print(f"Poppler available: {self.poppler_available}")
        print("=========================\n")

    def check_poppler(self):
        """Kiểm tra Poppler có sẵn"""
        try:
            if self.poppler_path is None:
                # Thử không chỉ định poppler_path (hy vọng nó đã trong PATH)
                convert_from_path("test.pdf", first_page=1, last_page=1)
                return True
            else:
                # Thử với poppler_path được chỉ định
                convert_from_path("test.pdf", poppler_path=self.poppler_path, first_page=1, last_page=1)
                return True
        except PDFInfoNotInstalledError:
            print("Poppler not installed or not in PATH")
            return False
        except Exception as e:
            print(f"Error checking poppler: {e}")
            return False

    def create_widgets(self):
        # Style
        style = ttk.Style()
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10, "bold"), padding=10)  # Tăng padding cho nút
        style.configure("Large.TButton", font=("Arial", 12, "bold"), padding=15)  # Style cho nút lớn
        style.configure("TRadiobutton", background="#f0f0f0", font=("Arial", 10))
        style.configure("TEntry", padding=5)
        
        # Main Frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=15, pady=15)  # Tăng padding
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=X, pady=(0, 15))  # Tăng khoảng cách
        
        # Logo (nếu có)
        try:
            logo_img = Image.open("logo.png")
            logo_img = logo_img.resize((160, 55), Image.LANCZOS)  # Tăng kích thước logo
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = Label(header_frame, image=self.logo, bg="#f0f0f0")
            logo_label.pack(side=LEFT, padx=(0, 15))  # Tăng padding
        except:
            pass
        
        title_label = Label(header_frame, text="PDF to Image Converter", font=("Arial", 18, "bold"), bg="#f0f0f0")
        title_label.pack(side=LEFT)
        
        # PDF Selection
        pdf_frame = ttk.LabelFrame(main_frame, text="Chọn file PDF", padding=12)  # Tăng padding
        pdf_frame.pack(fill=X, pady=8)  # Tăng khoảng cách
        
        pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=50, font=("Arial", 10))
        pdf_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        browse_pdf_btn = ttk.Button(pdf_frame, text="Duyệt...", style="TButton", command=self.browse_pdf)
        browse_pdf_btn.pack(side=RIGHT)
        
        # Output Folder
        output_frame = ttk.LabelFrame(main_frame, text="Thư mục lưu ảnh", padding=12)  # Tăng padding
        output_frame.pack(fill=X, pady=8)  # Tăng khoảng cách
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_folder, width=50, font=("Arial", 10))
        output_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        browse_output_btn = ttk.Button(output_frame, text="Duyệt...", style="TButton", command=self.browse_output)
        browse_output_btn.pack(side=RIGHT)
        
        # Conversion Options
        options_frame = ttk.LabelFrame(main_frame, text="Tùy chọn chuyển đổi", padding=12)  # Tăng padding
        options_frame.pack(fill=X, pady=8)  # Tăng khoảng cách
        
        # Radio buttons for conversion type
        all_pages_radio = ttk.Radiobutton(
            options_frame, 
            text="Chuyển tất cả các trang", 
            variable=self.convert_all, 
            value=True,
            command=self.toggle_page_entry
        )
        all_pages_radio.pack(anchor=W, pady=5)  # Thêm padding
        
        select_pages_radio = ttk.Radiobutton(
            options_frame, 
            text="Chọn trang cụ thể:", 
            variable=self.convert_all, 
            value=False,
            command=self.toggle_page_entry
        )
        select_pages_radio.pack(anchor=W, pady=5)  # Thêm padding
        
        # Page selection entry with example text
        self.page_entry = ttk.Entry(
            options_frame, 
            textvariable=self.pages_to_convert, 
            width=50,
            state=DISABLED,
            font=("Arial", 10)
        )
        self.page_entry.pack(fill=X, pady=(5, 5))  # Điều chỉnh padding
        
        example_label = ttk.Label(
            options_frame, 
            text="Ví dụ: 1 (trang 1), 1,3,5 (trang 1,3,5), 1-5 (trang 1 đến 5)",
            font=("Arial", 9),
            foreground="#666666"
        )
        example_label.pack(anchor=W, pady=(0, 5))  # Điều chỉnh padding
        
        # Progress bar
        self.progress_frame = ttk.LabelFrame(main_frame, text="Tiến trình", padding=12)  # Tăng padding
        self.progress_frame.pack(fill=X, pady=10)  # Tăng khoảng cách
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=X, pady=5)  # Thêm padding
        
        self.progress_label = ttk.Label(self.progress_frame, text="Sẵn sàng", font=("Arial", 9))
        self.progress_label.pack()
        
        # Convert button - Làm nút to hơn
        self.convert_btn = ttk.Button(
            main_frame, 
            text="CHUYỂN ĐỔI", 
            command=self.start_conversion_thread,
            style="TButton"  # Sử dụng style lớn
        )
        self.convert_btn.pack(pady=15, ipadx=20, ipady=8)  # Tăng padding và kích thước nội bộ
        
        # Tạo style cho nút chính
        style.configure("Accent.TButton", foreground="white", background="#4CAF50", font=("Arial", 12, "bold"))
        style.map("Accent.TButton", 
                 background=[("active", "#45a049"), ("disabled", "#cccccc")],
                 foreground=[("disabled", "#666666")])
    
    def toggle_page_entry(self):
        if self.convert_all.get():
            self.page_entry.config(state=DISABLED)
        else:
            self.page_entry.config(state=NORMAL)
    
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            # Tự động đặt thư mục output cùng thư mục với file PDF
            output_dir = os.path.dirname(file_path)
            self.output_folder.set(output_dir)
    
    def browse_output(self):
        folder_path = filedialog.askdirectory(title="Chọn thư mục lưu ảnh")
        if folder_path:
            self.output_folder.set(folder_path)
    
    def start_conversion_thread(self):
        # Kiểm tra Poppler trước khi chuyển đổi
        if not self.poppler_available:
            self.queue.put(("error", 
                "Không tìm thấy Poppler!\n\n"
                "Vui lòng thực hiện một trong các cách sau:\n"
                "1. Cài đặt Poppler và thêm vào PATH hệ thống\n"
                "2. Đặt thư mục poppler trong cùng thư mục với tool\n"
                "   (Cấu trúc: tool_folder/poppler/Library/bin/)\n"
                "3. Chỉ định đường dẫn đến thư mục bin của Poppler\n\n"
                "Bạn có thể tải Poppler tại:\n"
                "https://github.com/oschwartz10612/poppler-windows/releases"))
            return
            
        # Kiểm tra đầu vào
        if not self.pdf_path.get():
            self.queue.put(("error", "Vui lòng chọn file PDF"))
            return
        
        if not self.output_folder.get():
            self.queue.put(("error", "Vui lòng chọn thư mục lưu ảnh"))
            return
        
        if not self.convert_all.get() and not self.pages_to_convert.get():
            self.queue.put(("error", "Vui lòng nhập trang cần chuyển đổi hoặc chọn 'Chuyển tất cả các trang'"))
            return
        
        # Vô hiệu hóa nút chuyển đổi trong quá trình chuyển đổi
        self.convert_btn.config(state=DISABLED)
        
        # Chạy trong thread
        thread = threading.Thread(target=self.convert_pdf_to_images, daemon=True)
        thread.start()
    
    def process_queue(self):
        """Xử lý các message từ queue để cập nhật GUI"""
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "progress":
                    self._update_progress_gui(msg[1], msg[2])
                elif msg[0] == "error":
                    messagebox.showerror("Lỗi", msg[1])
                    self.convert_btn.config(state=NORMAL)
                elif msg[0] == "complete":
                    messagebox.showinfo("Thành công", msg[1])
                    self.convert_btn.config(state=NORMAL)
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)
    
    def convert_pdf_to_images(self):
        try:
            pdf_path = self.pdf_path.get()
            output_folder = self.output_folder.get()
            
            # Cập nhật tiến trình ban đầu
            self.queue.put(("progress", "Đang chuẩn bị...", 0))
            
            # Lấy danh sách trang cần chuyển đổi
            if self.convert_all.get():
                pages = None
            else:
                pages = self.parse_page_numbers(self.pages_to_convert.get())
                if not pages:
                    self.queue.put(("error", "Định dạng trang không hợp lệ"))
                    return
            
            # Chuyển đổi PDF sang ảnh
            poppler_path = self.poppler_path  # Sử dụng đường dẫn đã tìm thấy
            
            # Đầu tiên, lấy tổng số trang để tính toán tiến trình chính xác
            images = convert_from_path(
                pdf_path,
                poppler_path=poppler_path,
                first_page=1,
                last_page=1,  # Chỉ đọc trang đầu để lấy tổng số trang
                fmt='jpeg'
            )
            total_pages = len(images)
            
            # Nếu chọn trang cụ thể, tính tổng số trang sẽ chuyển đổi
            if pages and not self.convert_all.get():
                total_to_convert = len(pages)
            else:
                total_to_convert = total_pages
            
            self.queue.put(("progress", f"Đang xử lý 0/{total_to_convert} trang...", 0))
            
            # Thực hiện chuyển đổi thực sự
            images = convert_from_path(
                pdf_path,
                poppler_path=poppler_path,
                first_page=pages[0] if pages else None,
                last_page=pages[-1] if pages else None,
                thread_count=4,
                fmt='jpeg',
                jpegopt={'quality': 100}
            )
            
            # Lưu từng ảnh và cập nhật tiến trình
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            for i, image in enumerate(images):
                # Chỉ lấy những trang được chọn nếu có
                if pages and not self.convert_all.get():
                    current_page = pages[i]
                    if current_page > total_pages:
                        continue
                else:
                    current_page = i + 1
                
                output_path = os.path.join(output_folder, f"{base_name}_page_{current_page}.jpg")
                image.save(output_path, 'JPEG')
                
                # Cập nhật tiến trình
                progress = ((i + 1) / total_to_convert) * 100
                self.queue.put(("progress", f"Đang xử lý trang {current_page}/{total_to_convert}...", progress))
            
            self.queue.put(("progress", f"Hoàn thành {total_to_convert} trang!", 100))
            self.queue.put(("complete", f"Đã chuyển đổi thành công {total_to_convert} trang!"))
        
        except Exception as e:
            self.queue.put(("error", f"Đã xảy ra lỗi: {str(e)}"))
            self.queue.put(("progress", f"Lỗi: {str(e)}", 0))
    
    def parse_page_numbers(self, page_str):
        """Phân tích chuỗi trang thành danh sách số trang"""
        pages = set()
        parts = page_str.split(',')
        
        for part in parts:
            part = part.strip()
            if '-' in part:
                # Trường hợp khoảng trang (ví dụ: 1-5)
                start_end = part.split('-')
                if len(start_end) != 2:
                    return None
                try:
                    start = int(start_end[0])
                    end = int(start_end[1])
                    if start < 1 or end < 1 or start > end:
                        return None
                    pages.update(range(start, end + 1))
                except ValueError:
                    return None
            else:
                # Trường hợp trang đơn (ví dụ: 1,3,5)
                try:
                    page = int(part)
                    if page < 1:
                        return None
                    pages.add(page)
                except ValueError:
                    return None
        
        return sorted(pages)
    
    def _update_progress_gui(self, message, value):
        """Cập nhật GUI (phải được gọi từ main thread)"""
        self.progress_label.config(text=message)
        self.progress_bar['value'] = value
        self.root.update_idletasks()

if __name__ == "__main__":
    root = Tk()
    app = PDFToImageConverter(root)
    root.mainloop()
