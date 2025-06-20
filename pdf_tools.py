import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import PyPDF2
import os
from pathlib import Path
from PIL import Image, ImageTk
import pdf2image
import io

class PDFManager:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 문서 관리자")
        self.root.geometry("1200x700")
        self.root.configure(bg='#f0f0f0')
        
        # 파일 리스트 저장
        self.pdf_files = []
        self.current_preview_file = None
        self.current_page = 0
        self.total_pages = 0
        self.preview_images = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # 메인 컨테이너
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 좌측 프레임 (기능 탭들)
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))
        
        # 우측 프레임 (미리보기)
        right_frame = ttk.LabelFrame(main_container, text="PDF 미리보기", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 좌측 프레임 설정
        self.setup_left_panel(left_frame)
        
        # 우측 미리보기 설정
        self.setup_preview_panel(right_frame)
        
    def setup_left_panel(self, parent):
        # 탭 생성
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 병합 탭
        merge_frame = ttk.Frame(notebook, padding="10")
        notebook.add(merge_frame, text="PDF 병합")
        self.setup_merge_tab(merge_frame)
        
        # 분리 탭
        split_frame = ttk.Frame(notebook, padding="10")
        notebook.add(split_frame, text="PDF 분리")
        self.setup_split_tab(split_frame)
        
        # 회전 탭
        rotate_frame = ttk.Frame(notebook, padding="10")
        notebook.add(rotate_frame, text="PDF 회전")
        self.setup_rotate_tab(rotate_frame)
        
        # 좌측 패널 너비 고정
        parent.configure(width=500)
        
    def setup_preview_panel(self, parent):
        # 미리보기 컨트롤 프레임
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 파일 선택 버튼
        ttk.Button(control_frame, text="미리보기할 파일 선택", 
                  command=self.select_preview_file).pack(side=tk.LEFT, padx=(0, 10))
        
        # 현재 파일 표시
        self.preview_file_label = ttk.Label(control_frame, text="파일이 선택되지 않았습니다.")
        self.preview_file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # 미리보기 설정 도움말 버튼 추가
        ttk.Button(control_frame, text="미리보기 설정 도움말", 
                  command=self.show_preview_setup_help).pack(side=tk.RIGHT)
        
        # 페이지 네비게이션 프레임
        nav_frame = ttk.Frame(parent)
        nav_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(nav_frame, text="◀ 이전", command=self.prev_page).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(nav_frame, text="다음 ▶", command=self.next_page).pack(side=tk.LEFT, padx=(0, 10))
        
        self.page_label = ttk.Label(nav_frame, text="페이지: 0 / 0")
        self.page_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # 줌 컨트롤
        ttk.Label(nav_frame, text="줌:").pack(side=tk.LEFT, padx=(10, 5))
        self.zoom_var = tk.StringVar(value="100")
        zoom_combo = ttk.Combobox(nav_frame, textvariable=self.zoom_var, 
                                 values=["50", "75", "100", "125", "150", "200"], 
                                 width=8, state="readonly")
        zoom_combo.pack(side=tk.LEFT, padx=(0, 5))
        zoom_combo.bind('<<ComboboxSelected>>', self.on_zoom_change)
        
        ttk.Button(nav_frame, text="새로고침", command=self.refresh_preview).pack(side=tk.LEFT, padx=(10, 0))
        
        # 미리보기 이미지 프레임 (스크롤 가능)
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # 캔버스와 스크롤바
        self.preview_canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.preview_canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.preview_canvas.xview)
        
        self.preview_canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # 그리드 배치
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)
        
        # 마우스 휠 스크롤 바인딩
        self.preview_canvas.bind("<MouseWheel>", self.on_mousewheel)
        
    def setup_merge_tab(self, parent):
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(parent, text="PDF 파일 선택", padding="10")
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 파일 리스트박스
        listbox_frame = ttk.Frame(file_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.merge_listbox = tk.Listbox(listbox_frame, height=8, selectmode=tk.EXTENDED)
        self.merge_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.merge_listbox.bind('<<ListboxSelect>>', self.on_merge_file_select)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.merge_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.merge_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="파일 추가", command=self.add_files_to_merge).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="파일 제거", command=self.remove_files_from_merge).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="위로", command=self.move_up).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="아래로", command=self.move_down).pack(side=tk.LEFT, padx=(0, 5))
        
        # 병합 실행 프레임
        merge_exec_frame = ttk.LabelFrame(parent, text="병합 실행", padding="10")
        merge_exec_frame.pack(fill=tk.X)
        
        ttk.Button(merge_exec_frame, text="PDF 병합하기", command=self.merge_pdfs).pack()
        
    def setup_split_tab(self, parent):
        # 파일 선택 프레임
        split_file_frame = ttk.LabelFrame(parent, text="분리할 PDF 파일 선택", padding="10")
        split_file_frame.pack(fill=tk.X, pady=(0, 10))
        
        file_select_frame = ttk.Frame(split_file_frame)
        file_select_frame.pack(fill=tk.X)
        
        self.split_file_var = tk.StringVar()
        entry = ttk.Entry(file_select_frame, textvariable=self.split_file_var, state="readonly")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(file_select_frame, text="파일 선택", command=self.select_file_to_split).pack(side=tk.RIGHT)
        
        # 분리 옵션 프레임
        split_options_frame = ttk.LabelFrame(parent, text="분리 옵션", padding="10")
        split_options_frame.pack(fill=tk.BOTH, expand=True)
        
        # 분리 방식 선택
        self.split_method = tk.StringVar(value="pages")
        ttk.Radiobutton(split_options_frame, text="페이지별 분리", variable=self.split_method, 
                       value="pages").pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(split_options_frame, text="범위별 분리", variable=self.split_method, 
                       value="range").pack(anchor=tk.W, pady=2)
        
        # 범위 입력
        ttk.Label(split_options_frame, text="페이지 범위 (예: 1-5, 10-15):").pack(anchor=tk.W, pady=(10, 2))
        self.page_range_var = tk.StringVar()
        ttk.Entry(split_options_frame, textvariable=self.page_range_var).pack(fill=tk.X, pady=(0, 10))
        
        # 분리 실행
        ttk.Button(split_options_frame, text="PDF 분리하기", command=self.split_pdf).pack()
        
    def setup_rotate_tab(self, parent):
        # 파일 선택 프레임
        rotate_file_frame = ttk.LabelFrame(parent, text="회전할 PDF 파일 선택", padding="10")
        rotate_file_frame.pack(fill=tk.X, pady=(0, 10))
        
        file_select_frame = ttk.Frame(rotate_file_frame)
        file_select_frame.pack(fill=tk.X)
        
        self.rotate_file_var = tk.StringVar()
        entry = ttk.Entry(file_select_frame, textvariable=self.rotate_file_var, state="readonly")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(file_select_frame, text="파일 선택", command=self.select_file_to_rotate).pack(side=tk.RIGHT)
        
        # 회전 옵션 프레임
        rotate_options_frame = ttk.LabelFrame(parent, text="회전 옵션", padding="10")
        rotate_options_frame.pack(fill=tk.BOTH, expand=True)
        
        # 회전 각도
        ttk.Label(rotate_options_frame, text="회전 각도:").pack(anchor=tk.W, pady=(0, 5))
        self.rotation_angle = tk.StringVar(value="90")
        angle_frame = ttk.Frame(rotate_options_frame)
        angle_frame.pack(anchor=tk.W, pady=(0, 10))
        
        ttk.Radiobutton(angle_frame, text="90°", variable=self.rotation_angle, value="90").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(angle_frame, text="180°", variable=self.rotation_angle, value="180").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(angle_frame, text="270°", variable=self.rotation_angle, value="270").pack(side=tk.LEFT, padx=(0, 10))
        
        # 페이지 범위
        ttk.Label(rotate_options_frame, text="회전할 페이지 (예: 1-3, 5, 7-9, 전체는 공백):").pack(anchor=tk.W, pady=(0, 5))
        self.rotate_pages_var = tk.StringVar()
        ttk.Entry(rotate_options_frame, textvariable=self.rotate_pages_var).pack(fill=tk.X, pady=(0, 10))
        
        # 회전 실행
        ttk.Button(rotate_options_frame, text="PDF 회전하기", command=self.rotate_pdf).pack()
        
    def select_preview_file(self):
        file = filedialog.askopenfilename(
            title="미리보기할 PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.load_preview(file)
            
    def load_preview(self, file_path):
        try:
            self.preview_file_label.config(text="로딩 중...")
            self.root.update()

            try:
                import pdf2image
            except ImportError:
                raise Exception("pdf2image 라이브러리가 설치되지 않았습니다.\n'pip install pdf2image' 명령으로 설치해주세요.")

            poppler_dir = Path(__file__).parent / "poppler" / "Library" / "bin"
            poppler_path = str(poppler_dir) if poppler_dir.exists() else None

            try:
                if poppler_path:
                    pages = pdf2image.convert_from_path(file_path, dpi=150, poppler_path=poppler_path)
                else:
                    pages = pdf2image.convert_from_path(file_path, dpi=150)
            except Exception as pdf_error:
                if "poppler" in str(pdf_error).lower():
                    raise Exception(
                        "Poppler가 설치되지 않았습니다.\n\nWindows 사용자:\n"
                        "1. 'conda install -c conda-forge poppler'\n"
                        "2. https://github.com/oschwartz10612/poppler-windows 에서 다운로드 후 "
                        "'./poppler/Library/bin' 경로에 위치시키세요.\n"
                        "3. 또는 시스템 PATH에 poppler의 bin 경로를 추가하세요."
                    )
                else:
                    raise pdf_error

            self.preview_images = [page for page in pages]
            self.current_preview_file = file_path
            self.total_pages = len(self.preview_images)
            self.current_page = 0

            # 🔽 줌을 100%로 맞춰 캔버스 폭 기준 자동 확대되게 함
            self.zoom_var.set("100")

            self.preview_file_label.config(text=f"파일: {os.path.basename(file_path)}")
            self.update_page_display()

        except Exception as e:
            error_msg = str(e)
            if "poppler" in error_msg.lower() or "pdf2image" in error_msg.lower():
                self.show_preview_setup_help()
            else:
                messagebox.showerror("오류", f"PDF 미리보기 로드 중 오류가 발생했습니다:\n{error_msg}")
            self.preview_file_label.config(text="미리보기 사용 불가")

                
    def update_page_display(self):
        if not self.preview_images or self.current_page >= len(self.preview_images):
            return

        try:
            # 현재 페이지의 PIL 이미지
            pil_image = self.preview_images[self.current_page]

            # 현재 캔버스 너비 측정 (기본값은 800)
            self.preview_canvas.update_idletasks()  # 실제 크기 반영
            canvas_width = self.preview_canvas.winfo_width()
            if canvas_width <= 1:
                canvas_width = 800  # fallback 기본값

            # 줌 비율 처리
            zoom_str = self.zoom_var.get()
            if zoom_str == "100":
                # 캔버스 폭 기준으로 비율 자동 조절
                scale_ratio = canvas_width / pil_image.width
            else:
                # 사용자가 선택한 줌 비율 적용
                scale_ratio = int(zoom_str) / 100.0

            # 크기 계산 및 이미지 리사이징
            new_width = int(pil_image.width * scale_ratio)
            new_height = int(pil_image.height * scale_ratio)
            resized_image = pil_image.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # tkinter 이미지 객체 생성 및 표시
            self.current_photo = ImageTk.PhotoImage(resized_image)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.current_photo)

            # 스크롤 범위 설정
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

            # 페이지 라벨 업데이트
            self.page_label.config(text=f"페이지: {self.current_page + 1} / {self.total_pages}")

        except Exception as e:
            messagebox.showerror("오류", f"페이지 표시 중 오류가 발생했습니다:\n{str(e)}")

            
    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()
            
    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_page_display()
            
    def on_zoom_change(self, event=None):
        self.update_page_display()
        
    def refresh_preview(self):
        if self.current_preview_file:
            self.load_preview(self.current_preview_file)
            
    def on_mousewheel(self, event):
        self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def on_merge_file_select(self, event):
        selection = self.merge_listbox.curselection()
        if selection and self.pdf_files:
            selected_index = selection[0]
            if selected_index < len(self.pdf_files):
                selected_file = self.pdf_files[selected_index]
                try:
                    self.load_preview(selected_file)
                except Exception as e:
                    # 미리보기가 실패해도 선택은 유지
                    print(f"미리보기 로드 실패: {e}")
                    self.preview_file_label.config(text=f"미리보기 불가: {os.path.basename(selected_file)}")
        
    def add_files_to_merge(self):
        files = filedialog.askopenfilenames(
            title="병합할 PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.merge_listbox.insert(tk.END, os.path.basename(file))
                
    def remove_files_from_merge(self):
        selected_indices = self.merge_listbox.curselection()
        for i in reversed(selected_indices):
            self.merge_listbox.delete(i)
            del self.pdf_files[i]
            
    def move_up(self):
        selected = self.merge_listbox.curselection()
        if selected and selected[0] > 0:
            idx = selected[0]
            # 리스트박스에서 이동
            item = self.merge_listbox.get(idx)
            self.merge_listbox.delete(idx)
            self.merge_listbox.insert(idx-1, item)
            self.merge_listbox.selection_set(idx-1)
            
            # 파일 리스트에서도 이동
            self.pdf_files[idx], self.pdf_files[idx-1] = self.pdf_files[idx-1], self.pdf_files[idx]
            
    def move_down(self):
        selected = self.merge_listbox.curselection()
        if selected and selected[0] < self.merge_listbox.size() - 1:
            idx = selected[0]
            # 리스트박스에서 이동
            item = self.merge_listbox.get(idx)
            self.merge_listbox.delete(idx)
            self.merge_listbox.insert(idx+1, item)
            self.merge_listbox.selection_set(idx+1)
            
            # 파일 리스트에서도 이동
            self.pdf_files[idx], self.pdf_files[idx+1] = self.pdf_files[idx+1], self.pdf_files[idx]
            
    def select_file_to_split(self):
        file = filedialog.askopenfilename(
            title="분리할 PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.split_file_var.set(file)
            try:
                self.load_preview(file)
            except Exception as e:
                print(f"미리보기 로드 실패: {e}")
            
    def select_file_to_rotate(self):
        file = filedialog.askopenfilename(
            title="회전할 PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file:
            self.rotate_file_var.set(file)
            try:
                self.load_preview(file)
            except Exception as e:
                print(f"미리보기 로드 실패: {e}")
                
    def show_preview_setup_help(self):
        help_msg = """PDF 미리보기 기능을 사용하려면 추가 설정이 필요합니다.

필요한 라이브러리:
 1. pdf 파일의 처리를 위해 Poppler 가 필요하며, 내장되어 있습니다.
 2. 해당파일이 누락되면 미리보기가 되지 않습니다. 
 
(누락시)Windows 사용자 - Poppler 설치:
방법 1: conda install -c conda-forge poppler
방법 2: https://github.com/oschwartz10612/poppler-windows
        에서 다운로드 후 PATH에 추가

미리보기 없이도 PDF 병합/분리/회전 기능은 정상 작동합니다.
2025.6.20,  Make by. Chris"""
        
        messagebox.showinfo("미리보기 설정 안내", help_msg)
            
    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showwarning("경고", "병합할 PDF 파일을 선택해주세요.")
            return
            
        output_file = filedialog.asksaveasfilename(
            title="병합된 PDF 저장",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if not output_file:
            return
            
        try:
            merger = PyPDF2.PdfMerger()
            
            for file_path in self.pdf_files:
                merger.append(file_path)
                
            with open(output_file, 'wb') as output:
                merger.write(output)
                
            merger.close()
            messagebox.showinfo("완료", f"PDF 병합이 완료되었습니다.\n저장 위치: {output_file}")
            
        except Exception as e:
            messagebox.showerror("오류", f"PDF 병합 중 오류가 발생했습니다:\n{str(e)}")
            
    def split_pdf(self):
        if not self.split_file_var.get():
            messagebox.showwarning("경고", "분리할 PDF 파일을 선택해주세요.")
            return
            
        input_file = self.split_file_var.get()
        
        try:
            with open(input_file, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                
                if self.split_method.get() == "pages":
                    # 페이지별 분리
                    output_dir = filedialog.askdirectory(title="분리된 파일들을 저장할 폴더 선택")
                    if not output_dir:
                        return
                        
                    base_name = Path(input_file).stem
                    
                    for i, page in enumerate(reader.pages):
                        writer = PyPDF2.PdfWriter()
                        writer.add_page(page)
                        
                        output_path = os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf")
                        with open(output_path, 'wb') as output_file:
                            writer.write(output_file)
                            
                    messagebox.showinfo("완료", f"PDF가 {total_pages}개 파일로 분리되었습니다.\n저장 위치: {output_dir}")
                    
                else:
                    # 범위별 분리
                    page_ranges = self.page_range_var.get().strip()
                    if not page_ranges:
                        messagebox.showwarning("경고", "페이지 범위를 입력해주세요.")
                        return
                        
                    output_file = filedialog.asksaveasfilename(
                        title="분리된 PDF 저장",
                        defaultextension=".pdf",
                        filetypes=[("PDF files", "*.pdf")]
                    )
                    
                    if not output_file:
                        return
                        
                    writer = PyPDF2.PdfWriter()
                    page_numbers = self.parse_page_ranges(page_ranges, total_pages)
                    
                    for page_num in page_numbers:
                        writer.add_page(reader.pages[page_num - 1])
                        
                    with open(output_file, 'wb') as output:
                        writer.write(output)
                        
                    messagebox.showinfo("완료", f"지정된 페이지가 분리되었습니다.\n저장 위치: {output_file}")
                    
        except Exception as e:
            messagebox.showerror("오류", f"PDF 분리 중 오류가 발생했습니다:\n{str(e)}")
            
    def rotate_pdf(self):
        if not self.rotate_file_var.get():
            messagebox.showwarning("경고", "회전할 PDF 파일을 선택해주세요.")
            return
            
        input_file = self.rotate_file_var.get()
        angle = int(self.rotation_angle.get())
        
        output_file = filedialog.asksaveasfilename(
            title="회전된 PDF 저장",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if not output_file:
            return
            
        try:
            with open(input_file, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                writer = PyPDF2.PdfWriter()
                
                total_pages = len(reader.pages)
                pages_to_rotate = self.rotate_pages_var.get().strip()
                
                if not pages_to_rotate:
                    # 전체 페이지 회전
                    page_numbers = list(range(1, total_pages + 1))
                else:
                    # 지정된 페이지만 회전
                    page_numbers = self.parse_page_ranges(pages_to_rotate, total_pages)
                
                for i, page in enumerate(reader.pages):
                    if (i + 1) in page_numbers:
                        rotated_page = page.rotate(angle)
                        writer.add_page(rotated_page)
                    else:
                        writer.add_page(page)
                        
                with open(output_file, 'wb') as output:
                    writer.write(output)
                    
                messagebox.showinfo("완료", f"PDF 회전이 완료되었습니다.\n저장 위치: {output_file}")
                
        except Exception as e:
            messagebox.showerror("오류", f"PDF 회전 중 오류가 발생했습니다:\n{str(e)}")
            
    def parse_page_ranges(self, ranges_str, total_pages):
        """페이지 범위 문자열을 파싱하여 페이지 번호 리스트 반환"""
        page_numbers = set()
        
        for range_part in ranges_str.split(','):
            range_part = range_part.strip()
            
            if '-' in range_part:
                # 범위 (예: 1-5)
                start, end = map(int, range_part.split('-'))
                start = max(1, start)
                end = min(total_pages, end)
                page_numbers.update(range(start, end + 1))
            else:
                # 단일 페이지 (예: 3)
                page_num = int(range_part)
                if 1 <= page_num <= total_pages:
                    page_numbers.add(page_num)
                    
        return sorted(list(page_numbers))

def main():
    root = tk.Tk()
    app = PDFManager(root)
    root.mainloop()

if __name__ == "__main__":
    main()