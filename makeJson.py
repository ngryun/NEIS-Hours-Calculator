import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import json
import re

class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to JSON 변환기")
        self.root.geometry("800x600")
        
        # 메인 프레임
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # 입력 라벨과 버튼 프레임
        input_label_frame = ttk.Frame(main_frame)
        input_label_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        input_label_frame.columnconfigure(0, weight=1)
        
        ttk.Label(input_label_frame, text="엑셀 데이터 붙여넣기:", font=("Arial", 12, "bold")).grid(
            row=0, column=0, sticky=tk.W
        )
        
        # 붙여넣기 버튼
        ttk.Button(
            input_label_frame, 
            text="붙여넣기 (Ctrl+V)", 
            command=self.paste_from_clipboard
        ).grid(row=0, column=1, sticky=tk.E)
        
        # 입력 텍스트 영역
        self.input_text = scrolledtext.ScrolledText(
            main_frame, 
            height=10, 
            width=80,
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.input_text.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 키보드 단축키 바인딩
        self.input_text.bind('<Control-v>', self.on_paste)
        self.input_text.bind('<Control-V>', self.on_paste)
        self.input_text.bind('<Control-a>', self.on_select_all)
        self.input_text.bind('<Control-A>', self.on_select_all)
        
        # 우클릭 컨텍스트 메뉴 설정
        self.setup_context_menu()
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # 변환 버튼
        ttk.Button(
            button_frame, 
            text="JSON으로 변환", 
            command=self.convert_to_json,
            style="Accent.TButton"
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # 지우기 버튼
        ttk.Button(
            button_frame, 
            text="모두 지우기", 
            command=self.clear_all
        ).pack(side=tk.LEFT)
        
        # 출력 라벨
        ttk.Label(main_frame, text="JSON 결과:", font=("Arial", 12, "bold")).grid(
            row=4, column=0, columnspan=2, sticky=tk.W, pady=(10, 5)
        )
        
        # 출력 텍스트 영역
        self.output_text = scrolledtext.ScrolledText(
            main_frame, 
            height=12, 
            width=80,
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED
        )
        self.output_text.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 복사 버튼
        self.copy_button = ttk.Button(
            main_frame, 
            text="클립보드에 복사", 
            command=self.copy_to_clipboard,
            state=tk.DISABLED,
            style="Accent.TButton"
        )
        self.copy_button.grid(row=6, column=0, columnspan=2, pady=10)
        
        # 예시 데이터 추가
        example_text = """Key Value 국어 국어 화법과 작문 국어 독서 국어 언어와 매체 국어 문학 국어 실용 국어 국어 심화 국어 국어"""
        self.input_text.insert("1.0", example_text)
    
    def setup_context_menu(self):
        """우클릭 컨텍스트 메뉴 설정"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="잘라내기 (Ctrl+X)", command=self.cut_text)
        self.context_menu.add_command(label="복사 (Ctrl+C)", command=self.copy_text)
        self.context_menu.add_command(label="붙여넣기 (Ctrl+V)", command=self.paste_from_clipboard)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="모두 선택 (Ctrl+A)", command=self.select_all_text)
        
        # 우클릭 이벤트 바인딩
        self.input_text.bind("<Button-3>", self.show_context_menu)
    
    def show_context_menu(self, event):
        """컨텍스트 메뉴 표시"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def on_paste(self, event):
        """Ctrl+V 키 이벤트 처리"""
        self.paste_from_clipboard()
        return "break"  # 기본 이벤트 차단
    
    def on_select_all(self, event):
        """Ctrl+A 키 이벤트 처리"""
        self.select_all_text()
        return "break"  # 기본 이벤트 차단
    
    def paste_from_clipboard(self):
        """클립보드에서 텍스트 붙여넣기"""
        try:
            clipboard_content = self.root.clipboard_get()
            # 현재 커서 위치에 텍스트 삽입
            current_pos = self.input_text.index(tk.INSERT)
            self.input_text.insert(current_pos, clipboard_content)
            messagebox.showinfo("붙여넣기 완료", "클립보드의 내용이 붙여넣기되었습니다.")
        except tk.TclError:
            messagebox.showwarning("경고", "클립보드에 텍스트가 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"붙여넣기 중 오류가 발생했습니다: {str(e)}")
    
    def cut_text(self):
        """선택된 텍스트 잘라내기"""
        try:
            if self.input_text.selection_get():
                self.input_text.event_generate("<<Cut>>")
        except tk.TclError:
            pass  # 선택된 텍스트가 없는 경우
    
    def copy_text(self):
        """선택된 텍스트 복사"""
        try:
            if self.input_text.selection_get():
                self.input_text.event_generate("<<Copy>>")
        except tk.TclError:
            pass  # 선택된 텍스트가 없는 경우
    
    def select_all_text(self):
        """모든 텍스트 선택"""
        self.input_text.tag_add(tk.SEL, "1.0", tk.END)
        self.input_text.mark_set(tk.INSERT, "1.0")
        self.input_text.see(tk.INSERT)
    
    def parse_excel_data(self, text):
        """엑셀 데이터를 파싱하여 딕셔너리로 변환"""
        # 줄바꿈과 탭으로 분리
        lines = text.strip().split('\n')
        result = {}
        
        for line in lines:
            # 탭 또는 여러 공백으로 분리
            parts = re.split(r'\t+|\s{2,}', line.strip())
            parts = [part.strip() for part in parts if part.strip()]
            
            # 첫 번째 줄이 헤더인 경우 건너뛰기
            if len(parts) >= 2 and parts[0].lower() in ['key', 'keys']:
                continue
            
            # Key-Value 쌍 처리
            if len(parts) >= 2:
                key = parts[0]
                value = parts[1]
                result[key] = value
            elif len(parts) == 1:
                # 단일 값인 경우, 이전 패턴을 따라 처리
                continue
        
        # 만약 탭/공백 분리가 제대로 안된 경우, 연속된 한국어 단어들을 분리 시도
        if not result:
            # 모든 텍스트를 하나의 문자열로 합치고 한국어 단어 단위로 분리
            all_text = ' '.join(lines)
            words = re.findall(r'[가-힣]+(?:\s+[가-힣]+)*', all_text)
            
            # 짝수 개의 단어가 있다면 Key-Value 쌍으로 처리
            if len(words) % 2 == 0:
                for i in range(0, len(words), 2):
                    if i + 1 < len(words):
                        result[words[i].strip()] = words[i + 1].strip()
        
        return result
    
    def convert_to_json(self):
        """입력된 데이터를 JSON으로 변환"""
        input_data = self.input_text.get("1.0", tk.END).strip()
        
        if not input_data:
            messagebox.showwarning("경고", "변환할 데이터를 입력해주세요.")
            return
        
        try:
            # 데이터 파싱
            parsed_data = self.parse_excel_data(input_data)
            
            if not parsed_data:
                messagebox.showerror("오류", "데이터를 파싱할 수 없습니다. 형식을 확인해주세요.")
                return
            
            # JSON 형식으로 변환 (한국어 지원을 위해 ensure_ascii=False)
            json_output = json.dumps(parsed_data, indent=4, ensure_ascii=False)
            
            # 출력 텍스트 영역에 표시
            self.output_text.config(state=tk.NORMAL)
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", json_output)
            self.output_text.config(state=tk.DISABLED)
            
            # 복사 버튼 활성화
            self.copy_button.config(state=tk.NORMAL)
            
            messagebox.showinfo("성공", f"변환 완료! {len(parsed_data)}개의 항목이 변환되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다: {str(e)}")
    
    def copy_to_clipboard(self):
        """JSON 결과를 클립보드에 복사"""
        try:
            json_text = self.output_text.get("1.0", tk.END).strip()
            if json_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(json_text)
                self.root.update()
                messagebox.showinfo("복사 완료", "JSON 데이터가 클립보드에 복사되었습니다.")
            else:
                messagebox.showwarning("경고", "복사할 데이터가 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"클립보드 복사 중 오류가 발생했습니다: {str(e)}")
    
    def clear_all(self):
        """모든 텍스트 영역 지우기"""
        self.input_text.delete("1.0", tk.END)
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete("1.0", tk.END)
        self.output_text.config(state=tk.DISABLED)
        self.copy_button.config(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = ExcelToJsonConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()