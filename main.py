import pandas as pd
from fuzzywuzzy import fuzz
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import re
from datetime import datetime
import os

class OZKIZOrderSystem:
    def __init__(self):
        # 기본 디렉토리 설정
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir = os.path.join(self.base_dir, 'output')
        self.orders_dir = os.path.join(self.base_dir, 'orders')
        self.database_dir = os.path.join(self.base_dir, 'database')
        
        # 출력 디렉토리 생성
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.orders_dir, exist_ok=True)
        os.makedirs(self.database_dir, exist_ok=True)
        
        # 컬러 매핑 딕셔너리 추가
        self.color_mapping = {
            'cream': '크림',
            'ivory': '아이보리',
            'pink': '핑크',
            'blue': '블루',
            'white': '화이트',
            'black': '블랙',
            'gray': '그레이',
            'red': '레드',
            'yellow': '옐로우',
            'green': '그린',
            'purple': '퍼플',
            'brown': '브라운',
            'navy': '네이비',
            'beige': '베이지',
            'orange': '오렌지'
        }
        
        # GUI 초기화
        self.window = tk.Tk()
        self.window.title("OZKIZ 발주 시스템")
        self.window.geometry("1200x800")
        
        # 데이터베이스 로드
        self.load_database()
        
        # GUI 구성
        self.create_gui()

    def load_database(self):
        try:
            # 데이터베이스 파일 경로
            db_path = os.path.join(self.base_dir, 'database', 'inventory.xlsx')
            print(f"Loading database from: {db_path}")  # 디버깅용
            
            # 파일이 없으면 빈 DataFrame 생성
            if not os.path.exists(db_path):
                print("Database file not found, creating empty DataFrame")
                self.inventory_df = pd.DataFrame(columns=[
                    'product_code', 'product_name', 'option',
                    'price', 'origin', 'available_stock'
                ])
                return
            
            # Excel 파일 읽기
            self.inventory_df = pd.read_excel(db_path, engine='openpyxl')
            print(f"Loaded {len(self.inventory_df)} rows")  # 디버깅용
            
            # 컬럼명 설정
            self.inventory_df.columns = [
                'product_code',      # 상품코드
                'product_name',      # 상품명
                'option',            # 옵션
                'price',             # 판매가
                'origin',            # 원산지
                'available_stock'    # 가용재고
            ]
            print("Database loaded successfully")  # 디버깅용
            
        except Exception as e:
            print(f"Error loading database: {str(e)}")  # 디버깅용
            self.inventory_df = pd.DataFrame(columns=[
                'product_code', 'product_name', 'option',
                'price', 'origin', 'available_stock'
            ])

    def create_gui(self):
        # 메뉴바 생성
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="주문 파일 열기", command=self.load_order_file)
        file_menu.add_command(label="발주서 저장 위치 열기", command=self.open_output_folder)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.window.quit)
        
        # 주문 입력 프레임
        input_frame = ttk.LabelFrame(self.window, text="주문 입력", padding="10")
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.order_text = scrolledtext.ScrolledText(input_frame, height=10)
        self.order_text.pack(fill=tk.X)
        
        # 버튼 프레임
        button_frame = ttk.Frame(self.window, padding="5")
        button_frame.pack(fill=tk.X, padx=10)
        
        ttk.Button(button_frame, text="주문 파일 열기", 
                  command=self.load_order_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="주문 처리", 
                  command=self.process_orders).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="초기화", 
                  command=self.clear_all).pack(side=tk.LEFT, padx=5)
        
        # 결과 표시 프레임
        result_frame = ttk.LabelFrame(self.window, text="처리 결과", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(result_frame)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        # 상태바
        self.status_var = tk.StringVar()
        self.status_var.set("준비됨")
        status_bar = ttk.Label(self.window, textvariable=self.status_var, 
                             relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=10)

    def load_order_file(self):
        """주문 파일 로드"""
        file_path = filedialog.askopenfilename(
            initialdir=self.orders_dir,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    self.order_text.delete(1.0, tk.END)
                    self.order_text.insert(tk.END, file.read())
                self.status_var.set(f"파일 로드됨: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("오류", f"파일 로드 실패: {str(e)}")

    def parse_order(self, order_text):
        """주문 텍스트 파싱"""
        print(f"\n주문 파싱: {order_text}")  # 디버깅용
        
        try:
            product_info = order_text.strip()
            size_color_pairs = []
            
            # 수량 정보가 있는 부분들 찾기 (예: "100 (1)")
            quantity_matches = re.finditer(r'(\d+)\s*\((\d+)\)', product_info)
            last_end = 0
            current_color = None
            current_product = None
            
            # 첫 번째 숫자나 괄호가 나오기 전까지의 텍스트를 제품명으로 처리
            first_number = re.search(r'\d+\s*\(', product_info)
            if first_number:
                current_product = product_info[:first_number.start()].strip()
                words = current_product.split()
                if words:
                    current_color = words[-1]  # 마지막 단어를 색상으로 처리
                    current_product = ' '.join(words[:-1])  # 나머지를 제품명으로
            
            for match in quantity_matches:
                size = match.group(1)
                quantity = int(match.group(2))
                
                size_color_pairs.append({
                    'product_name': current_product,
                    'color': current_color,
                    'size': size,
                    'quantity': quantity
                })
                
                print(f"파싱 결과: 제품명: {current_product}, 색상: {current_color}, "
                      f"사이즈: {size}, 수량: {quantity}")  # 디버깅용
            
            return {
                'product_name': current_product,
                'variants': size_color_pairs
            }
            
        except Exception as e:
            print(f"주문 파싱 오류: {str(e)}")  # 디버깅용
            return {'product_name': '', 'variants': []}

    def translate_color(self, color):
        """영문 컬러명을 한글로 변환"""
        color_lower = str(color).lower().strip()
        return self.color_mapping.get(color_lower, color)

    def find_matching_product(self, search_info):
        """재고 DB에서 일치하는 제품 찾기"""
        print(f"검색 정보: {search_info}")  # 디버깅용
        best_match = None
        best_score = 0
        
        # 컬러 한글 변환
        search_color = self.translate_color(search_info['color'])
        print(f"검색 컬러 변환: {search_info['color']} -> {search_color}")  # 디버깅용
        
        for _, row in self.inventory_df.iterrows():
            # 제품명 정규화
            product_name_clean = str(search_info['product_name']).strip().lower().replace(" ", "")
            inventory_name_clean = str(row['product_name']).strip().lower().replace(" ", "")
            
            # 퍼지 매칭 점수 계산
            name_score = fuzz.ratio(product_name_clean, inventory_name_clean)
            
            print(f"비교: '{product_name_clean}' vs '{inventory_name_clean}' = {name_score}")  # 디버깅용
            
            if name_score > 60:  # 임계값을 60%로 낮춤
                # 옵션 문자열 파싱
                option_str = str(row['option']).strip()
                
                # 옵션 문자열에서 색상과 사이즈 추출
                # 예: "크림, :120" 형식 파싱
                option_parts = option_str.split(',')
                if len(option_parts) >= 2:
                    db_color = option_parts[0].strip()
                    db_size = option_parts[1].strip().replace(':', '').strip()
                    
                    # 색상 매칭 (한글 변환된 컬러와 비교)
                    color_match = (search_color.lower() == db_color.lower())
                    # 사이즈 매칭
                    size_match = (str(search_info['size']).strip() == db_size)
                    
                    print(f"옵션 비교: DB({db_color}, {db_size}) vs Search({search_color}, {search_info['size']})")
                    
                    if color_match and size_match:
                        if name_score > best_score:
                            best_score = name_score
                            best_match = row
                            print(f"매칭 발견! 점수: {name_score}")  # 디버깅용
        
        if best_match is None:
            print(f"매칭 실패: {search_info}")  # 디버깅용
        else:
            print(f"최종 매칭: {best_match['product_name']}")  # 디버깅용
        
        return best_match

    def process_orders(self):
        """주문 처리 메인 함수"""
        self.status_var.set("주문 처리 중...")
        orders = self.order_text.get(1.0, tk.END).strip().split('\n')
        
        if not orders or orders[0] == '':
            messagebox.showwarning("경고", "처리할 주문이 없습니다.")
            self.status_var.set("준비됨")
            return
        
        results = []
        current_product = None
        
        for order in orders:
            if order.strip():
                order_info = self.parse_order(order)
                print(f"\nProcessing order: {order}")
                
                if not order_info['variants']:
                    current_product = order_info['product_name']
                    print(f"New product group: {current_product}")
                    continue
                
                for variant in order_info['variants']:
                    search_info = {
                        'product_name': current_product or order_info['product_name'],
                        'color': variant['color'],
                        'size': variant['size']
                    }
                    
                    matching_product = self.find_matching_product(search_info)
                    
                    if matching_product is not None:
                        results.append({
                            'product_code': matching_product['product_code'],
                            'product_name': matching_product['product_name'],
                            'color': variant['color'],
                            'size': variant['size'],
                            'quantity': variant['quantity']
                        })
        
        if results:
            df = pd.DataFrame(results)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"어드민_발주서_{timestamp}.xlsx"
            output_path = os.path.join(self.output_dir, filename)
            df.to_excel(output_path, index=False, columns=[
                'product_code', 'product_name', 'color', 'size', 'quantity'
            ])
            self.show_results(results, filename)
            self.status_var.set(f"발주서 생성 완료: {filename}")
        else:
            messagebox.showwarning("경고", "매칭되는 제품을 찾을 수 없습니다.")
            self.status_var.set("준비됨")

    def show_results(self, results, filename):
        """처리 결과를 화면에 표시"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, f"발주서 생성 완료: {filename}\n\n")
        
        for result in results:
            self.result_text.insert(tk.END,
                f"SKU: {result['product_code']}\n"
                f"제품명: {result['product_name']}\n"
                f"색상: {result['color']}\n"
                f"사이즈: {result['size']}\n"
                f"수량: {result['quantity']}\n"
                f"----------------------------------------\n"
            )

    def clear_all(self):
        """입력창과 결과창 초기화"""
        self.order_text.delete(1.0, tk.END)
        self.result_text.delete(1.0, tk.END)
        self.status_var.set("준비됨")

    def open_output_folder(self):
        """결과 폴더 열기"""
        os.startfile(self.output_dir)

    def run(self):
        """프로그램 실행"""
        self.window.mainloop()

def match_product(order_product, stock_df):
    """
    제품 매칭 함수
    
    Args:
        order_product (str): 주문 상품명
        stock_df (DataFrame): 현재고조회 데이터프레임
    
    Returns:
        tuple: (매칭된 상품명, 유사도)
    """
    # 1. 정확한 매칭 시도
    exact_match = stock_df[stock_df['상품명'].str.contains(order_product, case=False, na=False)]
    if not exact_match.empty:
        return exact_match.iloc[0]['상품명'], 100.0
    
    # 2. 유사도 매칭
    max_similarity = 0
    best_match = None
    
    for stock_name in stock_df['상품명']:
        similarity = calculate_similarity(order_product, stock_name)
        if similarity >= 60 and similarity > max_similarity:
            max_similarity = similarity
            best_match = stock_name
            
    return (best_match, max_similarity) if best_match else (None, 0)

def process_order(order_row, stock_df):
    """
    주문 처리 함수
    
    Args:
        order_row (Series): 주문 정보
        stock_df (DataFrame): 현재고조회 데이터프레임
    
    Returns:
        dict: 매칭 결과
    """
    product_name = order_row['Order_Product']
    color = order_row['Order_Color']
    size = order_row['Order_Size']
    
    # 제품 매칭
    matched_name, similarity = match_product(product_name, stock_df)
    
    if matched_name:
        # 옵션 매칭 (색상, 사이즈)
        stock_item = stock_df[
            (stock_df['상품명'] == matched_name) & 
            (stock_df['옵션'].str.contains(f":{color}", na=False)) &
            (stock_df['옵션'].str.contains(f":{size}", na=False))
        ]
        
        if not stock_item.empty:
            return {
                'Matched_Name': matched_name,
                'Matched_Code': stock_item.iloc[0]['상품코드'],
                'Similarity': similarity,
                'Matched_Option': f":{color}, :{size}"
            }
            
    return {
        'Matched_Name': '매칭 실패',
        'Matched_Code': '매칭 실패',
        'Similarity': 0.0,
        'Matched_Option': ''
    }

if __name__ == "__main__":
    app = OZKIZOrderSystem()
    app.run()