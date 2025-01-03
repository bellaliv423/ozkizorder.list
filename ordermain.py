import pandas as pd
from fuzzywuzzy import fuzz
import numpy as np
import os
import sys
import glob
import re

def read_input_file(input_file):
    file_extension = os.path.splitext(input_file)[1].lower()
    if file_extension == '.csv':
        return pd.read_csv(input_file, encoding='utf-8')
    elif file_extension in ['.xlsx', '.xls']:
        return pd.read_excel(input_file)
    else:
        raise ValueError(f"지원하지 않는 파일 형식입니다: {file_extension}")

def get_product_mapping():
    """제품명 매핑 규칙 정의"""
    return {
        '원피스': ['원피스-', '-원피스'],
        '스키복': ['스키복-', '-스키복'],
        '부츠': ['부츠-', '-부츠'],
        '상의': ['상의-', '-상의'],
        '조끼': ['조끼-', '-조끼'],
        '팬츠': ['팬츠-', '-팬츠'],
        '슬립온': ['슬립온-', '-슬립온']
    }

def clean_product_name(name):
    """제품명 정제 함수"""
    # 기본 클리닝
    name = re.sub(r'\s*\([^)]*\)', '', name)  # 괄호와 그 안의 내용 제거
    name = name.replace('LED', '').strip()     # LED 제거
    
    # 불필요한 단어 제거
    remove_words = ['슈즈', '구두', '기모', '털안감', '세트', '가방세트', '잡화세트', 
                   '머리띠 세트', '패키지', '레깅스', '청바지', '점퍼', '양말', '실내화']
    
    for word in remove_words:
        name = name.replace(word, '').strip()
    
    return name.strip()

def calculate_price_35_percent(price):
    # 35% 가격 계산
    return int(price * 0.35)

def get_input_files():
    # orders 폴더의 모든 csv와 excel 파일 찾기
    csv_files = glob.glob('orders/*.csv')
    excel_files = glob.glob('orders/*.xlsx') + glob.glob('orders/*.xls')
    all_files = sorted(csv_files + excel_files)
    
    if not all_files:
        print("orders 폴더에 입력 파일이 없습니다.")
        sys.exit(1)
    
    # 파일 목록 출력
    print("\n=== 입력 파일 목록 ===")
    for idx, file in enumerate(all_files, 1):
        print(f"{idx}. {file}")
    
    # 파일 선택
    while True:
        try:
            choice = input("\n처리할 파일 번호를 입력하세요: ")
            if choice.lower() == 'q':
                print("프로그램을 종료합니다.")
                sys.exit(0)
            
            idx = int(choice) - 1
            if 0 <= idx < len(all_files):
                return all_files[idx]
            else:
                print("올바른 파일 번호를 입력하세요.")
        except ValueError:
            print("숫자를 입력하세요. (종료하려면 'q' 입력)")

def extract_core_product_name(name):
    """
    제품의 핵심 이름만 추출하는 함수
    예: "나이스 맨투맨 티셔츠" -> "나이스"
        "페세라 아트윅 기모 맨투맨 티셔츠" -> "페세라"
    """
    # 제거할 서술어 목록
    descriptive_terms = [
        '맨투맨', '티셔츠', '팬츠', '원피스', '스커트', '자켓', '코트', 
        '슬립온', '구두', '슈즈', '기모', '레깅스', '조끼', '스키복',
        '아트윅', '밴딩', '데님', '청바지', '세트', '가방', '머리띠',
        '양말', '부츠', '털안감', '코듀로이', '카고', '조거'
    ]
    
    # 이름 정제
    name = clean_product_name(name)
    words = name.split()
    
    # 단어가 없는 경우 빈 문자열 반환
    if not words:
        return ""
    
    # 첫 번째 단어를 핵심 이름으로 가정
    core_name = words[0]
    
    # 특정 제품명 매핑
    product_mapping = {
        '피넛츠': '하의-피넛츠',
        '리오': '하의-리오'
    }
    
    # 매핑된 제품명이 있으면 반환
    if core_name in product_mapping:
        return core_name
    
    # 두 번째 단어가 있고 서술어가 아닌 경우, 핵심 이름에 포함
    if len(words) > 1 and words[1].lower() not in [term.lower() for term in descriptive_terms]:
        core_name = f"{core_name} {words[1]}"
    
    return core_name.strip()

def calculate_similarity(order_name, inventory_name):
    """
    제품명 유사도 계산 함수
    - 핵심 제품 이름 매칭: 60% 이상 (필수)
    """
    def split_product_name(name):
        """제품 종류와 이름을 분리"""
        parts = name.split('-', 1)
        if len(parts) == 2:
            return parts[0].strip(), parts[1].strip()
        return '', name.strip()
    
    # 핵심 이름 추출
    order_core_name = extract_core_product_name(order_name).lower()
    _, inv_full_name = split_product_name(inventory_name)
    inv_core_name = extract_core_product_name(inv_full_name).lower()
    
    print(f"핵심 이름 비교: '{order_core_name}' vs '{inv_core_name}'")  # 디버깅용
    
    # 핵심 이름 유사도 계산
    name_similarity = fuzz.ratio(order_core_name, inv_core_name)
    
    # 특정 제품 직접 매칭
    direct_matches = {
        '나이스': '상의-나이스',
        '페세라': '상의-페세라',
        '바오': '하의-바오',
        '삐죽삐죽데님': '하의-삐죽삐죽데님',
        '폭신한맞춤': '부자재- 폭신한맞춤 깔창',
        '쿠쿠플라워': '원피스-쿠쿠플라워',
        '화이트리본': '원피스-화이트리본',
        '스노우베어': '스키복-스노우베어',
        '윈터베어': '스키복-윈터베어',
        '코니코니': '부츠-코니코니LED',
        '달콤스위티': '상의-달콤스위티',
        '피넛츠': '하의-피넛츠',
        '리오': '하의-리오'
    }
    
    # 직접 매칭 확인
    for key, value in direct_matches.items():
        if key.lower() in order_core_name and value.lower() in inventory_name.lower():
            return 100  # 완벽 매칭
    
    # 핵심 이름 유사도가 60% 이상인 경우 매칭 성공
    if name_similarity >= 60:
        return name_similarity
    
    return 0  # 매칭 실패

def get_color_mapping():
    """영문 컬러명을 한글로 매핑하는 딕셔너리 반환"""
    return {
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
        'orange': '오렌지',
        'crm': '크림',
        'wht': '화이트',
        'blk': '블랙',
        'gry': '그레이'
    }

def translate_color(color):
    """영문 컬러명을 한글로 변환"""
    color_mapping = get_color_mapping()
    color_lower = str(color).lower().strip()
    
    # 정확한 매칭 시도
    if color_lower in color_mapping:
        return color_mapping[color_lower]
    
    # 부분 매칭 시도
    for eng, kor in color_mapping.items():
        if eng in color_lower or color_lower in eng:
            return kor
    
    return color  # 매칭 실패시 원본 반환

def normalize_size(size_str):
    """사이즈 문자열 정규화"""
    # 문자열로 변환하고 모든 공백 제거
    size = str(size_str).replace(' ', '').strip()
    
    # '호' 제거
    size = size.replace('호', '')
    
    # 숫자만 추출
    numbers = re.findall(r'\d+', size)
    if numbers:
        return numbers[0]  # 첫 번째 숫자만 반환
    
    return size

def normalize_option(option_str, ignore_color=False):
    """
    옵션 문자열 정규화 함수
    ignore_color: True일 경우 컬러 매칭을 무시하고 사이즈만 비교
    """
    # 콜론 제거 및 공백 정리
    option = option_str.replace(':', '').replace(' ', '')
    
    # 옵션 부분 분리
    parts = option.split(',')
    if len(parts) >= 2:
        if ignore_color:
            # 사이즈만 반환
            size = parts[1].strip()
            return f",{size}"  # 컬러 부분은 비워두고 사이즈만 비교
        else:
            # 기존 방식대로 컬러와 사이즈 모두 반환
            color = parts[0].lower().strip()
            size = parts[1].strip()
            return f"{color},{size}"
    
    return option

def process_orders(input_file, inventory_file):
    """주문 처리 함수"""
    orders = read_input_file(input_file)
    inventory = pd.read_excel(inventory_file)
    results = []
    
    for _, order in orders.iterrows():
        best_match = None
        best_score = 0
        
        order_product = str(order['Product']).strip()
        order_color = str(order['Color']).strip()
        order_size = str(order['Size']).strip()
        
        print(f"\n처리 중인 주문: {order_product}")
        print(f"옵션: {order_color}, {order_size}")
        
        # 제품명 매칭
        for _, inv in inventory.iterrows():
            inv_name = str(inv['상품명']).strip()
            similarity = calculate_similarity(order_product, inv_name)
            
            if similarity >= 60:  # 60% 이상 유사도
                print(f"제품명 매칭 (유사도 {similarity:.1f}%): {inv_name}")
                
                # 옵션 매칭 시도 (먼저 컬러 포함해서 시도)
                inv_option = str(inv['옵션'])
                expected_option = f"{order_color}, :{order_size}"
                
                # 1차 시도: 컬러와 사이즈 모두 매칭
                inv_option_norm = normalize_option(inv_option)
                expected_option_norm = normalize_option(expected_option)
                
                match_found = inv_option_norm == expected_option_norm
                
                # 컬러 매칭 실패시, 해당 제품의 컬러가 1가지인지 확인
                if not match_found:
                    # 현재 제품의 모든 옵션 가져오기
                    product_options = inventory[inventory['상품명'] == inv_name]['옵션'].unique()
                    unique_colors = set()
                    
                    # 해당 제품의 모든 컬러 수집
                    for opt in product_options:
                        color = opt.split(',')[0].strip() if ',' in opt else ''
                        unique_colors.add(color)
                    
                    # 컬러가 1가지만 있는 경우
                    if len(unique_colors) == 1:
                        # 사이즈만으로 매칭 시도
                        inv_option_norm = normalize_option(inv_option, ignore_color=True)
                        expected_option_norm = normalize_option(expected_option, ignore_color=True)
                        match_found = inv_option_norm == expected_option_norm
                
                if match_found:
                    if similarity > best_score:
                        best_score = similarity
                        best_match = inv
                        print(f"매칭 성공! 상품코드: {inv['상품코드']}")
                        print(f"매칭된 옵션: {inv_option}")
        
        # 결과 저장
        if best_match is not None:
            results.append({
                'Order_Product': order_product,
                'Order_Color': order_color,
                'Order_Size': order_size,
                'Order_Quantity': order['Quantity'],
                'Order_Price_35': calculate_price_35_percent(best_match['판매가']),
                'Matched_Name': best_match['상품명'],
                'Matched_Code': best_match['상품코드'],
                'Matched_Price': best_match['판매가'],
                'Matched_stocks': best_match['가용재고'],
                'Matched_Option': best_match['옵션'],
                'Similarity': best_score
            })
        else:
            results.append({
                'Order_Product': order_product,
                'Order_Color': order_color,
                'Order_Size': order_size,
                'Order_Quantity': order['Quantity'],
                'Order_Price_35': 0,
                'Matched_Name': '매칭 실패',
                'Matched_Code': '매칭 실패',
                'Matched_Price': 0,
                'Matched_stocks': 0,
                'Matched_Option': '',
                'Similarity': 0
            })
    
    return pd.DataFrame(results)

def match_product(product_name, size, color=None):
    # 기존 코드...
    
    # 컬러 정보가 없는 경우 사이즈로만 매칭
    if not color:
        matching_products = [p for p in products 
                           if p['product_name'] == product_name 
                           and p['size'] == size]
        if matching_products:
            return matching_products[0]
    
    # 컬러 정보가 있는 경우 기존 로직대로 처리
    matching_products = [p for p in products 
                        if p['product_name'] == product_name 
                        and p['size'] == size 
                        and p['color'] == color]
    
    return matching_products[0] if matching_products else None

def preprocess_data(df):
    """
    데이터 전처리 함수
    """
    # 공백 제거
    df['Order_Product'] = df['Order_Product'].str.strip()
    df['Order_Color'] = df['Order_Color'].str.strip()
    
    # 특수문자 정규화
    df['Order_Product'] = df['Order_Product'].apply(normalize_special_chars)
    
    return df

if __name__ == "__main__":
    # 명령줄 인자로 파일을 지정한 경우
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        if not os.path.exists(input_file):
            print(f"파일을 찾을 수 없습니다: {input_file}")
            sys.exit(1)
    else:
        # 파일 선택 메뉴 표시
        input_file = get_input_files()
    
    inventory_file = 'database/현재고조회.xlsx'
    
    # 입력 파일명에서 확장자를 제외한 이름 추출
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    
    # output 폴더 생성
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # 주문 처리
    print(f"\n'{input_file}' 파일 처리 중...")
    results = process_orders(input_file, inventory_file)
    
    # 결과 출력
    print("\n=== 매칭 결과 ===")
    for idx, row in results.iterrows():
        print(f"\n[주문 {idx + 1}]")
        print(f"주문상품: {row['Order_Product']}")
        print(f"주문옵션: {row['Order_Color']}, {row['Order_Size']}")
        print(f"주문수량: {row['Order_Quantity']}개")
        print(f"35% 가격: {row['Order_Price_35']:,}원")
        print(f"매칭상품: {row['Matched_Name']}")
        print(f"상품코드: {row['Matched_Code']}")
        print(f"판매가격: {row['Matched_Price']:,}원")
        print(f"가용재고: {row['Matched_stocks']}개")
        print(f"매칭옵션: {row['Matched_Option']}")
        print("-" * 50)
    
    # 결과 파일 저장 (입력 파일명 기준으로 출력 파일명 생성)
    output_csv = f'output/{base_name}_results.csv'
    output_excel = f'output/{base_name}_results.xlsx'
    
    # 열 순서 지정
    column_order = [
        'Order_Product', 'Order_Color', 'Order_Size', 'Order_Quantity', 'Order_Price_35',
        'Matched_Name', 'Matched_Code', 'Matched_Price', 'Matched_stocks', 'Matched_Option', 'Similarity'
    ]
    results = results[column_order]
    
    # UTF-8 with BOM으로 저장하여 한글이 깨지지 않도록 함
    results.to_csv(output_csv, index=False, encoding='utf-8-sig')
    results.to_excel(output_excel, index=False)
    
    print(f"\n결과가 저장되었습니다:")
    print(f"CSV 파일: {output_csv}")
    print(f"Excel 파일: {output_excel}") 