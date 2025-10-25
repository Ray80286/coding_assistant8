
import pandas as pd
import math
from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Excel 파일 경로
EXCEL_FILE = 'c:/Users/USER/workspaces/coding_assistant7/stock7.xlsx'

def get_data():
    """Excel 파일에서 데이터를 읽어 DataFrame으로 반환합니다."""
    try:
        # 데이터 불러올 때, 종목코드를 문자열로 불러오도록 dtype 지정
        df = pd.read_excel(EXCEL_FILE, dtype={'종목코드': str})
    except FileNotFoundError:
        df = pd.DataFrame(columns=['종목코드', '회사명', '현재가', '거래량', '예측'])
        df.to_excel(EXCEL_FILE, index=False)
    return df

def save_data(df):
    """DataFrame을 Excel 파일로 저장합니다."""
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    """메인 페이지, 주식 목록을 검색, 정렬, 페이지네이션 기능과 함께 보여줍니다."""
    df = get_data()

    # 검색
    search_query = request.args.get('search', '')
    if search_query:
        df = df[df.apply(lambda row: search_query.lower() in str(row['회사명']).lower() or 
                                     search_query.lower() in str(row['종목코드']).lower(), axis=1)]

    # 정렬
    sort_by = request.args.get('sort_by', '회사명')
    order = request.args.get('order', 'asc')
    if sort_by in df.columns:
        ascending = (order == 'asc')
        df = df.sort_values(by=sort_by, ascending=ascending)

    # 페이지네이션
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 10, type=int)  # 페이지 당 항목 수
    total_items = len(df)
    total_pages = math.ceil(total_items / per_page)
    start = (page - 1) * per_page
    end = start + per_page
    df_page = df.iloc[start:end]

    # 다음 정렬 순서를 결정하기 위한 딕셔너리
    next_order = 'desc' if order == 'asc' else 'asc'

    return render_template('index.html', 
                           stocks=df_page.to_dict('records'),
                           page=page,
                           total_pages=total_pages,
                           search_query=search_query,
                           sort_by=sort_by,
                           order=order,
                           next_order=next_order,
                           per_page=per_page)

@app.route('/add')
def add_stock_form():
    """신규 종목 추가 폼을 보여줍니다."""
    return render_template('add.html')

@app.route('/add_stock', methods=['POST'])
def add_stock():
    """폼에서 받은 데이터로 신규 종목을 추가합니다."""
    df = get_data()
    
    # 폼 데이터 가져오기
    new_stock = {
        '종목코드': request.form['code'],
        '회사명': request.form['name'],
        '현재가': request.form['price'],
        '거래량': request.form['volume'],
        '예측': request.form['prediction']
    }
    
    # 새 데이터를 DataFrame에 추가
    new_df = pd.DataFrame([new_stock])
    df = pd.concat([df, new_df], ignore_index=True)
    
    # Excel 파일에 저장
    save_data(df)
    
    return redirect(url_for('index'))

@app.route('/edit/<stock_code>')
def edit_stock_form(stock_code):
    """기존 종목을 수정하는 폼을 보여줍니다."""
    df = get_data()
    # 종목코드로 해당 데이터 찾기
    stock = df[df['종목코드'] == stock_code].to_dict('records')
    if not stock:
        return "종목을 찾을 수 없습니다.", 404
    return render_template('edit.html', stock=stock[0])

@app.route('/update/<stock_code>', methods=['POST'])
def update_stock(stock_code):
    """폼에서 받은 데이터로 기존 종목을 수정합니다."""
    df = get_data()
    
    # 수정할 데이터의 인덱스 찾기
    stock_index = df[df['종목코드'] == stock_code].index
    
    if not stock_index.empty:
        # 폼 데이터로 해당 행 업데이트 (종목코드는 변경 불가)
        df.loc[stock_index, '회사명'] = request.form['name']
        df.loc[stock_index, '현재가'] = int(request.form['price'])
        df.loc[stock_index, '거래량'] = int(request.form['volume'])
        df.loc[stock_index, '예측'] = request.form['prediction']
        
        save_data(df)
        
    return redirect(url_for('index'))

@app.route('/delete/<stock_code>', methods=['POST'])
def delete_stock(stock_code):
    """해당 종목을 삭제합니다."""
    df = get_data()
    
    # 종목코드가 일치하지 않는 행만 남기기
    df = df[df['종목코드'] != stock_code]
    
    save_data(df)
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=8000)

