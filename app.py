import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

# --- ★★★変更点：Excelの書式設定に必要なライブラリをインポート★★★ ---
from openpyxl.styles import Font, PatternFill

# --------------------------------------------------------------------------------
# ★★★ここからが追加・復活したExcel整形外科医（format_excel_sheet関数）★★★
# --------------------------------------------------------------------------------
def format_excel_sheet(ws, df):
    """Excelシートの書式（列幅、数値フォーマットなど）を整える"""
    # 1. 列幅の自動調整
    for col_idx, column_cells in enumerate(ws.columns, 1):
        column_letter = column_cells[0].column_letter
        max_length = 0
        # セルの値の最大文字数を探す
        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        # ヘッダー（列名）の文字数も考慮
        header_text = ws.cell(row=1, column=col_idx).value
        if header_text:
            max_length = max(max_length, len(str(header_text)))
        # 少し余裕を持たせて列幅を設定
        ws.column_dimensions[column_letter].width = max_length + 3

    # 2. 数値の書式設定
    header = [c.value for c in ws[1]]
    red_font = Font(color="FF0000")
    
    # 3桁区切りとマイナス赤字を適用したい列
    cols_to_format = ["在庫数", "基準数量(自動)", "基準数量(手動)", "差し引き数量", "経過日数"]
    
    for col_name in cols_to_format:
        if col_name in header:
            col_idx = header.index(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'
                    if col_name == "差し引き数量" and cell.value < 0:
                        cell.font = red_font
    return ws

# --------------------------------------------------------------------------------
# 分析ロジック（変更なし）
# --------------------------------------------------------------------------------
def find_column_name(df_columns, possible_names):
    return next((name for name in possible_names if name in df_columns), None)

def analyze_inventory(src_file, rule_file, history_file):
    # (この analyze_inventory 関数の中身は、前回のコードと全く同じです)
    ws_key = pd.read_excel(rule_file, sheet_name="キー", header=None, dtype=str).fillna("")
    key_dict = {str(val).strip(): str(ws_key.iloc[0, col_idx]).strip() for col_idx in range(ws_key.shape[1]) for val in ws_key.iloc[1:, col_idx] if str(val).strip()}
    manual_quantities = {}
    if "リスト" in pd.ExcelFile(rule_file).sheet_names:
        df_list = pd.read_excel(rule_file, sheet_name="リスト", dtype=str)
        quantity_col_name = find_column_name(df_list.columns, ["基準数量（手動）", "数量"])
        if quantity_col_name:
            manual_quantities = df_list.set_index('商品名')[quantity_col_name].apply(pd.to_numeric, errors='coerce').dropna().astype('Int64').to_dict()
    df_history = pd.DataFrame()
    if history_file is not None:
        try:
            df_history_raw = pd.read_excel(history_file, sheet_name='Data', engine='xlrd')
            column_rename_map = {"订单发行日": "order_date", "注文発行日": "order_date", "订单数量": "order_quantity", "注文数量": "order_quantity", "商品名称": "product_name"}
            df_history = df_history_raw.rename(columns=lambda c: column_rename_map.get(c, c))
            df_history["order_date"] = pd.to_datetime(df_history["order_date"], errors='coerce')
            df_history["order_quantity"] = pd.to_numeric(df_history["order_quantity"], errors='coerce')
            df_history.dropna(subset=["product_name", "order_date", "order_quantity"], inplace=True)
        except Exception as e:
            st.warning(f"発注履歴ファイルの読み込み中にエラーが発生しました: {e}")
    ws_src = pd.read_excel(src_file, header=10, dtype=str).fillna("")
    inventory_col = find_column_name(ws_src.columns, ['客户在库', '在库数量', '在庫数量'])
    if not inventory_col:
        st.error("エラー: 元在庫表に在庫数量を示す列が見つかりません。")
        return None, None, None, None
    cols_to_map = {'商品名称': '商品名称', inventory_col: 'INVENTORY_LEVEL', '最终出荷日': '最终出荷日'}
    ws_src.rename(columns={k: v for k, v in cols_to_map.items() if k in ws_src.columns}, inplace=True)
    current_inventory_map = {str(row["商品名称"]).strip(): int(pd.to_numeric(row["INVENTORY_LEVEL"], errors='coerce')) for _, row in ws_src.iterrows() if pd.notna(row["商品名称"]) and pd.notna(row["INVENTORY_LEVEL"])}
    consumption_dict = {}
    if not df_history.empty:
        df_agg = df_history.groupby('product_name').agg(total_incoming=('order_quantity', 'sum'), first_order_date=('order_date', 'min')).reset_index()
        today = datetime.now()
        for _, row in df_agg.iterrows():
            p_name, total_in, first_date = row['product_name'], row['total_incoming'], row['first_order_date']
            months = max(1, (today.year - first_date.year) * 12 + (today.month - first_date.month))
            current_stock = current_inventory_map.get(p_name, 0)
            total_consumption = total_in - current_stock
            if total_consumption > 0:
                monthly_con = round(total_consumption / months)
                if monthly_con > 0: consumption_dict[p_name] = int(monthly_con)
    low_stock_auto, low_stock_manual, long_term_stock = [], [], []
    one_year_ago = datetime.now() - relativedelta(years=1)
    def assign_brand(product_name):
        return next((bname for key, bname in key_dict.items() if key in str(product_name)), "OTHER")
    ws_src['ブランド'] = ws_src['商品名称'].apply(assign_brand)
    df_src_for_report = ws_src.drop_duplicates(subset=['商品名称'])
    for _, row in df_src_for_report.iterrows():
        p_name = str(row['商品名称']).strip()
        if not p_name: continue
        brand_name = row['ブランド']
        inv_qty, auto_qty, manual_qty = current_inventory_map.get(p_name, 0), consumption_dict.get(p_name), manual_quantities.get(p_name, 0)
        if auto_qty and inv_qty < auto_qty: low_stock_auto.append({"ブランド": brand_name, "商品名": p_name, "在庫数": inv_qty, "基準数量(自動)": auto_qty, "差し引き数量": inv_qty - auto_qty})
        if manual_qty and inv_qty < manual_qty: low_stock_manual.append({"ブランド": brand_name, "商品名": p_name, "在庫数": inv_qty, "基準数量(手動)": manual_qty, "差し引き数量": inv_qty - manual_qty})
        ship_date = pd.to_datetime(str(row["最终出荷日"]).strip(), errors='coerce')
        if pd.notna(ship_date) and ship_date < one_year_ago: long_term_stock.append({"ブランド": brand_name, "商品名": p_name, "最終出荷日": ship_date.date(), "経過日数": (datetime.now() - ship_date).days, "在庫数": inv_qty})
    df_auto = pd.DataFrame(low_stock_auto); df_manual = pd.DataFrame(low_stock_manual); df_long = pd.DataFrame(long_term_stock)
    return ws_src, df_auto, df_manual, df_long

# --- ★★★変更点：Excelダウンロード用の関数を強化★★★ ---
def to_excel(df_full, df_auto, df_manual, df_long):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # サマリーシートの書き込みと書式設定
        df_auto.to_excel(writer, sheet_name='不足在庫_自動', index=False)
        format_excel_sheet(writer.sheets['不足在庫_自動'], df_auto)
        
        df_manual.to_excel(writer, sheet_name='不足在庫_手動', index=False)
        format_excel_sheet(writer.sheets['不足在庫_手動'], df_manual)
        
        df_long.to_excel(writer, sheet_name='長期在庫', index=False)
        format_excel_sheet(writer.sheets['長期在庫'], df_long)
        
        # ブランドごとのシートの書き込みと書式設定
        cols_to_drop = ['INVENTORY_LEVEL']
        df_to_write = df_full.drop(columns=[col for col in cols_to_drop if col in df_full.columns])
        brands = sorted(df_to_write['ブランド'].unique())
        for brand in brands:
            brand_df = df_to_write[df_to_write['ブランド'] == brand].drop(columns=['ブランド'])
            brand_df.to_excel(writer, sheet_name=brand, index=False)
            format_excel_sheet(writer.sheets[brand], brand_df) # 各ブランドシートにも書式設定を適用

    return output.getvalue()

# --------------------------------------------------------------------------------
# Streamlit UI部分（変更なし）
# --------------------------------------------------------------------------------
st.set_page_config(layout="wide")
st.title('📈 在庫分析ダッシュボード')

BASE_PATH = Path(__file__).resolve().parent
DEFAULT_RULE_FILE = BASE_PATH / "振り分けルール.xlsx"
DEFAULT_HISTORY_FILE = BASE_PATH / "発注履歴.xls"

if not DEFAULT_RULE_FILE.exists():
    st.error("エラー: アプリケーションに「振り分けルール.xlsx」が同梱されていません。")
    st.stop()
rule_file = DEFAULT_RULE_FILE
history_file = DEFAULT_HISTORY_FILE if DEFAULT_HISTORY_FILE.exists() else None

st.info("👇 分析したい「元在庫表」をアップロードしてください。")
uploaded_src_file = st.file_uploader("", type=['xlsx', 'xls'], label_visibility="collapsed")

st.sidebar.header("⚙️ 設定")
st.sidebar.markdown("""
このダッシュボードは、同梱されているマスターファイルを使用します。
- **振り分けルール:** `振り分けルール.xlsx`
- **発注履歴:** `発注履歴.xls` (存在する場合)
""")
with st.sidebar.expander("もし、特別なファイルで試したい場合はこちら"):
    uploaded_rule_override = st.file_uploader("特別な「振り分けルール」", type=['xlsx', 'xls'])
    uploaded_history_override = st.file_uploader("特別な「発注履歴」", type=['xlsx', 'xls'])
    if uploaded_rule_override: rule_file = uploaded_rule_override
    if uploaded_history_override: history_file = uploaded_history_override

if uploaded_src_file:
    st.success(f"「{uploaded_src_file.name}」を読み込みました。")
    with st.spinner('在庫データを分析中...'):
        df_full, df_auto, df_manual, df_long = analyze_inventory(uploaded_src_file, rule_file, history_file)
    
    if df_full is not None:
        st.success('分析が完了しました！')
        st.header('分析結果')
        
        excel_data = to_excel(df_full, df_auto, df_manual, df_long)
        st.download_button(
            label="📄 見やすいExcel形式で全データをダウンロード",
            data=excel_data,
            file_name=f"在庫レポート_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        tab1, tab2, tab3 = st.tabs([f"不足在庫(自動) ({len(df_auto)})", f"不足在庫(手動) ({len(df_manual)})", f"長期在庫 ({len(df_long)})"])
        with tab1: st.dataframe(df_auto)
        with tab2: st.dataframe(df_manual)
        with tab3: st.dataframe(df_long.sort_values(by="経過日数", ascending=False) if not df_long.empty else df_long)

        st.divider()
        st.header('全在庫リスト（ブランド別詳細）')
        brand_list = ["全ブランド表示"] + sorted(df_full['ブランド'].unique())
        selected_brand = st.selectbox('表示したいブランドを選択してください:', brand_list)
        if selected_brand == "全ブランド表示":
            st.dataframe(df_full)
        else:
            st.dataframe(df_full[df_full['ブランド'] == selected_brand])
            st.dataframe(df_full)
        else:

            st.dataframe(df_full[df_full['ブランド'] == selected_brand])
