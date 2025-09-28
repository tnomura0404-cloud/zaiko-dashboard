import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
from openpyxl.styles import Font, PatternFill

# --------------------------------------------------------------------------------
# Excel整形関数（元のコードからハイライト機能などを移植）
# --------------------------------------------------------------------------------
def format_excel_sheet(ws, df):
    is_shortage_report = ws.title.startswith("不足在庫")
    is_long_term_report = ws.title == "長期在庫"
    
    for col_idx, column_cells in enumerate(ws.columns, 1):
        column_letter = column_cells[0].column_letter
        # ★★★改善点：特定レポートの列幅を固定★★★
        if (is_shortage_report and column_letter in ['C', 'D', 'E', 'F']) or \
           (is_long_term_report and column_letter in ['C', 'D', 'E', 'F']):
            ws.column_dimensions[column_letter].width = 15
        else:
            max_length = 0
            for cell in column_cells:
                if cell.value is not None: max_length = max(max_length, len(str(cell.value)))
            header_text = ws.cell(row=1, column=col_idx).value
            if header_text: max_length = max(max_length, len(str(header_text)))
            ws.column_dimensions[column_letter].width = max_length + 3

    header = [c.value for c in ws[1]]
    red_font = Font(color="FF0000")
    cols_to_format = ["在庫数", "基準数量(自動)", "基準数量(手動)", "差し引き数量", "経過日数"]
    money_columns = ["贩卖单价", "合計金額"]
    
    for col_name in cols_to_format:
        if col_name in header:
            col_idx = header.index(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'
                    if col_name == "差し引き数量" and cell.value < 0:
                        cell.font = red_font
    
    for col_name in money_columns:
        if col_name in header:
            col_idx = header.index(col_name) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
    return ws

# --------------------------------------------------------------------------------
# 分析ロジック（is_low_stockなどの判定結果を返すように変更）
# --------------------------------------------------------------------------------
def find_column_name(df_columns, possible_names):
    return next((name for name in possible_names if name in df_columns), None)

def analyze_inventory(src_file, rule_file, history_file):
    # (ルール, 履歴, 在庫表の読み込み部分は変更なし)
    ws_key = pd.read_excel(rule_file, sheet_name="キー", header=None, dtype=str).fillna("")
    key_dict = {str(val).strip(): str(ws_key.iloc[0, col_idx]).strip() for col_idx in range(ws_key.shape[1]) for val in ws_key.iloc[1:, col_idx] if str(val).strip()}
    manual_quantities = {}
    if "リスト" in pd.ExcelFile(rule_file).sheet_names:
        df_list = pd.read_excel(rule_file, sheet_name="リスト", dtype=str)
        quantity_col_name = find_column_name(df_list.columns, ["基準数量（手動）", "数量"])
        if quantity_col_name: manual_quantities = df_list.set_index('商品名')[quantity_col_name].apply(pd.to_numeric, errors='coerce').dropna().astype('Int64').to_dict()
    df_history = pd.DataFrame()
    if history_file is not None:
        try:
            df_history_raw = pd.read_excel(history_file, sheet_name='Data', engine='xlrd')
            column_rename_map = {"订单发行日": "order_date", "注文発行日": "order_date", "订单数量": "order_quantity", "注文数量": "order_quantity", "商品名称": "product_name"}
            df_history = df_history_raw.rename(columns=lambda c: column_rename_map.get(c, c))
            df_history["order_date"] = pd.to_datetime(df_history["order_date"], errors='coerce')
            df_history["order_quantity"] = pd.to_numeric(df_history["order_quantity"], errors='coerce')
            df_history.dropna(subset=["product_name", "order_date", "order_quantity"], inplace=True)
        except Exception as e: st.warning(f"発注履歴ファイルの読み込み中にエラーが発生しました: {e}")
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
            if total_consumption > 0: monthly_con = round(total_consumption / months);
            if monthly_con > 0: consumption_dict[p_name] = int(monthly_con)

    # ★★★改善点：ハイライト用の判定結果をws_srcに直接追加★★★
    def assign_brand(product_name): return next((bname for key, bname in key_dict.items() if key in str(product_name)), "OTHER")
    ws_src['ブランド'] = ws_src['商品名称'].apply(assign_brand)
    one_year_ago = datetime.now() - relativedelta(years=1)
    
    low_stock_auto_products = set()
    low_stock_manual_products = set()
    long_term_products = set()

    for _, row in ws_src.iterrows():
        p_name = str(row["商品名称"]).strip()
        if not p_name: continue
        inv_qty = current_inventory_map.get(p_name, 0)
        auto_qty = consumption_dict.get(p_name)
        manual_qty = manual_quantities.get(p_name)
        if auto_qty and inv_qty < auto_qty: low_stock_auto_products.add(p_name)
        if manual_qty and inv_qty < manual_qty: low_stock_manual_products.add(p_name)
        ship_date = pd.to_datetime(str(row["最终出荷日"]).strip(), errors='coerce')
        if pd.notna(ship_date) and ship_date < one_year_ago: long_term_products.add(p_name)
    
    # ws_srcに判定結果をブール値で追加
    ws_src['is_low_stock'] = ws_src['商品名称'].isin(low_stock_auto_products | low_stock_manual_products)
    ws_src['is_long_term'] = ws_src['商品名称'].isin(long_term_products)

    # レポート用データフレームの作成
    df_src_for_report = ws_src.drop_duplicates(subset=['商品名称'])
    report_auto = df_src_for_report[df_src_for_report['商品名称'].isin(low_stock_auto_products)].copy()
    report_manual = df_src_for_report[df_src_for_report['商品名称'].isin(low_stock_manual_products)].copy()
    report_long = df_src_for_report[df_src_for_report['商品名称'].isin(long_term_products)].copy()

    return ws_src, report_auto, report_manual, report_long

# --- Excelダウンロード用の関数 ---
def to_excel(df_full, df_auto, df_manual, df_long):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # レポートシート
        df_auto.to_excel(writer, sheet_name='不足在庫_自動', index=False)
        format_excel_sheet(writer.sheets['不足在庫_自動'], df_auto)
        df_manual.to_excel(writer, sheet_name='不足在庫_手動', index=False)
        format_excel_sheet(writer.sheets['不足在庫_手動'], df_manual)
        df_long.to_excel(writer, sheet_name='長期在庫', index=False)
        format_excel_sheet(writer.sheets['長期在庫'], df_long)
        
        # ブランド別シート
        brands = sorted(df_full['ブランド'].unique())
        for brand in brands:
            brand_df = df_full[df_full['ブランド'] == brand].copy()
            # is_low_stock, is_long_term 列はこのExcelには不要なので削除
            brand_df_cleaned = brand_df.drop(columns=['ブランド', 'is_low_stock', 'is_long_term'], errors='ignore')
            brand_df_cleaned.to_excel(writer, sheet_name=brand, index=False)
            ws = writer.sheets[brand]
            format_excel_sheet(ws, brand_df_cleaned)

            # ★★★改善点：ハイライト処理をここで行う★★★
            low_fill = PatternFill(fill_type="solid", fgColor="FFFF00")  # 黄色
            long_fill = PatternFill(fill_type="solid", fgColor="FFCCCC")  # 赤色
            header = [c.value for c in ws[1]]
            try:
                p_idx = header.index("商品名称") + 1
                s_idx = header.index("最终出荷日") + 1
                # 元の brand_df (判定結果を持つ) を使ってループ
                for r_idx, row in enumerate(brand_df.itertuples(), 2):
                    if row.is_low_stock:
                        ws.cell(row=r_idx, column=p_idx).fill = low_fill
                    if row.is_long_term:
                        ws.cell(row=r_idx, column=s_idx).fill = long_fill
            except (ValueError, AttributeError):
                pass # 列がない場合は何もしない

    return output.getvalue()

# --------------------------------------------------------------------------------
# Streamlit UI部分
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
        with tab3: st.dataframe(df_long)

        st.divider()
        st.header('全在庫リスト（ブランド別詳細）')
        brand_list = ["全ブランド表示"] + sorted(df_full['ブランド'].unique())
        selected_brand = st.selectbox('表示したいブランドを選択してください:', brand_list)
        df_display = df_full.drop(columns=['is_low_stock', 'is_long_term'], errors='ignore')
        if selected_brand == "全ブランド表示":
            st.dataframe(df_display)
        else:
            st.dataframe(df_display[df_display['ブランド'] == selected_brand])
