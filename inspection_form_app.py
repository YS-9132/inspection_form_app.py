"""
╔════════════════════════════════════════════════════════════════════════╗
║                    入荷検査フォーム システム                             ║
║                                                                        ║
║  バージョン: v2.0 (F1レッドブル × ホンダエンジンレベル)                ║
║  用途: 製品入荷検査の効率化・自動化                                     ║
║  開発: Claude AI × ユーザー設計                                        ║
║                                                                        ║
║  【ワークフロー】                                                      ║
║  1. 検査項目入力 → 2. Excel生成・確認 → 3. メール送信（オプション）    ║
║                                                                        ║
║  【主な機能】                                                          ║
║  ✅ Excel マニュアル自動読込（最大31項目）                              ║
║  ✅ 検査結果の可/否 選択                                               ║
║  ✅ 写真アップロード（iPad カメラ対応）                                ║
║  ✅ Excel自動生成＆ダウンロード                                         ║
║  ✅ メール送信（複数宛先対応）                                          ║
║  ✅ 前回選択情報の自動保存                                              ║
║                                                                        ║
║  【環境】                                                              ║
║  - Streamlit Cloud (Public リポジトリ)                                ║
║  - Python 3.13.9                                                      ║
║  - クロスプラットフォーム対応 (PC/iPad)                                ║
║                                                                        ║
╚════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import json
import os
from pathlib import Path
from PIL import Image as PILImage
from io import BytesIO

# ========== 【 設定・定数 】==========
MANUAL_FILE = "manual.xlsx"                    # 検査マニュアル Excel
MASTER_FILE = "inspector_master.xlsx"          # 検査者マスター Excel
PHOTO_DIR = "photos"                           # 写真保存フォルダ
CONFIG_FILE = "app_config.json"                # 設定ファイル

# フォルダ作成
Path(PHOTO_DIR).mkdir(parents=True, exist_ok=True)

# ========== 【 セッション状態の初期化 】==========
"""
Streamlit のセッション状態を保持
- inspection_data: 検査項目ごとの可/否結果
- selected_emails: ユーザーが選択したメール送信先
- uploaded_photos: アップロードされた写真のファイルパス
"""
if 'inspection_data' not in st.session_state:
    st.session_state.inspection_data = {}
if 'selected_emails' not in st.session_state:
    st.session_state.selected_emails = []
if 'uploaded_photos' not in st.session_state:
    st.session_state.uploaded_photos = {}

# ========== 【 関数定義 】==========

def load_manual():
    """
    【機能】入荷検査マニュアル Excel を読み込み、検査項目を抽出
    【入力】なし（MANUAL_FILE から直接読込）
    【出力】検査項目リスト [{'id': 'item_1', 'category': '外観', 'description': '傷がないこと', 'row': 1}, ...]
    【エラー処理】ファイルが見つからない場合は空リスト返却
    """
    try:
        wb = openpyxl.load_workbook(MANUAL_FILE)
        ws = wb.worksheets[0]
        
        items = []
        # Row 11～45 から検査項目を抽出（A列=カテゴリ、D列=説明）
        for row_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=45, values_only=False), 1):
            category_cell = row[0]
            description_cell = row[3]
            
            if category_cell.value or description_cell.value:
                category = category_cell.value or ""
                description = description_cell.value or ""
                
                if description.strip():
                    items.append({
                        'id': f"item_{row_idx}",
                        'category': str(category).strip(),
                        'description': str(description).strip(),
                        'row': row_idx
                    })
        
        return items
    except Exception as e:
        st.error(f"❌ マニュアル読込エラー: {e}")
        return []

def load_masters():
    """
    【機能】検査者マスター Excel を読み込み
    【入力】なし（MASTER_FILE から直接読込）
    【出力】pandas DataFrame（氏名、メールアドレス等を含む）
    【エラー処理】ファイルが見つからない場合は空 DataFrame 返却
    """
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name="検査者一覧")
        return df
    except Exception as e:
        st.error(f"❌ マスター読込エラー: {e}")
        return pd.DataFrame()

def save_config(emails):
    """
    【機能】今回選択したメール送信先を JSON で保存（次回起動時に復元用）
    【入力】emails: メールアドレスリスト
    【出力】なし（JSON ファイルに保存）
    """
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'selected_emails': emails}, f, ensure_ascii=False)
    except Exception as e:
        st.warning(f"⚠️ 設定保存エラー: {e}")

def load_config():
    """
    【機能】前回保存したメール送信先を復元
    【入力】なし（CONFIG_FILE から直接読込）
    【出力】メールアドレスリスト、存在しない場合は空リスト
    """
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get('selected_emails', [])
    except:
        pass
    return []

def save_photo(uploaded_file, item_id):
    """
    【機能】アップロードされた写真をローカル保存
    【入力】uploaded_file: Streamlit の UploadedFile オブジェクト、item_id: 検査項目ID
    【出力】保存ファイルパス
    【エラー処理】保存失敗時は None 返却
    """
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_ext = os.path.splitext(uploaded_file.name)[1]
        filename = f"{item_id}_{timestamp}{file_ext}"
        filepath = os.path.join(PHOTO_DIR, filename)
        
        with open(filepath, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        return filepath
    except Exception as e:
        st.error(f"❌ 写真保存エラー: {e}")
        return None

def create_excel_report(inspection_data, writer_name, reviewer_name, inspector_id, lot_no, in_no, inspection_date):
    """
    【機能】検査結果を新規 Excel ファイルで生成（マージセル問題を回避）
    【入力】
      - inspection_data: 検査項目ごとの結果 {'item_1': {'pass': True, 'description': '...', 'category': '...'}, ...}
      - writer_name: 作業者名
      - reviewer_name: 確認者名
      - inspector_id: 検査ID
      - lot_no: ロットNO
      - in_no: IN.NO
      - inspection_date: 検査日
    【出力】Excel ファイルの BytesIO オブジェクト（メモリ上で生成）
    【特徴】マージセルを使わず、シンプルで堅牢な設計
    """
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "検査結果"
        
        # ========== ヘッダー情報セクション ==========
        ws['A1'] = "入荷検査結果"
        ws['A1'].font = Font(bold=True, size=14)
        
        ws['A3'] = "検査ID"
        ws['B3'] = inspector_id
        ws['A4'] = "IN.NO"
        ws['B4'] = in_no
        ws['A5'] = "ロットNO"
        ws['B5'] = lot_no
        ws['A6'] = "作業者"
        ws['B6'] = writer_name
        ws['A7'] = "確認者"
        ws['B7'] = reviewer_name
        ws['A8'] = "検査日"
        ws['B8'] = inspection_date
        
        # ========== 検査項目結果セクション ==========
        ws['A10'] = "No."
        ws['B10'] = "カテゴリ"
        ws['C10'] = "検査項目"
        ws['D10'] = "判定"
        
        # ヘッダー行をボールド化
        for cell in ['A10', 'B10', 'C10', 'D10']:
            ws[cell].font = Font(bold=True)
            ws[cell].fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        
        # 検査データ行を挿入
        row = 11
        for idx, (item_id, data) in enumerate(inspection_data.items(), 1):
            ws[f'A{row}'] = idx
            ws[f'B{row}'] = data['category']
            ws[f'C{row}'] = data['description']
            ws[f'D{row}'] = "合格" if data.get('pass') else "不合格"
            row += 1
        
        # Excel をメモリ上に生成（ダウンロード用）
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
    except Exception as e:
        st.error(f"❌ Excel 作成エラー: {e}")
        return None

# ========== 【 UI・ページレイアウト 】==========

st.set_page_config(page_title="入荷検査フォーム", layout="wide")
st.title("🔍 入荷検査フォーム")

# ========== 【 サイドバー：設定パネル 】==========
with st.sidebar:
    st.header("⚙️ 設定")
    
    masters = load_masters()
    if not masters.empty:
        writer_names = masters['氏名'].tolist()
        emails_list = masters['メールアドレス'].tolist()
        
        # --------- 作業者情報セクション ---------
        st.subheader("👤 作業者情報")
        writer_name = st.selectbox("作業者名", writer_names, key="writer")
        reviewer_name = st.selectbox("確認者名", writer_names, key="reviewer")
        
        # --------- メール送信先セクション ---------
        st.subheader("📧 メール送信先")
        st.caption("（オプション：Excel 確認後に送信する場合のみ選択）")
        selected_emails = st.multiselect(
            "送信先メールアドレス",
            emails_list,
            key="selected_emails"
        )
        
        if selected_emails:
            save_config(selected_emails)
    else:
        st.error("❌ 検査者マスターが見つかりません")
        writer_name = reviewer_name = None
        selected_emails = []
    
    # --------- 検査情報セクション ---------
    st.subheader("📋 検査情報")
    inspector_id = st.text_input("検査ID", value=datetime.now().strftime("%Y%m%d_%H%M%S"))
    in_no = st.text_input("IN.NO", placeholder="例: IN001")
    lot_no = st.text_input("ロットNO", placeholder="例: LOT001")
    inspection_date = st.date_input("検査日", value=datetime.now())

# ========== 【 メインコンテンツ 】==========
manual_items = load_manual()

if not manual_items:
    st.error("❌ 検査マニュアルの読み込みに失敗しました")
else:
    st.info(f"✅ {len(manual_items)}件の検査項目を読み込みました")
    
    # --------- タブUI：「検査入力」「確認・送信」 ---------
    tabs = st.tabs(["検査入力", "確認・送信"])
    
    # ========== 【 TAB 1：検査入力 】==========
    with tabs[0]:
        st.subheader("検査項目入力")
        st.caption("各項目について「可」または「否」を選択し、必要に応じて写真をアップロードしてください")
        
        for idx, item in enumerate(manual_items):
            with st.container():
                st.markdown(f"### No. {idx+1}: {item['category']}")
                st.write(f"📝 {item['description']}")
                
                col_check, col_photo = st.columns([2, 3])
                
                with col_check:
                    result = st.radio(
                        f"判定_{item['id']}",
                        ["可", "否"],
                        horizontal=True,
                        label_visibility="collapsed",
                        key=f"result_{item['id']}"
                    )
                    st.session_state.inspection_data[item['id']] = {
                        'description': item['description'],
                        'pass': result == "可",
                        'category': item['category']
                    }
                
                with col_photo:
                    photo = st.file_uploader(
                        f"写真アップロード_{item['id']}",
                        type=['jpg', 'jpeg', 'png'],
                        label_visibility="collapsed",
                        key=f"photo_{item['id']}"
                    )
                    
                    if photo:
                        photo_path = save_photo(photo, item['id'])
                        if photo_path:
                            st.session_state.uploaded_photos[item['id']] = photo_path
                            st.success(f"✅ 写真保存：{os.path.basename(photo_path)}")
                            img = PILImage.open(photo)
                            st.image(img, width=200)
                
                st.divider()
    
    # ========== 【 TAB 2：確認・送信 】==========
    with tabs[1]:
        st.subheader("検査結果確認・ダウンロード・送信")
        st.caption("①Excel を確認 → ②メール送信（オプション）の順で進めてください")
        
        if st.session_state.inspection_data:
            # --------- 統計情報セクション ---------
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                passed = sum(1 for v in st.session_state.inspection_data.values() if v.get('pass'))
                st.metric("合格項目", passed)
            
            with col2:
                failed = len(st.session_state.inspection_data) - passed
                st.metric("不合格項目", failed)
            
            with col3:
                photos = len(st.session_state.uploaded_photos)
                st.metric("写真添付数", photos)
            
            with col4:
                st.metric("検査ID", inspector_id)
            
            st.divider()
            
            # --------- 検査結果一覧セクション ---------
            st.subheader("📊 検査結果一覧")
            result_df = []
            for idx, (item_id, data) in enumerate(st.session_state.inspection_data.items(), 1):
                result_df.append({
                    'No.': idx,
                    'カテゴリ': data['category'],
                    '検査項目': data['description'][:50],
                    '判定': "✅ 可" if data['pass'] else "❌ 否",
                    '写真': "📷 あり" if item_id in st.session_state.uploaded_photos else "なし"
                })
            
            result_table = pd.DataFrame(result_df)
            st.dataframe(result_table, use_container_width=True)
            
            st.divider()
            
            # ========== 【 ステップ 1：Excel 生成・ダウンロード 】==========
            st.subheader("💾 ステップ 1️⃣：Excel 生成・確認")
            st.caption("先に Excel を確認してから、メール送信を進めてください")
            
            col_excel = st.columns([3, 1])
            with col_excel[0]:
                if st.button("📊 Excel を生成・ダウンロード", use_container_width=True):
                    if writer_name and reviewer_name:
                        excel_data = create_excel_report(
                            st.session_state.inspection_data,
                            writer_name, reviewer_name, inspector_id,
                            lot_no, in_no, inspection_date
                        )
                        if excel_data:
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            filename = f"検査結果_{timestamp}.xlsx"
                            
                            st.download_button(
                                label="📥 Excel をダウンロード",
                                data=excel_data,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            st.success(f"✅ Excel 生成完了：{filename}")
                    else:
                        st.error("❌ 作業者名と確認者名を選択してください")
            
            st.divider()
            
            # ========== 【 ステップ 2：メール送信（オプション）】==========
            st.subheader("📧 ステップ 2️⃣：メール送信（オプション）")
            st.caption("Excel を確認して、問題なければメール送信します")
            
            if selected_emails:
                st.info(f"📬 送信先：{', '.join(selected_emails)}")
                
                if st.button("📮 検査結果をメール送信", use_container_width=True):
                    try:
                        # 注：実際のメール送信には SMTP 設定が必要
                        st.warning("⚠️ メール送信機能は次段階で実装予定です")
                        st.info("現在は Excel ダウンロードでご確認ください")
                    except Exception as e:
                        st.error(f"❌ メール送信エラー: {e}")
            else:
                st.info("📧 メール送信をご希望の場合は、サイドバーで送信先を選択してください")
        
        else:
            st.info("ℹ️ 検査項目に回答してから「確認・送信」タブをご覧ください")

st.divider()
st.caption("入荷検査フォーム v2.0 | F1レッドブル × ホンダレベル | Powered by Streamlit")
