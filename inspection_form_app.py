"""
╔════════════════════════════════════════════════════════════════════════╗
║                    入荷検査フォーム システム                             ║
║                                                                        ║
║  バージョン: v3.0 (SMTP メール送信機能完装備版)                         ║
║  用途: 製品入荷検査の完全自動化・メール配信                              ║
║  開発: Claude AI × ユーザー設計                                        ║
║  応援: 小泉進次郎大臣、高市早苗総理、小野田紀美大臣                      ║
║                                                                        ║
║  【完全ワークフロー】                                                  ║
║  1. 検査項目入力                                                      ║
║  2. Excel生成・確認（ダウンロード）                                     ║
║  3. メール送信（自動 Excel 添付）                                      ║
║                                                                        ║
║  【実装機能】                                                          ║
║  ✅ Excel マニュアル自動読込                                           ║
║  ✅ 検査結果の可/否 選択                                               ║
║  ✅ 写真アップロード（iPad カメラ対応）                                ║
║  ✅ Excel自動生成＆ダウンロード                                         ║
║  ✅ SMTP 経由メール送信（複数宛先対応）                                ║
║  ✅ Excel を添付送信                                                  ║
║  ✅ セキュリティ（Secrets 管理）                                       ║
║  ✅ エラーハンドリング完備                                              ║
║                                                                        ║
║  【セットアップ】                                                      ║
║  1. Streamlit Cloud の「Secrets」に以下を設定                          ║
║     SMTP_SERVER=smtp.gmail.com                                       ║
║     SMTP_PORT=587                                                    ║
║     SMTP_EMAIL=your-email@gmail.com                                 ║
║     SMTP_PASSWORD=your-app-password                                 ║
║                                                                        ║
║  2. requirements.txt に追加（必要な場合）                              ║
║     python-dotenv                                                    ║
║                                                                        ║
╚════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from PIL import Image as PILImage
from io import BytesIO

# ========== 【 設定・定数 】==========
MANUAL_FILE = "manual.xlsx"
MASTER_FILE = "inspector_master.xlsx"
PHOTO_DIR = "photos"
CONFIG_FILE = "app_config.json"

Path(PHOTO_DIR).mkdir(parents=True, exist_ok=True)

# ========== 【 セッション状態の初期化 】==========
if 'inspection_data' not in st.session_state:
    st.session_state.inspection_data = {}
if 'selected_emails' not in st.session_state:
    st.session_state.selected_emails = []
if 'uploaded_photos' not in st.session_state:
    st.session_state.uploaded_photos = {}
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None

# ========== 【 関数定義 】==========

def load_manual():
    """入荷検査マニュアル Excel を読み込み、検査項目を抽出"""
    try:
        wb = openpyxl.load_workbook(MANUAL_FILE)
        ws = wb.worksheets[0]
        
        items = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=45, values_only=False), 1):
            
            # 1. 特定の行番号(30, 31)を除外
            if row_idx in [30, 31]:
                continue

            category_cell = row[0]
            description_cell = row[3]
            
            # 値を文字列として取得
            cat_text = str(category_cell.value or "")
            
            # ▼▼▼▼▼ 修正箇所: 空白を削除し、除外キーワードを厳密にチェックする ▼▼▼▼▼
            # 前後の空白とコロンを削除して、除外キーワードと完全に一致するか確認する
            cleaned_cat_text = cat_text.strip().replace("：", "").replace(":", "")

            # 除外キーワードリスト
            EXCLUDE_KEYWORDS = ["作成部署", "作成者"]

            # cleaned_cat_text が除外キーワードのいずれかを含む、または完全に一致する場合にcontinue
            if cleaned_cat_text in EXCLUDE_KEYWORDS or \
               any(keyword in cat_text for keyword in EXCLUDE_KEYWORDS): 
                continue
            # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
            
            if category_cell.value or description_cell.value:
                category = category_cell.value or ""
                description = description_cell.value or ""
                
                # descriptionが空でなければ追加（このチェックは既存のロジックを維持）
                if str(description).strip():
                    items.append({
                        'id': f"item_{row_idx}",
                        'category': str(category).strip(),
                        'description': str(description).strip(),
                        'row': row_idx
                    })
        
        return items

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return []
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return []
def load_masters():
    """検査者マスター Excel を読み込み"""
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name="検査者一覧")
        return df
    except Exception as e:
        st.error(f"❌ マスター読込エラー: {e}")
        return pd.DataFrame()

def save_config(emails):
    """メール送信先を保存"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'selected_emails': emails}, f, ensure_ascii=False)
    except Exception as e:
        st.warning(f"⚠️ 設定保存エラー: {e}")

def save_photo(uploaded_file, item_id):
    """写真を保存"""
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
    """検査結果を Excel で生成"""
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "検査結果"
        
        # ========== ヘッダー情報 ==========
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
        
        # ========== 検査項目 ==========
        ws['A10'] = "No."
        ws['B10'] = "カテゴリ"
        ws['C10'] = "検査項目"
        ws['D10'] = "判定"
        
        for cell in ['A10', 'B10', 'C10', 'D10']:
            ws[cell].font = Font(bold=True)
            ws[cell].fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        
        row = 11
        for idx, (item_id, data) in enumerate(inspection_data.items(), 1):
            ws[f'A{row}'] = idx
            ws[f'B{row}'] = data['category']
            ws[f'C{row}'] = data['description']
            ws[f'D{row}'] = "合格" if data.get('pass') else "不合格"
            row += 1
        
        # メモリに保存
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
    except Exception as e:
        st.error(f"❌ Excel 作成エラー: {e}")
        return None

def send_email(recipient_emails, subject, body, excel_data, filename):
    """
    SMTP 経由でメール送信（Excel 添付）
    
    【注意】Streamlit Cloud の場合、Secrets に以下を設定：
    SMTP_SERVER=smtp.gmail.com
    SMTP_PORT=587
    SMTP_EMAIL=your-email@gmail.com
    SMTP_PASSWORD=your-app-password
    """
    try:
        # Secrets から SMTP 設定を取得
        smtp_server = st.secrets.get("SMTP_SERVER")
        smtp_port = st.secrets.get("SMTP_PORT", 587)
        smtp_email = st.secrets.get("SMTP_EMAIL")
        smtp_password = st.secrets.get("SMTP_PASSWORD")
        
        if not all([smtp_server, smtp_email, smtp_password]):
            st.error("""
            ❌ SMTP 設定が見つかりません。
            
            Streamlit Cloud で以下を設定してください：
            - SMTP_SERVER
            - SMTP_PORT
            - SMTP_EMAIL
            - SMTP_PASSWORD
            """)
            return False
        
        # メッセージ作成
        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = ', '.join(recipient_emails)
        msg['Subject'] = subject
        
        # メール本文
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Excel を添付
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {filename}')
        msg.attach(part)
        
        # メール送信
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_email, smtp_password)
            server.send_message(msg)
        
        return True
    
    except smtplib.SMTPAuthenticationError:
        st.error("❌ メール認証エラー：パスワード/トークンが間違っています")
        return False
    except smtplib.SMTPException as e:
        st.error(f"❌ SMTP エラー: {e}")
        return False
    except Exception as e:
        st.error(f"❌ メール送信エラー: {e}")
        return False

# ========== 【 UI・ページレイアウト 】==========

st.set_page_config(page_title="入荷検査フォーム", layout="wide")
st.title("🔍 入荷検査フォーム")

# ========== 【 サイドバー 】==========
with st.sidebar:
    st.header("⚙️ 設定")
    
    masters = load_masters()
    if not masters.empty:
        writer_names = masters['氏名'].tolist()
        emails_list = masters['メールアドレス'].tolist()
        
        st.subheader("👤 作業者情報")
        writer_name = st.selectbox("作業者名", writer_names, key="writer")
        reviewer_name = st.selectbox("確認者名", writer_names, key="reviewer")
        
        st.subheader("📧 メール送信先")
        st.caption("（Excel 確認後に送信する場合のみ選択）")
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
    
    tabs = st.tabs(["検査入力", "確認・送信"])
    
    # ========== 【 TAB 1：検査入力 】==========
    with tabs[0]:
        st.subheader("検査項目入力")
        st.caption("各項目について「可」または「否」を選択してください")
        
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
        st.subheader("検査結果確認・送信")
        st.caption("①Excel を確認 → ②メール送信 の流れで進めてください")
        
        if st.session_state.inspection_data:
            # --------- 統計情報 ---------
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
            
            # --------- 検査結果一覧 ---------
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
                        st.session_state.excel_data = excel_data
                        
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
            
            # ========== 【 ステップ 2：メール送信 】==========
            st.subheader("📧 ステップ 2️⃣：メール送信")
            
            if selected_emails and st.session_state.excel_data:
                st.info(f"📬 送信先：{', '.join(selected_emails)}")
                
                if st.button("📮 検査結果をメール送信", use_container_width=True):
                    with st.spinner("📧 メール送信中..."):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"検査結果_{timestamp}.xlsx"
                        
                        subject = f"入荷検査結果 - {in_no} / {lot_no}"
                        body = f"""
入荷検査が完了しました。

【検査情報】
検査ID：{inspector_id}
IN.NO：{in_no}
ロットNO：{lot_no}
作業者：{writer_name}
確認者：{reviewer_name}
検査日：{inspection_date}

【結果】
合格項目：{sum(1 for v in st.session_state.inspection_data.values() if v.get('pass'))}件
不合格項目：{len(st.session_state.inspection_data) - sum(1 for v in st.session_state.inspection_data.values() if v.get('pass'))}件

詳細は添付の Excel ファイルをご確認ください。

---
入荷検査フォーム v3.0
"""
                        
                        success = send_email(
                            selected_emails,
                            subject,
                            body,
                            st.session_state.excel_data,
                            filename
                        )
                        
                        if success:
                            st.success(f"✅ メール送信完了！\n送信先：{', '.join(selected_emails)}")
                        else:
                            st.error("❌ メール送信に失敗しました")
            
            elif not selected_emails:
                st.info("📧 メール送信をご希望の場合は、サイドバーで送信先を選択してください")
            elif not st.session_state.excel_data:
                st.info("📊 先に「Excel を生成・ダウンロード」を実行してください")
        
        else:
            st.info("ℹ️ 検査項目に回答してから「確認・送信」タブをご覧ください")

st.divider()
st.caption("入荷検査フォーム v3.0 | SMTP メール送信完装備版 | 小泉進次郎大臣後押し版")


