import sys
import os
import pytesseract 

def resource_path(relative_path):
    """ PyInstallerでバンドルされたファイルへのパスを取得 """
    try:
        # PyInstallerは一時フォルダに展開し、そのパスを _MEIPASS に格納する
        base_path = sys._MEIPASS
    except Exception:
        # PyInstallerで実行されていない場合（開発時など）はカレントディレクトリ
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Tesseract OCRのパス設定 ---
# アプリ(.exe)と同じ階層に 'tesseract_engine' というフォルダを作り、
# その中にtesseract.exeやtessdataを置いたと仮定
bundled_tesseract_path = resource_path(os.path.join("tesseract_engine", "tesseract.exe"))

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'): # PyInstallerで実行されているかチェック
    print(f"PyInstaller環境を検出。Tesseractのパスを検索: {bundled_tesseract_path}")
    if os.path.exists(bundled_tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = bundled_tesseract_path
        # tessdataのパスも指定する必要がある場合がある (tesseract.exeと同じ階層にtessdataがあれば通常は不要)
        # tessdata_dir_config = f'--tessdata-dir "{resource_path(os.path.join("tesseract_engine", "tessdata"))}"'
        # print(f"Tesseractのtessdataパス設定(試行): {tessdata_dir_config}")
        # (pytesseract.image_to_string の config に tessdata_dir_config を追加することも検討)
        print(f"同梱Tesseractを使用: {pytesseract.pytesseract.tesseract_cmd}")
    else:
        print(f"警告: 同梱Tesseractが見つかりません ({bundled_tesseract_path})。システムPATHのTesseractを探します。")
        # ここでエラーにするか、ユーザーに設定を促すなどの処理が必要
        # もし同梱しない場合は、以下のハードコードされたパスやPATH通しに依存する
        default_tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(default_tesseract_path):
             pytesseract.pytesseract.tesseract_cmd = default_tesseract_path
        else:
            print("エラー: Tesseract OCRがシステムPATHにもデフォルトパスにも見つかりません。")
            # messagebox.showerror("エラー", "Tesseract OCRが見つかりません。") # GUI起動前なのでmessageboxはまだ使えない
else: # 通常のPython環境で実行されている場合
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # 開発時のパス
    print(f"開発環境。Tesseractのパス: {pytesseract.pytesseract.tesseract_cmd}")

# この後に import pytesseract を書くか、既に書かれている場合はこのコードブロックをそれより前に移動
# import pytesseract # ← ここで改めてインポートするか、この設定ブロックを既存のimportより前に

from PIL import Image, UnidentifiedImageError, ImageTk
import requests 
from bs4 import BeautifulSoup 
import urllib.parse 
from googlesearch import search # Google検索用ライブラリ
import pygetwindow # ウィンドウ情報取得用
import mss # スクリーンキャプチャ用
# import mss.tools # Pillow Image変換に直接は不要

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext # scrolledtext をインポート
import re # 正規表現モジュール
import threading # スレッド処理用

# Tesseract OCRのパス指定
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# --- グローバル定数・設定 ---
# !!! 重要：これらの座標はユーザーがキャプチャ画像に合わせて調整する必要があります !!!
MISSION_SLOT_COORDINATES = [
    (670, 371, 1417, 397),  # 1番目の任務名領域 (ユーザー設定値に置き換えてください)
    (670, 473, 1417, 499),  # 2番目の任務名領域 (ユーザー設定値に置き換えてください)
    (670, 575, 1417, 601),  # 3番目の任務名領域 (ユーザー設定値に置き換えてください)
    (670, 677, 1417, 703),  # 4番目の任務名領域 (ユーザー設定値に置き換えてください)
    (670, 779, 1417, 805)   # 5番目の任務名領域 (ユーザー設定値に置き換えてください)
]
captured_kancolle_image_for_gui = None # キャプチャした画像を保持するグローバル変数
ocr_results_for_selection = [] # OCR結果をGUI間で共有するためのリスト（現在は直接使われていない）

# --- グローバル変数 (GUIウィジェットとStringVar) ---
root = None 
status_label_var = None
slot_entry_var = None
mission_name_var = None
site_name_var = None
url_var = None
content_text = None
rewards_text = None
sortie_info_text = None
expedition_info_text = None
arsenal_info_text = None
capture_button_widget = None
slot_entry_widget = None
process_slots_button_widget = None


# --- OCRとウェブページ解析のためのヘルパー関数群 ---

def is_plausible_title_pattern(line_text):
    """その行が任務タイトルらしいパターンか判定する補助関数（最終版に近いもの）"""
    line = line_text.strip()
    if not line or len(line) > 70:
        return False
    norm_line = line.replace('[', '「').replace('【', '「')
    norm_line = norm_line.replace(']', '」').replace('】', '」')
    title_keywords = [
        "開発", "任務", "拡充", "配備", "編成", "計画", "演習", "遠征", "出撃", 
        "改装", "近代化改修", "挑戦", "兵装", "兵站", "哨戒", "作戦"
    ]
    if norm_line.startswith("「") and "」" in norm_line:
        first_closing_bracket_idx = norm_line.find("」") 
        if first_closing_bracket_idx > 0 : 
            text_within_brackets = norm_line[1:first_closing_bracket_idx]
            text_after_closing_bracket = norm_line[first_closing_bracket_idx+1:].strip()
            if any(text_after_closing_bracket.startswith(kw) for kw in title_keywords): return True
            if not text_after_closing_bracket and any(text_within_brackets.endswith(kw) for kw in title_keywords): return True
            if not text_after_closing_bracket and len(text_within_brackets) > 2 and len(text_within_brackets) < 35:
                is_simple_action_verb_inside = any(verb in text_within_brackets for verb in ["廃棄せよ", "準備", "開発せよ"])
                contains_title_keyword_inside = any(kw in text_within_brackets for kw in title_keywords)
                if not is_simple_action_verb_inside or contains_title_keyword_inside: return True
    elif not (line.startswith("「") or line.startswith("【") or line.startswith("[")): 
        if any(line.endswith(kw) for kw in title_keywords):
            if line.count("、") <= 2 and line.count("「") <= 1 and " x" not in line and "NO." not in line.upper():
                if not (len(line.split()) == 2 and any(line.endswith(verb) for verb in ["準備"])): return True
    if line.startswith("精鋭「"): return True
    if line == "敵艦隊を撃破せよ!" or line == "敵艦隊を撃破せよ": return True
    return False

def capture_kancolle_window(
    exact_kancolle_title="艦隊これくしょん -艦これ- - オンラインゲーム - DMM GAMES", 
    generic_kancolle_hints=["艦隊これくしょん -艦これ-", "「艦隊これくしょん -艦これ-」"], 
    browser_hints=["Google Chrome", "Microsoft Edge"]):
    print("\n艦これウィンドウを検索中...")
    try:
        all_windows = pygetwindow.getAllWindows(); target_window = None
        print(f"ステップ1: 「{exact_kancolle_title}」に完全一致するウィンドウを検索します...")
        for window in all_windows:
            if window.title == exact_kancolle_title:
                if window.visible and not window.isMinimized: target_window = window; print(f"  -> 発見: 「{target_window.title}」"); break
        if not target_window:
            print(f"ステップ1では見つかりませんでした。ステップ2: 汎用ヒントで検索します...")
            # print(f"  (汎用タイトルヒント: {generic_kancolle_hints}, ブラウザヒント: {browser_hints})") # 詳細ログは必要なら
            possible_windows = []
            for window in all_windows:
                if any(hint in window.title for hint in generic_kancolle_hints) and \
                   any(browser_hint in window.title for browser_hint in browser_hints):
                    if window.visible and not window.isMinimized: possible_windows.append(window)
            if possible_windows: target_window = possible_windows[0]; print(f"  -> 汎用ヒントで発見: 「{target_window.title}」")
        if not target_window: 
            print("エラー: 対象の艦これウィンドウが見つかりませんでした。\n以下の点を確認してください:")
            print("  - 艦これが起動しており、ウィンドウが表示されている（最小化されていない）。")
            print(f"  - ウィンドウタイトルが「{exact_kancolle_title}」であるか、")
            print(f"    または、タイトルに「{generic_kancolle_hints}」のいずれかと「{browser_hints}」のいずれかが含まれている。")
            return None
        print(f"キャプチャ対象ウィンドウ: 「{target_window.title}」 (サイズ: {target_window.width}x{target_window.height})")
        try:
            if target_window.isMinimized: target_window.restore()
            target_window.activate()
        except Exception as e_act: print(f"警告: ウィンドウのアクティブ化に失敗 (無視して続行): {e_act}")
        monitor_region = {"top": target_window.top, "left": target_window.left, "width": target_window.width, "height": target_window.height}
        with mss.mss() as sct:
            sct_img = sct.grab(monitor_region)
            img = Image.frombytes("RGB", (sct_img.width, sct_img.height), sct_img.rgb); return img
    except pygetwindow.PyGetWindowException as e_gw: print(f"ウィンドウ情報取得エラー (pygetwindow): {e_gw}"); return None
    except UnidentifiedImageError: print("エラー: キャプチャ画像をPillowが認識できませんでした。"); return None
    except Exception as e: print(f"ウィンドウキャプチャ中に予期せぬエラー: {e}"); return None

def ocr_specific_slot(base_image, slot_coords, slot_number_for_debug=0): # デバッグ用引数追加
    if not base_image: print("エラー: ocr_specific_slot 画像がありません。"); return ""
    try:
        x1, y1, x2, y2 = slot_coords
        if x1 >= x2 or y1 >= y2: print(f"警告: スロット {slot_number_for_debug} の座標 {slot_coords} が不正。"); return ""
        slot_img = base_image.crop((x1, y1, x2, y2))
        # デバッグ用にスロット画像を保存したい場合は以下のコメントを解除
        # slot_img.save(f"debug_slot_{slot_number_for_debug}.png")
        # print(f"デバッグ: スロット{slot_number_for_debug}の画像を debug_slot_{slot_number_for_debug}.png として保存しました。")
        text = pytesseract.image_to_string(slot_img, lang='jpn', config='--psm 7').strip()
        return text
    except Exception as e: print(f"スロット {slot_number_for_debug} (座標 {slot_coords}) のOCRエラー: {e}"); return ""

def find_mission_page_url_on_zekamashi(mission_name): # zekamashi.net 直接検索
    cleaned_name = mission_name.replace("|", " ").replace("!", "").replace("[", "").replace("]", "")
    cleaned_name = " ".join(cleaned_name.split()).strip()
    if not cleaned_name: print("エラー: 検索名が空(zekamashi)"); return None
    print(f"\n「{cleaned_name}」で「zekamashi.net」内を検索中...")
    encoded_query = urllib.parse.quote(cleaned_name)
    search_url = f"https://zekamashi.net/?s={encoded_query}"
    print(f"検索URL (zekamashi): {search_url}")
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(search_url, headers=headers, timeout=10)
        response.raise_for_status(); response.encoding = response.apparent_encoding
        soup = BeautifulSoup(response.text, 'html.parser')
        no_results_tag = soup.find(string=lambda text: text and ("何も見つかりませんでした" in text or "お探しのページは見つかりませんでした" in text))
        if no_results_tag and (soup.find(class_="no-results") or soup.find(id="content", class_="no-results") or soup.find("div", class_="error404")): # 色々な「結果なし」パターン
            print("zekamashi.net: 指定された条件では何も見つかりませんでした。"); return None
        
        search_results_articles = soup.select('article.post, article.page, div.search-entry, div.post-item') # 一般的なコンテナ
        if not search_results_articles : search_results_articles = soup.find_all('article') # フォールバック

        query_keywords = [kw for kw in cleaned_name.split() if len(kw) > 1 and not kw.isdigit()]
        best_match_url = None; highest_score = 0

        for article in search_results_articles:
            title_tag = article.find(['h1', 'h2', 'h3'], class_='entry-title') 
            if not title_tag: title_tag = article.find(['h1', 'h2', 'h3']) # クラスなしも
            link_tag = None
            if title_tag:
                link_tag = title_tag.find('a', href=True)
                if not link_tag and title_tag.name == 'a' and title_tag.has_attr('href'): link_tag = title_tag
            if not link_tag: link_tag = article.find('a', href=True) # article直下の最初のリンクも候補に

            if link_tag and link_tag.has_attr('href'):
                link_text = link_tag.get_text(strip=True); href = link_tag['href']; current_score = 0
                for kw in query_keywords:
                    if kw in link_text: current_score += 3 # タイトルテキストの一致を重視
                    elif kw.lower() in link_text.lower(): current_score +=2 # 小文字での一致も少し加点
                    elif kw in href: current_score +=1 
                if "任務" in link_text or "/任務" in href or "quest" in href.lower(): current_score += 2
                if "編成例" in link_text: current_score +=1 # 編成例ページは関連性が高いかも
                
                if current_score > highest_score:
                    highest_score = current_score; best_match_url = urllib.parse.urljoin("https://zekamashi.net/", href)
        
        if best_match_url and highest_score >= 3: # 最低スコアを設定 (調整可能)
            print(f"zekamashi.net: 最も関連性の高いページ候補 -> {best_match_url} (スコア: {highest_score})"); return best_match_url
        print(f"zekamashi.net: 関連性の高いページは見つかりませんでした (最高スコア: {highest_score})。"); return None
    except Exception as e: print(f"zekamashi.net 検索エラー: {e}"); return None

def get_mission_source_urls(selected_mission_name):
    # zekamashi直接検索をまず試す（スコアが良ければそれを採用する改良も可能）
    # direct_zekamashi_url = find_mission_page_url_on_zekamashi(selected_mission_name)
    # if direct_zekamashi_url: # もしここで十分なスコアならGoogle検索をスキップする判断もできる
    #     return [(direct_zekamashi_url, "zekamashi.net")]

    google_query = f"{selected_mission_name} 艦これ 攻略" 
    print(f"\nGoogleで「{google_query}」を検索し、結果を解析します...")
    google_results_urls_from_lib = []
    manual_google_url = f"https://www.google.com/search?q={urllib.parse.quote(google_query)}"
    try:
        print("  Google検索ライブラリで検索実行中..."); temp_results = []
        for url in search(google_query): # 最もシンプルな呼び出し
            temp_results.append(url)
            if len(temp_results) >= 5: break # 上位5件まで
        google_results_urls_from_lib = temp_results
        if google_results_urls_from_lib: print(f"  Google検索から {len(google_results_urls_from_lib)} 件のURL候補を取得。")
        else: print("  Google検索結果なし。")
    except TypeError as te: # search()関数の引数に関するエラーの場合
        print(f"  Google検索関数の呼び出しで引数エラーが発生しました: {te}")
    except Exception as e_general: print(f"  Google検索ライブラリ使用中に予期せぬエラー: {e_general}")

    if not google_results_urls_from_lib:
        print("プログラムによるGoogle検索に失敗したか、結果がありませんでした。手動確認用のURLを提示します。")
        return [(manual_google_url, "Google検索 (手動確認用)")]

    for url in google_results_urls_from_lib:
        if url and "zekamashi.net" in urllib.parse.urlparse(url).netloc:
            print(f"Google検索結果からzekamashi.netのページを優先的に使用: {url}")
            return [(url, "zekamashi.net (via Google)")] 

    print(f"Google検索結果にzekamashi.netは見つかりませんでした。他の上位サイトを提示します (最大3件):")
    top_results_to_return = []
    for url in google_results_urls_from_lib[:3]:
        if url: domain = urllib.parse.urlparse(url).netloc; top_results_to_return.append((url, domain))
    if top_results_to_return: return top_results_to_return
    print("Google検索で有望な結果が見つかりませんでした（フィルタリング後）。")
    return [(manual_google_url, "Google検索 (手動確認用)")]

def parse_tablepress_table(table_soup):
    data = []; headers = []
    if not table_soup: return data
    thead = table_soup.find('thead'); 
    if thead:
        hr = thead.find('tr')
        if hr: 
            for th in hr.find_all('th'): headers.append(th.get_text(strip=True))
    tbody = table_soup.find('tbody')
    if tbody:
        for row_tr in tbody.find_all('tr'):
            row = {}; cells = row_tr.find_all(['td', 'th'])
            for i, cell in enumerate(cells):
                text = cell.get_text(strip=True).replace('\n', ' ')
                if headers and i < len(headers): row[headers[i]] = text
                else: row[f"col_{i+1}"] = text # ヘッダがないか不足する場合のフォールバック
            # colspanを持つセルが１つだけで、それが注釈の場合、その行はスキップ（テーブルデータとしては扱わない）
            if len(cells) == 1 and cells[0].has_attr('colspan') and "参考：" in cells[0].get_text():
                continue
            if row: data.append(row)
    return data

def extract_specific_mission_details(soup, mission_details):
    print("\n詳細情報の抽出を開始します...")
    main_content_area = soup.find(class_=["entry-content", "post-content", "article-body", "main-content", "td-post-content"]) # td-post-content も追加
    if not main_content_area: main_content_area = soup.body
    if not main_content_area: print("エラー: 主要コンテンツエリアが見つかりません。"); return

    # --- 1. 基本的な任務内容と報酬の抽出 ---
    if not mission_details.get("任務内容") and main_content_area: 
        current_mission_content = []; condition_heading_keywords = ["任務情報", "任務内容", "達成条件", "クリア条件", "出現条件", "概要", "任務概要"]; condition_section_found = False
        for kw in condition_heading_keywords:
            heading_tag = main_content_area.find(['h2', 'h3', 'h4'], string=lambda text: text and kw in text.strip())
            if heading_tag:
                # print(f"見出し「{heading_tag.get_text(strip=True)}」から任務内容を抽出試行...")
                for sibling in heading_tag.find_next_siblings():
                    if sibling.name in ['h2', 'h3', 'h4'] or (sibling.name == 'p' and ('報酬は' in sibling.get_text() or 'クリア報酬に' in sibling.get_text())): break
                    if sibling.name == 'ul':
                        for li in sibling.find_all('li', recursive=False): 
                            clean_text = li.get_text(strip=True).replace('\n', ' '); 
                            if clean_text: current_mission_content.append(clean_text)
                        if current_mission_content: condition_section_found = True; break 
                    elif sibling.name == 'p':
                        clean_text = sibling.get_text(strip=True).replace('\n', ' '); 
                        if clean_text: current_mission_content.append(clean_text); # 複数のpが続くことを許容するため、ここではbreakしない
                if current_mission_content: condition_section_found = True # ループ後にフラグを立てる
                if condition_section_found and current_mission_content: break 
        if current_mission_content: mission_details["任務内容"] = current_mission_content
        elif not mission_details.get("任務内容") and main_content_area: 
            reward_p_tag_for_content = main_content_area.find(lambda tag: tag.name == 'p' and ('報酬は' in tag.get_text() or 'クリア報酬に' in tag.get_text()))
            if reward_p_tag_for_content:
                candidate_uls = reward_p_tag_for_content.find_previous_siblings('ul')
                if candidate_uls:
                    condition_ul_tag = candidate_uls[0] 
                    temp_content = [li.get_text(strip=True).replace('\n', ' ') for li in condition_ul_tag.find_all('li', recursive=False) if li.get_text(strip=True)]
                    if temp_content: mission_details["任務内容"] = temp_content
    
    if not mission_details.get("報酬") and main_content_area: 
        current_rewards = []
        reward_p_tag = main_content_area.find(lambda tag: tag.name == 'p' and ('報酬は' in tag.get_text() or 'クリア報酬に' in tag.get_text()))
        if reward_p_tag:
            reward_ul_tag = reward_p_tag.find_next_sibling('ul')
            if reward_ul_tag: current_rewards = [li.get_text(strip=True).replace('\n', ' ') for li in reward_ul_tag.find_all('li', recursive=False) if li.get_text(strip=True)]
        if current_rewards: mission_details["報酬"] = current_rewards

    # --- 2. 工廠任務特有のテーブル抽出 ---
    dev_recipe_table_found = False; dev_keywords = ["開発", "レシピ", "工廠", "改修"] # 「改修」も追加
    if main_content_area:
        possible_headings_for_dev = main_content_area.find_all(['h2', 'h3', 'h4', 'h5'])
        for heading in possible_headings_for_dev:
            if any(kw in heading.get_text(strip=True) for kw in dev_keywords):
                table = heading.find_next_sibling('table', class_=lambda x: x and 'tablepress' in x)
                if table: mission_details["開発レシピ表"] = parse_tablepress_table(table); dev_recipe_table_found = True; break 
    if not dev_recipe_table_found and main_content_area: 
        all_tables = main_content_area.find_all('table', class_=lambda x: x and 'tablepress' in x)
        for table_cand in all_tables: # ページ内の全てのtablepressをチェック
            caption = table_cand.find('caption'); first_th = table_cand.find('th')
            if (caption and any(kw in caption.get_text(strip=True) for kw in dev_keywords)) or \
               (first_th and any(kw in first_th.get_text(strip=True) for kw in dev_keywords)):
                mission_details["開発レシピ表"] = parse_tablepress_table(table_cand); break # 最初に見つかったものを採用

    # --- 3. 遠征任務特有の情報抽出 ---
    expedition_details_list = []
    if main_content_area:
        expedition_name_headings = main_content_area.find_all('h3') 
        for h3_tag in expedition_name_headings:
            h3_text = h3_tag.get_text(strip=True)
            match = re.match(r"^(ID:)?([A-Z]?\d{1,2}(\-[A-Z\d]{1,2})?)\s?[:：]?\s*(.+)", h3_text)
            is_generic_explanation = "とは？" in h3_text or "まとめ" in h3_text or "一覧" in h3_text or "について" in h3_text or "関連記事" in h3_text
            if match and not is_generic_explanation :
                expedition_name_from_h3 = h3_text
                expedition_table = h3_tag.find_next_sibling('table', class_=lambda x: x and 'tablepress' in x)
                if expedition_table:
                    table_data = parse_tablepress_table(expedition_table)
                    expedition_details_list.append({"遠征名": expedition_name_from_h3, "情報表": table_data})
    if expedition_details_list: mission_details["遠征詳細"] = expedition_details_list

    # --- 4. 出撃任務特有の情報抽出 ---
    sortie_details_list = []
    processed_specific_map_sortie = False # 個別の海域名が見つかったかのフラグ
    if main_content_area:
        map_name_headings = main_content_area.find_all(['h3', 'h4']) 
        
        for h_tag in map_name_headings:
            h_text = h_tag.get_text(strip=True)
            map_match_pattern = r"^(\d-\d(?:-\w)?|\d-\d\S*|\S*\d-\d\S*|EO海域|鎮守府海域(?:-\d)?|西方海域|中部海域|北方海域|南方海域|Extra Operation)"
            map_match = re.match(map_match_pattern, h_text, re.IGNORECASE)
            span_id_match = h_tag.find('span', id=lambda x: x and x.startswith('i-'))
            is_generic_heading = "とは？" in h_text or "まとめ" in h_text or "一覧" in h_text or "について" in h_text or "その他" in h_text or "関連記事" in h_text or "コメント" in h_text

            if (map_match or span_id_match) and not is_generic_heading:
                map_name = h_text
                current_sortie_info = {"海域": map_name, "編成例": [], "編成備考": []}
                
                for sibling in h_tag.find_next_siblings():
                    if sibling.name in ['h2','h3','h4'] or (sibling.name=='p' and ('報酬は' in sibling.get_text(strip=True) or 'クリア報酬に' in sibling.get_text(strip=True))): break
                    text_to_add = ""
                    if sibling.name == 'p': text_to_add = sibling.get_text(strip=True).replace('\n', ' ')
                    elif sibling.name == 'ul':
                        for li in sibling.find_all('li', recursive=False):
                            li_text = li.get_text(strip=True).replace('\n', ' '); 
                            if li_text: current_sortie_info["編成備考"].append(li_text)
                        continue
                    if text_to_add:
                        is_primary_fleet = False
                        if ("【" in text_to_add and "】" in text_to_add and any(st in text_to_add for st in ["駆","軽","重","戦","航","潜","母","巡","艦","海防"])) or \
                           re.match(r"^([^\s「【●※]+?\d{1,2})+.*", text_to_add):
                            if not (text_to_add.lstrip().startswith("●") or text_to_add.lstrip().startswith("※")): is_primary_fleet = True
                        if is_primary_fleet: current_sortie_info["編成例"].append(text_to_add)
                        else: current_sortie_info["編成備考"].append(text_to_add)
                if current_sortie_info["編成例"] or current_sortie_info["編成備考"]:
                    sortie_details_list.append(current_sortie_info)
                    processed_specific_map_sortie = True # 個別海域情報を処理したフラグを立てる
        
        # --- ↓↓↓ もし上記のh3/h4での海域情報が見つからなかった場合のフォールバック ↓↓↓ ---
        if not processed_specific_map_sortie: # (または sortie_details_list が空の場合)
            # h2 タグなどで「編成例」という見出しを探す
            general_fleet_heading = main_content_area.find(['h2', 'h3'], string=lambda text: text and "編成例" in text.strip() and not ("とは？" in text.strip() or "まとめ" in text.strip()))
            if general_fleet_heading:
                print(f"一般的な「{general_fleet_heading.get_text(strip=True)}」セクションから情報を抽出します。")
                # この場合の「海域」名は、ページタイトルから取るか、固定の文字列にする
                page_title_text = mission_details.get("タイトル", "").split('｜')[0].strip() # ページの主タイトル部分
                map_name_for_general = page_title_text if page_title_text else "(主要攻略)" # タイトルが取れなければ汎用名

                current_sortie_info = {"海域": map_name_for_general, "編成例": [], "編成備考": []}
                for sibling in general_fleet_heading.find_next_siblings():
                    if sibling.name in ['h2','h3','h4'] or (sibling.name=='p' and ('報酬は' in sibling.get_text(strip=True) or 'クリア報酬に' in sibling.get_text(strip=True))): break
                    text_to_add = ""
                    if sibling.name == 'p': text_to_add = sibling.get_text(strip=True).replace('\n', ' ')
                    elif sibling.name == 'ul':
                        for li in sibling.find_all('li', recursive=False):
                            li_text = li.get_text(strip=True).replace('\n', ' '); 
                            if li_text: current_sortie_info["編成備考"].append(li_text)
                        continue
                    if text_to_add:
                        is_primary_fleet = False
                        if ("【" in text_to_add and "】" in text_to_add and any(st in text_to_add for st in ["駆","軽","重","戦","航","潜","母","巡","艦","海防"])) or \
                           re.match(r"^([^\s「【●※]+?\d{1,2})+.*", text_to_add):
                            if not (text_to_add.lstrip().startswith("●") or text_to_add.lstrip().startswith("※")): is_primary_fleet = True
                        if is_primary_fleet: current_sortie_info["編成例"].append(text_to_add)
                        else: current_sortie_info["編成備考"].append(text_to_add)
                if current_sortie_info["編成例"] or current_sortie_info["編成備考"]:
                    sortie_details_list.append(current_sortie_info)

    if sortie_details_list:
        mission_details["出撃情報"] = sortie_details_list


# --- GUIイベントハンドラ関数 ---
def clear_mission_details_gui():
    global mission_name_var, site_name_var, url_var, content_text, rewards_text, sortie_info_text, expedition_info_text, arsenal_info_text
    if mission_name_var: mission_name_var.set("")
    if site_name_var: site_name_var.set("")
    if url_var: url_var.set("")
    widgets = [content_text, rewards_text, sortie_info_text, expedition_info_text, arsenal_info_text]
    for widget in widgets:
        if widget and widget.winfo_exists(): widget.config(state=tk.NORMAL); widget.delete('1.0', tk.END); widget.config(state=tk.DISABLED)

def update_mission_details_gui(details_dict):
    """抽出された詳細情報 (details_dict) を対応するGUIウィジェットに表示する (サブタブ対応・スタイル適用版)"""
    # グローバル変数として定義されたGUIウィジェットとStringVarを参照
    global mission_name_var, site_name_var, url_var
    global content_text, rewards_text, arsenal_info_text # これらは直接 ScrolledText を参照
    global tab_sortie, tab_expedition # これらはサブNotebookを配置する親の ttk.Frame を参照
    global root

    if not root or not root.winfo_exists(): 
        print("エラー: update_mission_details_gui - rootウィンドウが無効です。")
        return 

    clear_mission_details_gui() # まず既存の表示をクリア

    # 任務名、サイト名、URLをStringVar経由でラベルに設定
    if mission_name_var: mission_name_var.set(details_dict.get("タイトル", "N/A"))
    if site_name_var: site_name_var.set(details_dict.get("サイト名", "N/A"))
    if url_var: url_var.set(details_dict.get("URL", "N/A"))

    # --- 任務内容タブの処理 ---
    if content_text and content_text.winfo_exists():
        content_text.config(state=tk.NORMAL); content_text.delete('1.0', tk.END)
        data_content = details_dict.get("任務内容")
        if data_content and isinstance(data_content, list) and data_content:
            for item in data_content: content_text.insert(tk.END, f"- {item}\n")
        else: content_text.insert(tk.END, "(情報なし)\n")
        content_text.config(state=tk.DISABLED)

    # --- 報酬タブの処理 ---
    if rewards_text and rewards_text.winfo_exists():
        rewards_text.config(state=tk.NORMAL); rewards_text.delete('1.0', tk.END)
        data_rewards = details_dict.get("報酬")
        if data_rewards and isinstance(data_rewards, list) and data_rewards:
            for item in data_rewards: rewards_text.insert(tk.END, f"- {item}\n")
        else: rewards_text.insert(tk.END, "(情報なし)\n")
        rewards_text.config(state=tk.DISABLED)

    # --- 出撃情報タブの処理 (サブタブ形式) ---
    if 'tab_sortie' in globals() and tab_sortie and tab_sortie.winfo_exists():
        for widget in tab_sortie.winfo_children(): # 既存のウィジェットをクリア
            widget.destroy()

        sortie_data_list = details_dict.get("出撃情報")
        if sortie_data_list and isinstance(sortie_data_list, list) and sortie_data_list:
            sortie_notebook = ttk.Notebook(tab_sortie) # 出撃情報タブ内に新しいNotebook
            sortie_notebook.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)

            for sortie_item in sortie_data_list:
                map_name = sortie_item.get('海域', '海域不明')
                map_tab_frame = ttk.Frame(sortie_notebook) 
                sortie_notebook.add(map_tab_frame, text=map_name[:20]) # タブ名は20文字まで

                map_details_st = scrolledtext.ScrolledText(map_tab_frame, wrap=tk.WORD, height=10, relief=tk.GROOVE, borderwidth=1) # heightは適宜調整
                map_details_st.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)
                
                # スタイルタグを設定 (フォント設定は __main__ で定義したものを参照する形が良いが、ここでは直接定義)
                font_family = "MS Gothic" # または適切なフォント
                font_size_normal = 9
                font_size_large_bold = font_size_normal + 3
                
                map_details_st.tag_configure("fleet_example_style", font=(font_family, font_size_large_bold, "bold"), foreground="blue")
                map_details_st.tag_configure("sub_header_style", font=(font_family, font_size_normal + 1, "bold"))

                map_details_st.config(state=tk.NORMAL)
                if sortie_item.get("編成例"):
                    map_details_st.insert(tk.END, "編成例:\n", "sub_header_style")
                    for comp in sortie_item.get("編成例"):
                        map_details_st.insert(tk.END, f"  - {comp}\n", "fleet_example_style") 
                if sortie_item.get("編成備考"):
                    map_details_st.insert(tk.END, "\n編成備考:\n", "sub_header_style")
                    for note in sortie_item.get("編成備考"):
                        map_details_st.insert(tk.END, f"  - {note}\n")
                map_details_st.config(state=tk.DISABLED)
        else:
            ttk.Label(tab_sortie, text="(出撃情報なし)").pack(padx=5, pady=5)

    # --- 遠征詳細タブの処理 (サブタブ形式) ---
    if 'tab_expedition' in globals() and tab_expedition and tab_expedition.winfo_exists():
        for widget in tab_expedition.winfo_children(): # 既存のウィジェットをクリア
            widget.destroy()
        expedition_data_list = details_dict.get("遠征詳細")
        if expedition_data_list and isinstance(expedition_data_list, list) and expedition_data_list:
            expedition_notebook = ttk.Notebook(tab_expedition)
            expedition_notebook.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)
            for ed_item in expedition_data_list:
                exp_name = ed_item.get('遠征名', '遠征名不明')
                exp_tab_frame = ttk.Frame(expedition_notebook)
                expedition_notebook.add(exp_tab_frame, text=exp_name[:20])

                exp_details_st = scrolledtext.ScrolledText(exp_tab_frame, wrap=tk.WORD, height=6, relief=tk.GROOVE, borderwidth=1)
                exp_details_st.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)
                exp_details_st.config(state=tk.NORMAL)

                # 遠征詳細用のスタイルタグ (必要なら)
                # font_family = "MS Gothic"; font_size_normal = 9
                # exp_details_st.tag_configure("sub_header_style", font=(font_family, font_size_normal + 1, "bold"))

                if ed_item.get("情報表"):
                    # exp_details_st.insert(tk.END, "情報表:\n", "sub_header_style") # オプション
                    for table_row_dict in ed_item.get("情報表", []):
                        row_str = "  "
                        for rk, rv in table_row_dict.items():
                            row_str += f"{rk}: {rv}; "
                        exp_details_st.insert(tk.END, row_str.strip().rstrip(';') + "\n")
                exp_details_st.config(state=tk.DISABLED)
        else:
            ttk.Label(tab_expedition, text="(遠征詳細なし)").pack(padx=5, pady=5)
            
    # --- 開発レシピ表タブの処理 ---
    if 'arsenal_info_text' in globals() and arsenal_info_text and arsenal_info_text.winfo_exists():
        widget = arsenal_info_text
        widget.config(state=tk.NORMAL); widget.delete('1.0', tk.END)
        data_arsenal = details_dict.get("開発レシピ表")
        if data_arsenal and isinstance(data_arsenal, list) and data_arsenal:
            for item_idx, item in enumerate(data_arsenal):
                if isinstance(item, dict):
                    display_text = f"  ● ({item_idx + 1}) "
                    for dk, dv in item.items(): display_text += f"{dk}: {dv}; "
                    widget.insert(tk.END, display_text.strip().rstrip(';') + "\n")
                else: widget.insert(tk.END, f"- {item}\n")
        else: widget.insert(tk.END, "(情報なし)\n")
        widget.config(state=tk.DISABLED)
            
    if root and root.winfo_exists(): 
        root.update_idletasks()

def process_one_mission_in_thread(slot_idx, ocr_name, status_var_ref, root_ref):
    global processed_missions_details_gui # スレッド内からは直接GUIリストを更新しない方が安全
    
    def schedule_update(func, *args):
        if root_ref and root_ref.winfo_exists(): root_ref.after(0, lambda: func(*args))
    def update_status(msg):
        if status_var_ref and isinstance(status_var_ref, tk.StringVar): schedule_update(status_var_ref.set, msg)

    update_status(f"スロット{slot_idx+1}:「{ocr_name[:15]}...」検索中...")
    final_details = None
    try:
        source_opts = get_mission_source_urls(ocr_name)
        if not source_opts: schedule_update(messagebox.showwarning, "情報源なし", f"「{ocr_name}」の情報源が見つかりません。"); update_status(f"スロット{slot_idx+1}:情報源なし"); return
        
        chosen_url, chosen_site = source_opts[0] # 最初の候補を使用
        # GUIへの途中経過表示は update_mission_details_gui でまとめて行う
        if "Google検索 (手動確認用)" in chosen_site:
            schedule_update(messagebox.showinfo, "手動確認", f"「{ocr_name}」はGoogle検索URL参照:\n{chosen_url}"); update_status(f"スロット{slot_idx+1}:Google手動確認"); return

        update_status(f"スロット{slot_idx+1}:「{chosen_site}」から取得中...")
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        page_resp = requests.get(chosen_url, headers=headers, timeout=15); page_resp.raise_for_status(); page_resp.encoding = page_resp.apparent_encoding
        html = page_resp.text; soup_obj = BeautifulSoup(html, 'html.parser')
        title_tag = soup_obj.find('title'); title = title_tag.get_text(strip=True) if title_tag else "タイトル不明"
        
        final_details = {"OCR任務名": ocr_name, "タイトル": title, "サイト名": chosen_site, "URL": chosen_url, "任務内容": [], "報酬": [], "出撃情報": [], "遠征詳細": [], "開発レシピ表": []}
        extract_specific_mission_details(soup_obj, final_details)
        
        schedule_update(update_mission_details_gui, final_details) 
        # processed_missions_details_gui.append(final_details) # メインスレッド側で追加した方が安全
        update_status(f"スロット{slot_idx+1}:「{ocr_name[:15]}...」表示完了。")
    except Exception as e:
        error_msg = f"「{ocr_name}」処理中エラー: {e}"; schedule_update(messagebox.showerror, "処理エラー", error_msg); update_status(f"スロット{slot_idx+1}:エラー発生。")

def handle_capture_button_click():
    global captured_kancolle_image_for_gui, status_label_var, slot_entry_widget, process_slots_button_widget, root
    print("キャプチャボタンクリック。"); status_label_var.set("艦これウィンドウを検索中..."); root.update_idletasks()
    img = capture_kancolle_window()
    if img:
        captured_kancolle_image_for_gui = img; messagebox.showinfo("成功", "艦これウィンドウキャプチャ成功！\nステップ2で処理スロットを指定してください。")
        status_label_var.set("キャプチャ成功！ステップ2へ。"); 
        if slot_entry_widget: slot_entry_widget.config(state=tk.NORMAL)
        if process_slots_button_widget: process_slots_button_widget.config(state=tk.NORMAL)
        clear_mission_details_gui()
    else: messagebox.showerror("エラー", "艦これウィンドウをキャプチャできませんでした。"); status_label_var.set("キャプチャ失敗。")

def handle_process_slots_button_click():
    global captured_kancolle_image_for_gui, status_label_var, slot_entry_var, root, MISSION_SLOT_COORDINATES
    if not captured_kancolle_image_for_gui: messagebox.showwarning("注意", "先にウィンドウをキャプチャ。"); return
    choice_str = slot_entry_var.get()
    if not choice_str: messagebox.showwarning("入力エラー", "スロット番号未入力。"); return
    clear_mission_details_gui(); status_label_var.set("選択スロットの処理を開始します..."); root.update_idletasks()
    selected_indices = []
    try:
        if choice_str.lower() == 'all': selected_indices = list(range(len(MISSION_SLOT_COORDINATES)))
        else:
            parts = choice_str.split(','); temp_indices = []; valid = True
            for p_str in parts:
                s_idx_str = p_str.strip(); 
                if not s_idx_str: continue
                idx = int(s_idx_str) - 1
                if 0 <= idx < len(MISSION_SLOT_COORDINATES): 
                    if idx not in temp_indices: temp_indices.append(idx)
                else: messagebox.showerror("入力エラー", f"{idx+1}は無効な番号。"); valid=False; break
            if valid and temp_indices: selected_indices = sorted(temp_indices)
            elif valid and not temp_indices: messagebox.showwarning("入力エラー", "有効な番号なし。"); status_label_var.set("入力エラー。"); return
            elif not valid: status_label_var.set("入力エラー。"); return
    except ValueError: messagebox.showerror("入力エラー", "番号をカンマ区切りで。例:1,2,3 or all"); status_label_var.set("入力エラー。"); return
    except Exception as e: messagebox.showerror("エラー", f"スロット番号処理中エラー: {e}"); status_label_var.set("処理エラー。"); return
    if not selected_indices: status_label_var.set("処理スロット未選択。"); return
    
    status_label_var.set(f"処理対象スロット: {[i+1 for i in selected_indices]} の処理を開始...")
    root.update_idletasks()
    
    for slot_idx in selected_indices:
        coords = MISSION_SLOT_COORDINATES[slot_idx]
        ocr_name = ocr_specific_slot(captured_kancolle_image_for_gui, coords, slot_idx + 1) # デバッグ用に番号渡し
        if not ocr_name: messagebox.showwarning("OCR結果なし", f"スロット {slot_idx + 1} からテキストを読み取れませんでした。スキップします。"); continue
        
        thread = threading.Thread(target=process_one_mission_in_thread, args=(slot_idx, ocr_name, status_label_var, root), daemon=True)
        thread.start()
    # 全てのスレッドが開始された後のメッセージ
    # status_label_var.set("全選択スロットの処理を開始しました。(バックグラウンド実行)") # これは各スレッド完了時に更新

def copy_url_to_clipboard():
    """表示されているURLをクリップボードにコピーする"""
    global url_var, root, status_label_var # root と url_var, status_label_var をグローバル変数として参照
    
    # root や url_var が初期化されているか（GUIが作られているか）確認
    if root is None or url_var is None:
        messagebox.showerror("エラー", "GUIが初期化されていません。")
        if status_label_var: status_label_var.set("GUI未初期化エラー")
        return

    url_to_copy = url_var.get()
    if url_to_copy and url_to_copy != "N/A" and "Google検索 (手動確認用)" not in url_to_copy : # URLが有効な場合
        try:
            root.clipboard_clear() # クリップボードをクリア
            root.clipboard_append(url_to_copy) # URLをクリップボードに追加
            # root.update() # 即時反映 (通常は不要)
            messagebox.showinfo("コピー完了", f"以下のURLをクリップボードにコピーしました:\n{url_to_copy}")
            if status_label_var: status_label_var.set("URLをクリップボードにコピーしました。")
        except tk.TclError:
            # クリップボードが利用できない環境など
            messagebox.showerror("コピー失敗", "クリップボードへのアクセスに失敗しました。")
            if status_label_var: status_label_var.set("クリップボードへのコピーに失敗しました。")
        except Exception as e:
            messagebox.showerror("コピー失敗", f"予期せぬエラーが発生しました: {e}")
            if status_label_var: status_label_var.set(f"URLコピー中にエラー: {e}")
    elif "Google検索 (手動確認用)" in url_to_copy:
        messagebox.showwarning("コピー対象外", "これは手動確認用のGoogle検索URLのため、コピー機能の対象外です。")
        if status_label_var: status_label_var.set("手動確認用URLはコピー対象外です。")
    else:
        messagebox.showwarning("コピー対象なし", "コピーする有効なURLがありません。")
        if status_label_var: status_label_var.set("コピーするURLがありません。")

# --- Tkinter GUIのメイン処理 ---
if __name__ == "__main__":
    root = tk.Tk() 
    status_label_var = tk.StringVar()
    slot_entry_var = tk.StringVar()
    mission_name_var = tk.StringVar()
    site_name_var = tk.StringVar()
    url_var = tk.StringVar()

    root.title("艦これ任務サポート GUI (v0.7)")
    root.geometry("850x750") 

    style = ttk.Style()
    try: 
        if 'vista' in style.theme_names(): style.theme_use('vista')
    except tk.TclError: pass

    main_frame = ttk.Frame(root, padding="10")
    main_frame.pack(expand=True, fill=tk.BOTH)

    input_controls_frame = ttk.Frame(main_frame) # キャプチャとスロット指定をまとめるフレーム
    input_controls_frame.pack(fill=tk.X, pady=5, padx=5)

    capture_frame = ttk.LabelFrame(input_controls_frame, text="ステップ1: ウィンドウキャプチャ", padding="10")
    capture_frame.pack(side=tk.LEFT, padx=(0,5), fill=tk.X) # expand=True を削除または調整
    capture_button_widget = ttk.Button(capture_frame, text="艦これウィンドウをキャプチャ", command=handle_capture_button_click)
    capture_button_widget.pack(pady=5, padx=5)

    slot_selection_frame = ttk.LabelFrame(input_controls_frame, text="ステップ2: 処理スロット指定 & 実行", padding="10")
    slot_selection_frame.pack(side=tk.LEFT, padx=(5,0), fill=tk.X, expand=True) # こちらを expand=True に
    slot_entry_label = ttk.Label(slot_selection_frame, text="処理スロット番号 (例: 1,2 or all):")
    slot_entry_label.pack(side=tk.LEFT, padx=(0,5))
    slot_entry_widget = ttk.Entry(slot_selection_frame, textvariable=slot_entry_var, width=20)
    slot_entry_widget.pack(side=tk.LEFT, padx=5); slot_entry_widget.config(state=tk.DISABLED)
    process_slots_button_widget = ttk.Button(slot_selection_frame, text="選択スロットの情報取得・表示", command=handle_process_slots_button_click)
    process_slots_button_widget.pack(side=tk.LEFT, padx=5); process_slots_button_widget.config(state=tk.DISABLED)

    results_display_frame = ttk.LabelFrame(main_frame, text="ステップ3: 抽出された任務詳細", padding="10")
    results_display_frame.pack(expand=True, fill=tk.BOTH, pady=5, padx=5) # ← 元の設定に戻してみる
    
    top_info_frame = ttk.Frame(results_display_frame)
    top_info_frame.pack(fill=tk.X, pady=(5,10))

    ttk.Label(top_info_frame, text="任務名 (タイトル):", anchor=tk.W).grid(row=0, column=0, sticky=tk.NW, padx=2, pady=2)
    mission_name_label_widget = ttk.Label(top_info_frame, textvariable=mission_name_var, wraplength=700, anchor=tk.NW, justify=tk.LEFT, relief=tk.FLAT, borderwidth=1, padding=2)
    mission_name_label_widget.grid(row=0, column=1, sticky=tk.EW, padx=2, pady=2)

    ttk.Label(top_info_frame, text="情報源サイト名:", anchor=tk.W).grid(row=1, column=0, sticky=tk.NW, padx=2, pady=2)
    site_name_label_widget = ttk.Label(top_info_frame, textvariable=site_name_var, anchor=tk.NW, justify=tk.LEFT, relief=tk.FLAT, borderwidth=1, padding=2)
    site_name_label_widget.grid(row=1, column=1, sticky=tk.EW, padx=2, pady=2)
    
    url_frame_for_copy = ttk.Frame(top_info_frame) # URLとコピーボタンをまとめるフレーム
    url_frame_for_copy.grid(row=2, column=1, sticky=tk.EW, padx=2, pady=2)
    ttk.Label(top_info_frame, text="情報源URL:", anchor=tk.W).grid(row=2, column=0, sticky=tk.NW, padx=2, pady=2)
    url_label_widget = ttk.Label(url_frame_for_copy, textvariable=url_var, wraplength=600, anchor=tk.NW, justify=tk.LEFT, relief=tk.FLAT, borderwidth=1, padding=2) 
    url_label_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
    copy_url_button = ttk.Button(url_frame_for_copy, text="コピー", command=copy_url_to_clipboard, width=6)
    copy_url_button.pack(side=tk.LEFT, padx=(5,0))

    top_info_frame.columnconfigure(1, weight=1) 

    notebook = ttk.Notebook(results_display_frame) # results_display_frame を親に
    notebook.pack(expand=True, fill=tk.BOTH, pady=5) # notebook自体は親フレーム内で拡張

    tab_content = ttk.Frame(notebook); notebook.add(tab_content, text='任務内容')
    content_text = scrolledtext.ScrolledText(tab_content, wrap=tk.WORD, height=4, state=tk.DISABLED, relief=tk.GROOVE, borderwidth=1) # 例: 2->4
    content_text.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)

    tab_rewards = ttk.Frame(notebook); notebook.add(tab_rewards, text='報酬')
    rewards_text = scrolledtext.ScrolledText(tab_rewards, wrap=tk.WORD, height=4, state=tk.DISABLED, relief=tk.GROOVE, borderwidth=1) # 例: 2->4
    rewards_text.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)

    tab_sortie = ttk.Frame(notebook); notebook.add(tab_sortie, text='出撃情報')
    sortie_info_text = scrolledtext.ScrolledText(tab_sortie, wrap=tk.WORD, height=25, state=tk.DISABLED, relief=tk.GROOVE, borderwidth=1) # 例: 4->10 (ここは情報量が多いので大きく)
    sortie_info_text.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)
    
    # --- ↓↓↓ 「出撃情報」の「編成例」用スタイルタグを設定 ↓↓↓ ---
    # ベースとなるフォント情報を取得（またはデフォルト値を設定）
    # ウィジェットが実際に表示されるまでは正確なフォント取得が難しい場合があるため、
    # 固定値で設定するか、一般的なフォントを指定するのが無難です。
    # ここでは、現在のフォントサイズを基準に少し大きくしてみます。
    # (実際のフォントファミリーやサイズは、お使いの環境や好みで調整してください)
    default_font_family = "MS Gothic" 
    default_font_size = 9          # ScrolledTextの基本フォントサイズを想定
    
    sortie_info_text.tag_configure(
        "fleet_example_style", 
        font=(default_font_family, default_font_size + 3, "bold"), # サイズを+3に、太字
        foreground="blue"                                        # 文字色を青に (下線指定を削除)
    )
    
    tab_expedition = ttk.Frame(notebook); notebook.add(tab_expedition, text='遠征詳細')
    expedition_info_text = scrolledtext.ScrolledText(tab_expedition, wrap=tk.WORD, height=6, state=tk.DISABLED, relief=tk.GROOVE, borderwidth=1) # 例: 3->6
    expedition_info_text.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)

    tab_arsenal = ttk.Frame(notebook); notebook.add(tab_arsenal, text='開発レシピ')
    arsenal_info_text = scrolledtext.ScrolledText(tab_arsenal, wrap=tk.WORD, height=6, state=tk.DISABLED, relief=tk.GROOVE, borderwidth=1) # 例: 3->6
    arsenal_info_text.pack(expand=True, fill=tk.BOTH, padx=2, pady=2)

    status_bar = ttk.Label(root, textvariable=status_label_var, relief=tk.SUNKEN, anchor=tk.W, padding=(5,2))
    status_label_var.set("「艦これウィンドウをキャプチャ」ボタンを押してください。")
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # processed_missions_details_gui = [] # これはメインループの外、関数の外でグローバルとして初期化済み想定

    root.mainloop()