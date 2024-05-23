'''
License

Copyright © 2024 Taichi Oikawa

This project is licensed under the MIT License, see the LICENSE.txt file for details.
'''

import os
import math
import json
import csv
import time

import openpyxl as xl
from openpyxl.styles.borders import Border, Side
import questionary
from termcolor import cprint


# questionary while関数
def questionary_text(prompt: str) -> str:
    while True:
        answer: str = questionary.text(prompt).ask()
        if answer != '':
            return answer

def questionary_int(prompt: str) -> int:
    while True:
        answer: str = questionary.text(prompt).ask()
        if answer != '' and answer.isdigit():
            return int(answer)

# 実数か判定する関数 (float)
def is_float(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False

# 時間を計測する関数
def time_count(start: float, end: float) -> str:
    time = end - start
    s = int(time)
    ms = int((time - s) * 1000)
    return f'{s}s {ms}ms'

# ファイルを読み込む関数
def read_text_file(file_path: str) -> list[str]:
    with open(file_path, 'r', encoding="Shift-JIS") as f:
        return f.readlines()

# ファイルに書き込む関数
def write_text_file(file_path: str, data: str) -> None:
    with open(file_path, 'w', encoding="Shift-JIS") as f:
        f.write(data)

# 1文字ずつカンマを挿入する関数
def insert_commas(text: str) -> str:
    result: list[str] = []
    for i, t in enumerate(text):
        if i == len(text) - 1:
            result.append(t)
        else:
            result.append(f'{t},')
    return ''.join(result)

# 辞書型の値を合計する関数
def add_dictionaries(dict1: dict, dict2: dict) -> dict:
    result: dict = {}

    for key in dict1:
        result[key] = dict1[key]

    for key in dict2:
        if key in result:
            result[key] += dict2[key]
        else:
            result[key] = dict2[key]

    return result


## MAIN ##

# タイトル表示, ライセンス表示
cprint('\n================= [全校制作] モザイク画 CSV変換プログラム =================\n', 'green')
print('\nCopyright © 2024 Taichi Oikawa\nThis project is licensed under the MIT License, see the LICENSE.txt file for details.\n')

# カレントディレクトリを取得
path: str = os.getcwd()
# ファイルリストを取得
filelist: list[str] = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]

cprint('\n--- プログラムを開始します ---\n', 'light_magenta')

# outputディレクトリが存在しない場合は作成
if os.path.isdir(os.path.join(path, 'output')) == False:
    try:
        os.mkdir(os.path.join(path, 'output'))
    except Exception as e:
        cprint(f'outputディレクトリの作成に失敗しました。\n---------- エラーメッセージ ----------\n{e}', 'red')
        input('\n -- 終了するには何かキーを押してください -- ')
        os._exit(0)


# jsonファイルを読み込む
if os.path.exists('settings.json'):
    with open('settings.json', 'r') as f:
        json_data = json.load(f)

    classes: list[str] = json_data['ClassNames'].split(',') if 'ClassNames' in json_data else None
    pieces_per_origami: int = int(json_data['PiecesPerOrigami']) if 'PiecesPerOrigami' in json_data else None
    blocks_per_class: int = int(json_data['BlocksPerClass']) if 'BlocksPerClass' in json_data else None

    EXCEL_DATA: dict[str, str] = json_data['Excel'] if 'Excel' in json_data else None

    if classes == None:
        cprint('設定ファイルにクラス名のデータがありません。', 'red')
        input('\n -- 終了するには何かキーを押してください -- ')
        os._exit(0)
else:
    cprint('設定ファイルが見つかりませんでした。', 'red')
    input('\n -- 終了するには何かキーを押してください -- ')
    os._exit(0)



# モザイク画元データを取得
file_choices = []
for file in filelist:
    if file.endswith('.txt'):
        file_choices.append(file)

if file_choices == []:
    cprint('モザイク画データが見つかりませんでした。', 'red')
    input('\n -- 終了するには何かキーを押してください -- ')
    os._exit(0)


# プロパティの入力1
mosaic_data_file: str = questionary.select('モザイク画データを選択してください: ', choices=file_choices).ask()
output_file_name: str = questionary_text('出力ファイル名を入力してください（拡張子は不要）: ')

if pieces_per_origami == None:
    print('\n※ 1cm角の場合は15*15で225となります')
    pieces_per_origami: int = questionary_int('1枚の折り紙で作成できるピース数を入力してください: ')
else:
    print(f'\n1枚の折り紙で作成できるピース数 (設定ファイル): {pieces_per_origami}')

if blocks_per_class == None:
    blocks_per_class: int = questionary_int('1クラスで作成するモザイク画のブロック数を入力してください: ')
else:
    print(f'1クラスで作成するモザイク画のブロック数 (設定ファイル): {blocks_per_class}')


# モザイク画データを読み込む
mosaic_data_lines: list[str] = read_text_file(os.path.join(path, mosaic_data_file))
cprint(f'\n\nモザイク画データを読み込みました。ファイル名: {mosaic_data_file}\n\n', 'light_blue')

# プロパティの入力2
start_row_total_pixels: str = questionary.select('合計ピクセルが記載されている開始行を選択してください: ', choices=mosaic_data_lines[-8:]).ask()
start_row_num: int = mosaic_data_lines.index(start_row_total_pixels)
end_row_total_pixels: str = questionary.select('合計ピクセルが記載されている終了行を選択してください', choices=[mosaic_data_lines[start_row_num].replace('\n', ' (selected) \n')] + mosaic_data_lines[start_row_num + 1 :]).ask()
end_row_num: int = mosaic_data_lines.index(end_row_total_pixels.replace(' (selected) \n', '\n'))

cprint(f'\n計算を開始します。', 'light_green')
time_start = time.time()

# モザイク画データからピクセル合計データを取得
pixel_sum_data_list: list = mosaic_data_lines[start_row_num: end_row_num + 1]
# ピクセル合計データを整形  改行を置換
pixel_sum_data: str = ''.join(pixel_sum_data_list).replace('\n', '  ')

# ピクセル合計データを辞書型に変換
pixel_sum_dict: dict[str, int] = {}
pairs: list[str] = pixel_sum_data.split()
for pair in pairs:
    key, value = pair.split('=')
    pixel_sum_dict[key] = int(value)

# 必要折り紙枚数の合計を計算
total_paper_dict: dict[str, int] = {key: math.ceil(value / pieces_per_origami) for key, value in pixel_sum_dict.items()}

# 結果をリストに格納
calc_result: list[str] = []
for key in total_paper_dict:
    calc_result.append(f'{key}: {pixel_sum_dict[key]}  ->  {total_paper_dict[key]}枚\n')

# 合計の2行は削除
mosaic_data_lines = mosaic_data_lines[:start_row_num] + mosaic_data_lines[end_row_num + 1:]

# ブロックごとのデータを取得
block_data_list: list[str] = []
for i, line in enumerate(mosaic_data_lines):
    if line.startswith('---') and i != 0:
        # 区切り文字を追加
        block_data_list.append('///')
        block_data_list.append(line)
    else:
        block_data_list.append(line)

# 区切り文字でリストを分割し再代入
block_data_list = ''.join(block_data_list).split('///')

# '='が含まれる行を取得
calc_data_list: list[str] = []
for block_data in block_data_list:
    lines = block_data.split('\n')
    for line in lines:
        if '=' in line:
            calc_data_list.append(line)
    # 区切り文字を追加
    calc_data_list.append('///')

# 区切り文字でリストを分割し再代入 最後の要素は削除
calc_data_list = ''.join(calc_data_list).split('///')[:-1]


# 各クラスごとのピクセル合計データを取得
calc_sum_class_dict: dict[str, int] = {}
calc_class_result: list[str] = []
class_number: int = 0
for i, line in enumerate(calc_data_list):
    # 辞書型に変換
    pixel_sum_class_dict: dict[str, int] = {}
    pairs: list[str] = line.split()
    for pair in pairs:
        key, value = pair.split('=')
        pixel_sum_class_dict[key] = int(value)
    calc_sum_class_dict = add_dictionaries(calc_sum_class_dict, pixel_sum_class_dict)

    # 各クラス、最初の行にクラス情報を追加 結果をリストに格納
    if (i+1)%blocks_per_class == 0:
        calc_class_result.append(f'--------------- {classes[class_number]} ---------------')
        for key in calc_sum_class_dict:
            calc_class_result.append(f'{key}: {calc_sum_class_dict[key]}   [ {math.ceil(calc_sum_class_dict[key] / pieces_per_origami)}枚 ]')
        calc_class_result.append('')
        calc_sum_class_dict = {}
        class_number += 1

calc_result = ''.join(calc_result) + '\n'*3 + '\n'.join(calc_class_result)

# 計算した結果をファイルに書き込む
write_text_file(os.path.join(path, 'output', f'{output_file_name}_origami.txt'), ''.join(calc_result))

time_end = time.time()
cprint(f'\n計算が完了しました。', 'light_blue')
cprint(f'ファイル名: {output_file_name}_origami.txt     計算時間: {time_count(time_start, time_end)}\n', 'light_blue')



cprint('\nデータを整形してCSVファイルに変換する処理を開始します。', 'light_green')
time_start = time.time()

summarized_result: list[str] = []
for block_data in block_data_list:
    lines = block_data.split('\n')[:-1]
    # '='が含まれる行は削除
    block_data = [line for line in lines if '=' not in line]
    # 1文字ずつカンマを挿入 (先頭行は除く)
    block_data = [insert_commas(line) if i != 0 else line for i, line in enumerate(block_data)]
    # 先頭行を最終行に移動
    block_data.append(block_data.pop(0))

    summarized_result.append('\n'.join(block_data))

# 調整したデータをcsvファイルに書き込む
write_text_file(os.path.join(path, 'output', f'{output_file_name}.csv'), '\n'.join(summarized_result))

time_end = time.time()
cprint(f'\nCSVファイルに変換しました。', 'light_blue')
cprint(f'ファイル名: {output_file_name}.csv     処理時間: {time_count(time_start, time_end)}\n', 'light_blue')


if EXCEL_DATA != None:
    # 全設定データがあるか確認、数値か確認
    settings_int = ['width', 'height', 'margin_top', 'margin_bottom', 'margin_left', 'margin_right', 'margin_header', 'margin_footer']
    settings = ['print_grid_line']
    is_settings = True
    for s in settings_int:
        if s not in EXCEL_DATA:
            is_settings = False
            cprint(f'Excel設定データに{s}がありません。', 'red')
        else:
            if not is_float(EXCEL_DATA[s]):
                is_settings = False
                cprint(f'Excel設定データの{EXCEL_DATA[s]}は整数値ではありません。', 'red')
            else:
                EXCEL_DATA[s] = float(EXCEL_DATA[s])
    for s in settings:
        if s not in EXCEL_DATA:
            is_settings = False
            cprint(f'Excel設定データに{s}がありません。', 'red')


    if is_settings:
        cprint('\nExcel設定データを読み込みました。', 'light_blue')
        cprint('Excelファイルを作成します。\n', 'light_green')
        time_start = time.time()

        # csvファイルを読み込み、Excelファイルに変換
        wb = xl.Workbook()
        wb.remove(wb.active)
        sheet = wb.create_sheet()

        with open(os.path.join(path, 'output', f'{output_file_name}.csv'), 'r', encoding="Shift-JIS") as f:
            reader = csv.reader(f)
            for row in reader:
                sheet.append(row)

        # 枠線を追加
        side = Side(style='thin', color='000000')
        border = Border(top=side, bottom=side, left=side, right=side)

        width = json_data['Excel']['width']
        height = json_data['Excel']['height']

        for row in sheet:
            # 行幅を調整
            sheet.row_dimensions[row[0].row].height = height
            for cell in row:
                if sheet[cell.coordinate].value:
                    if not sheet[cell.coordinate].value.startswith('---'):
                        sheet[cell.coordinate].border = border
                        # 中央揃え
                        sheet[cell.coordinate].alignment = xl.styles.Alignment(horizontal='center', vertical='center')
                    else:
                        # 上下中央揃え
                        sheet[cell.coordinate].alignment = xl.styles.Alignment(vertical='center')


        # 1行目の全てのセルを取得し、その幅を調整する
        for cell in sheet[1]:
            sheet.column_dimensions[cell.column_letter].width = width

        # ページレイアウトを設定
        # 印刷の向き: 横
        sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
        # 余白
        sheet.page_margins.top = json_data['Excel']['margin_top']
        sheet.page_margins.bottom = json_data['Excel']['margin_bottom']
        sheet.page_margins.left = json_data['Excel']['margin_left']
        sheet.page_margins.right = json_data['Excel']['margin_right']
        sheet.page_margins.header = json_data['Excel']['margin_header']
        sheet.page_margins.footer = json_data['Excel']['margin_footer']
        # 枠線を印刷
        if json_data['Excel']['print_grid_line'] == "True":
            sheet.print_options.gridLines = True
        elif json_data['Excel']['print_grid_line'] == "False":
            sheet.print_options.gridLines = False
        else:
            cprint('Excel設定データの print_grid_line が True または False ではありません。\n枠線印刷の設定は実行されませんでした\n', 'red')

        try:
            wb.save(os.path.join(path, 'output', f'{output_file_name}.xlsx'))
        except Exception as e:
            cprint(f'Excelファイルの保存に失敗しました。\n---------- エラーメッセージ ----------\n{e}', 'red')
            os._exit(0)
        else:
            time_end = time.time()
            cprint(f'Excelファイルを保存しました。', 'light_blue')
            cprint(f'ファイル名: {output_file_name}.xlsx     処理時間: {time_count(time_start, time_end)}', 'light_blue')

    else:
        cprint('Excel設定データの読み込みに失敗しました。Excelファイルを作成できませんでした。', 'red')
else:
    cprint('Excel設定データがありません。Excelファイルを作成できませんでした。', 'red')

cprint('\n--- プログラムを終了します ---', 'light_magenta')