"""
データの位置を修正、標準偏差を求めるプログラム
"""
import openpyxl
import time
import os
import csv
from tkinter import Tk, filedialog
#標準偏差を求めるためのモジュールをインポート
import statistics
import re
from datetime import datetime
import xlwings as xw


def collect_zone_change_indices(sheet):
    """ゾーン変化箇所の行番号を取得する関数"""
    column_data = {col: [] for col in ["A", "B", "C", "D", "E", "F", "I", "J"]}
    max_row = sheet.max_row
    non_empty_rows = 0
    # データ取得
    for row in range(2, max_row + 1):
        for col in column_data.keys():
            cell_value = sheet[f"{col}{row}"].value
            column_data[col].append(cell_value)
            if cell_value is not None:
                 has_data= True
        if has_data:
            non_empty_rows += 1
    # ゾーン変化検出
    zone_data = column_data["E"]
    change_indices = []

    for i in range(1, len(zone_data)):
        if zone_data[i] != zone_data[i - 1]:
            change_indices.append(i + 1)  # Excelの行番号に対応
            # if len(change_indices) == 1:#実空間に合わせてデータの数は1つだけにする
            #     break

    return column_data, change_indices,non_empty_rows

def save_data(sheet, column_data, shift_diff, target_change_index, sampling_interval, measurement_count):
    """ゾーン変化箇所の差分に応じてデータをシフトし、Excelに書き込む"""
    max_row = len(column_data["A"]) + 2 +shift_diff

    # データシフト（元のデータは消去せずに移動）
    for row in range(max_row - 1, 1, -1):  # 逆順でシフト
        for col, data in column_data.items():
            new_row = row + shift_diff
            if 2 <= new_row < max_row:  # 範囲内に収まる場合
                sheet[f"{col}{new_row}"] = sheet[f"{col}{row}"].value
            if shift_diff > 0:  # シフト後の余白部分を空白にする
                if sheet[f"{col}{row}"].value is not None:  # 元のセルに値がある場合のみ処理
                    sheet[f"{col}{row}"] = None  # セルを空白にする
                    #if col == "A":  # 1回だけK列に1を入れる
                    if row <= shift_diff+1:
                        sheet[f"K{row}"] = 1

def process_all_sheets(file_path, write_path):
    """1/2 全てのシートを処理する関数"""
    wb = openpyxl.load_workbook(file_path)
    sheet_names = [name for name in wb.sheetnames if name != "0"]
    all_change_indices = {}  # 全てのゾーン変化箇所行番号を保持（シート名と共に）
    sheet_data = {}  # 各シートのデータとゾーン変化箇所を保持

    # 最初に読み込んだシートから測定数とサンプリング間隔を取得
    first_sheet = wb[sheet_names[0]]
    measurement_count = int(first_sheet["H2"].value)
    sampling_interval = float(first_sheet["H3"].value)

    sheet_names = [name for name in wb.sheetnames if name != "0"]
    num = len(sheet_names)   
    # データ収集フェーズ
    for sheet_index, sheet_name in enumerate(sheet_names, start=1):
        if sheet_index > num:
            break
        sheet = wb[sheet_name]
        print(f'Analyzing {sheet_name}...')
        try:
            column_data, change_indices, non_empty_rows = collect_zone_change_indices(sheet)#関数の呼び出し
            print(f"{sheet_name} のデータを収集しました。データ数は{non_empty_rows}行です。")

            if change_indices:
                max_change_index = min(change_indices)  # 最大のゾーン変化箇所
                all_change_indices[sheet_name] = max_change_index
                sheet_data[sheet_name] = (sheet, column_data, max_change_index)

                print(f"ゾーン変化箇所: {change_indices}")
                print(f"{sheet_name} のゾーン変化箇所で一番小さい行番号: {max_change_index}")
        except Exception as e:
            print(f'Error analyzing {sheet_name}: {e}')

    # ゾーン変化箇所のランキングを表示
    sorted_change_indices = sorted(all_change_indices.items(), key=lambda x: x[1])
    print(f"ゾーン変化箇所のランキング: {sorted_change_indices}")

    if len(sorted_change_indices) >= 100:
        target_change_index = sorted_change_indices[99][1]#
        target_sheet_name = sorted_change_indices[99][0]
        print(f"100番目のゾーン変化箇所の行番号: {target_change_index}")
    else:
        target_change_index = sorted_change_indices[-1][1]  # 100個未満の場合は最大のものを使用
        target_sheet_name = sorted_change_indices[-1][0]
        print(f"100個未満のため、最大のゾーン変化箇所の行番号: {target_change_index}")

    # 101番目以降のシートを削除
    if len(sorted_change_indices) > 100 or sheet_name !="0":#0のシートは削除しない
        for sheet_name, _ in sorted_change_indices[100:]:
            wb.remove(wb[sheet_name])
            print(f"シート {sheet_name} を削除しました。")

    # シフトと書き込みフェーズ
    for sheet_name, (sheet, column_data, max_change_index) in sheet_data.items():
        shift_diff = target_change_index - max_change_index
        if shift_diff > 0:
            print(f"{sheet_name} のデータを {shift_diff} 行下にシフトします。")
            save_data(sheet, column_data, shift_diff, target_change_index, sampling_interval, measurement_count)#関数の呼び出し

    # 上書き保存（タイムスタンプ付きファイル名）

    """位置を動かしたファイル(file_path)を保存"""
    base_name = os.path.basename(file_path)  # ファイル名のみを抽出 データを読み込むのに使ったファイル名を使う
    updated_time = datetime.now().strftime("%Y%m%d_%H%M%S")  # 新しい時間
    new_name = re.sub(r'(\d{8}_\d{6})', updated_time, base_name)  # 古い時間部分を新しい時間に置き換え
    new_file_path = os.path.join(os.path.dirname(file_path), f"a0_{new_name}")  # "devi_" を追加 標準偏差を求める
    #new_file_path = file_path.replace(".xlsx", f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


    print("saving...")
    wb.save(new_file_path)
    wb.close()  # 読み込み用のワークブックを閉じる
    print(f'Saved to {new_file_path}')
    print(f"全シートの最大ゾーン変化箇所の行番号: {target_change_index}")


    """ 2/2 write_path(記録を残すようのファイル)への記録とCSVファイルへの出力"""
        # シート名が "0" のシートに target_change_index と target_sheet_name を保存
    try:
        app = xw.App(visible=False)  # Excelアプリケーションを非表示で起動
        wb = app.books.open(write_path)


        sheet_0 = wb.sheets("0")
        sheet_0.range("G8").value = "ゾーンが変化した時間[ms]"
        sheet_0.range("G9").value = "100個のデータがそろう行番号"
        sheet_0.range("H8").value= target_sheet_name
        sheet_0.range("H9").value = target_change_index
        
        # E列に0を入れ、その前後にサンプリング間隔ごとに値を増減させる
        for i in range(2, measurement_count + 1):#1はインデックスなのでパス
            if i == target_change_index:
                sheet_0.range(f"E{i}").value = 0
            elif i < target_change_index:
                sheet_0.range(f"E{i}").value = -(target_change_index - i) * sampling_interval
            else:
                sheet_0.range(f"E{i}").value = (i - target_change_index) * sampling_interval
        
        """write_pathの名前を変更"""
        timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")  # 新しい時間
        new_write_path = f"{write_path.rsplit('.', 1)[0]}_{timestamp}.xlsx"
        
        
        # base_name = os.path.basename(write_path)  # ファイル名のみを抽出 データを読み込むのに使ったファイル名を使う
        # updated_time = datetime.now().strftime("%Y%m%d_%H%M%S")  # 新しい時間
        # new_name = re.sub(r'(\d{8}_\d{6})', updated_time, base_name)  # 古い時間部分を新しい時間に置き換え
        # new_write_path = os.path.join(os.path.dirname(write_path), f"a1_{new_name}") 
        
        
        wb.save(new_write_path)
        wb.close()  # 書き込み用のワークブックを閉じる
    finally:
        app.quit()  # Excelアプリケーションを終了
        print(f"データが '{new_write_path}' に保存されました。")


    # ランキングをCSVファイルに出力
    ranking_file_path = os.path.join(os.path.dirname(file_path), f"ranking_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    with open(ranking_file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["ランキング", "シート名", "ゾーン変化箇所"])
        for rank, (sheet_name, change_index) in enumerate(sorted_change_indices, start=1):
            writer.writerow([rank, sheet_name, change_index])
    print(f'ランキングを {ranking_file_path} に保存しました。')
    print(new_file_path)
    print(write_path)

    return  new_file_path, new_write_path
    #return new_file_path, new_write_path, ranking_file_path#将来的には必要になるかもしれないので、戻り値を追加する

""" 標準偏差を求めるプログラム"""

def calculate_std_dev(new_file_path, write_path,column_letter='F'):
    start_time = time.time()
    print(new_file_path)
    print(write_path)
        # Excelファイルを読み込む
    try:
        app = xw.App(visible=False)  # Excelアプリケーションを非表示で起動
        workbook = app.books.open(new_file_path)
        writebook=app.books.open(write_path)
                
        sheet0 = writebook.sheets['0']
        
        # '0'以外のすべてのシートを処理対象として取得
        sheets_to_process = [workbook.sheets[sheet_name] for sheet_name in workbook.sheets if sheet_name.name != '0']
        if not sheets_to_process:
            print("処理するシートが見つかりません。ワークブックを確認してください。")
            app.quit()
            return
        print(f"処理対象のシート数: {len(sheets_to_process)}")
        
        # # 未使用シートを記録するリスト
        # unused_sheets = {sheet.title: True for sheet in sheets_to_process}  # 初期値として全シートを未使用とみなす
        total_used_data_count = 0  # 使用されたデータ数を記録する変数
        # 任意のシートからG2-G5およびH2-H5のデータをシート '0' にコピー

        """ 以下の3行は一旦コメントアウト 測定条件を書き込み用のシートに書き込む"""
        # for i in range(2, 6):
        #     sheet0[f'G{i}'] = sheets_to_process[0][f'G{i}'].value  # 最初のシートからコピー
        #     sheet0[f'H{i}'] = sheets_to_process[0][f'H{i}'].value


        # 各行の処理を行い、シート '0' の指定列に標準偏差と平均を計算して書き込み
        row = 2
        #total_data_count = 0  # データの総数をカウント
        flag=0
        while True:
            # 処理対象シートの指定列からデータを収集
            data = []
            for sheet in sheets_to_process:
                shift_value=sheet.range(f'K{row}').value
                cell_value = sheet.range(f'{column_letter}{row}').value
                if shift_value != 1 and cell_value is not None:  # shift_valueが1でない場合はデータを収集
                    data.append(cell_value)
                    #unused_sheets[sheet.title] = False  # データが使用されたシートを「未使用ではない」と記録
    
            # データがない場合、終了
            if not data:
                sheet0.range('A2').value = len(sheets_to_process)  # 処理したシート数を記録
                break
            
            # 標準偏差と平均を計算
            std_dev = statistics.stdev(data) if len(data) > 1 else 0
            mean_value = statistics.mean(data) if data else 0
            print(f"Row {row}: 標準偏差 = {std_dev}, 平均 = {mean_value}")# デバッグ用
            
            # シート '0' の指定セルに標準偏差を書き込む
            sheet0.range(f'{column_letter}{row}').value = std_dev
            
            # 平均をJ列に記録
            sheet0.range(f'J{row}').value = mean_value
            
            # 使用したデータ数をI列に記録
            data_count = len(data)
            if data_count == 100 and flag==0:
                sheet0.range('H10').value=row-1  # 100個のデータがそろう時間[ms]を記録
                flag=1
            sheet0.range(f'I{row}').value = data_count
            total_used_data_count += data_count
            
            # 次の行へ
            row += 1
        
        base_name = os.path.basename(write_path)  # ファイル名のみを抽出 データを読み込むのに使ったファイル名を使う
        updated_time = datetime.now().strftime("%Y%m%d_%H%M%S")  # 新しい時間
        new_name = re.sub(r'(\d{8}_\d{6})', updated_time, base_name)  # 古い時間部分を新しい時間に置き換え
        new_write_path = os.path.join(os.path.dirname(write_path), f"stdev_{new_name}")  # "devi_" を追加
        
        writebook.save(new_write_path)  # write_pathに保存
        writebook.close()  # 書き込み用のワークブックを閉じる
        workbook.close()  # 読み込み用のワークブックを閉じる
    finally:
        app.quit()  # Excelアプリケーションを終了

        end_time = time.time()
        print(f"処理時間: {end_time - start_time:.2f}秒")
        print(f"データが '{new_write_path}' に保存されました。標準偏差はシート '0' の {column_letter} 列の2行目から書き込みました。")


"""メイン関数に相当する部分"""
# Step1 位置合わせ
Tk().withdraw()
print("位置合わせに使うファイルを選んでください")
file_path = filedialog.askopenfilename(title="測定結果をまとめたエクセルファイルを選択", filetypes=[("Excel files", "*.xlsx")])
print("記録を残すファイルを選んでください")
write_path = filedialog.askopenfilename(title="標準偏差を書き込むエクセルファイルを選択", filetypes=[("Excel files", "*.xlsx")])

start_time = time.time()#開始時間
new_file_path, new_write_path=process_all_sheets(file_path,write_path)#関数の呼び出し
end_time = time.time()#終了時間
print(f'Process all sheets is done. Elapsed time: {end_time - start_time:.2f} seconds.')

# Step2 標準偏差を求める
# Tk().withdraw()
# print("計算に使うファイルを選んでください")#読み込むときにファイル名を変えたものを読み込ませておく必要がある、、書き換えたものをどういう風にして置き換えればいいのか
# new_file_path = filedialog.askopenfilename(title="位置を修正したエクセルファイルを選択", filetypes=[("Excel files", "*.xlsx")])#最終的にはここも無くしたい

#print("記録を残すファイルを選んでください")
#write_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
print ("標準偏差を求めます")
calculate_std_dev(new_file_path, new_write_path,column_letter='F')#標準偏差を求める関数
print(new_file_path)
print(write_path)
