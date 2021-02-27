#!/usr/bin/env python3
import os
import openpyxl
import sys
import yaml

import postgresclient

# 引数判定
if len(sys.argv) != 4:
    print("Usage: python " & os.path.basename(__file__) & " <format_file> <output_file_path> <yyyymm>")
    sys.exit(1)

# フォーマット定義読み込み
with open(sys.argv[1], "r") as formatfile:
    format_config = yaml.load(formatfile)

# DB接続情報
l_dirname = os.path.dirname(os.path.abspath(__file__))
db_config = os.path.normpath(os.path.join(l_dirname, "../db_connection.yaml"))
db_client = postgresclient.PostgresClient(yaml.load(db_config["postgres"]))
# SQLファイル配置場所
sql_dir = os.path.normpath(os.path.join(l_dirname, "../sql"))

# ---- 出力ファイル作成処理 ----
# フォーマット確認(TODO: csv増える予定)
if format_config.get("format") not in ["xlsx"]:
    print(format_config.get("format") " is not valid format.")
    sys.exit(1)

# -- xlsx出力処理 --
# ベースになるファイル読み込み
basefile = format_config.get("basefile")
if not os.path.isfile(basefile):
    print("File could not be found " + basefile + ".")
    sys.exit(1)

wb = openpyxl.load_workbook(sys.argv[1])

for sheet in format_config["sheets"]:
    if sheet["name"] not in wb.sheetnames:
        wb.create_sheet(index=sheet["index"], title=sheet["name"])
    ws = wb.worksheets[sheet["name"]]

    row_i = sheet["row_padding"] + 1
    col_i = sheet["col_padding"] + 1

    # column headers
    col_headers = []
    col_header_confs = sorted(sheet["col_headers"], key=lambda x: x["index"])
    for col_header_conf in col_header_confs:
        ds = col_header_conf["source"]
        h_table = db_client.execute(os.path.normpath(os.path.join(sql_dir, ds["sql"])))
        h_table = sorted(h_table, key=lambda x: x[ds["order"]])
        col_header = {
            size: len(h_table)
            datas: [x[ds["data"]]] for x in h_table]
        }
    
    
    # sql.format(datas)



#Completed.
print("Successfully generated.")
sys.exit(0)
