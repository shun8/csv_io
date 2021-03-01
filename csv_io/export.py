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

p_format_file = sys.argv[1]
p_output_file_path = sys.argv[2]
p_yyyymm = sys.argv[3]

# フォーマット定義読み込み
with open(p_format_file, "r") as format_file:
    format_config = yaml.load(format_file)

# DB接続情報
l_dirname = os.path.dirname(os.path.abspath(__file__))
db_config = os.path.normpath(os.path.join(l_dirname, "../db_connection.yaml"))
db_client = postgresclient.PostgresClient(yaml.load(db_config["postgres"]))
# SQLファイル配置場所
sql_dir = os.path.normpath(os.path.join(l_dirname, "../sql"))

# アドレス計算用関数
def calc_header_span(p_index, p_col_headers):
    if p_index >= len(p_col_headers):
        return 1
    return p_col_headers[p_index+1]["size"] * calc_col_index(p_index+1, p_col_headers)

# ---- 出力ファイル作成処理 ----
# フォーマット確認(TODO: csv増える予定)
if format_config.get("format") not in ["xlsx"]:
    print(format_config.get("format") " is not valid format.")
    sys.exit(1)

# -- xlsx出力処理 --
# ベースになるファイル読み込み
basefile = format_config.get("basefile")
if basefile is None:
    wb = openpyxl.Workbook()
elif os.path.isfile(basefile):
    wb = openpyxl.load_workbook(p_output_file_path)
else:
    print("File could not be found " + basefile + ".")
    sys.exit(1)

# create sheet
for sheet in format_config["sheets"]:
    if sheet["name"] not in wb.sheetnames:
        wb.create_sheet(index=sheet["index"], title=sheet["name"])
    ws = wb.worksheets[sheet["name"]]

    row_i = sheet["row_padding"] + 1
    col_i = sheet["col_padding"] + 1

    # query column headers
    col_headers = []
    col_header_confs = sorted(sheet["col_headers"], key=lambda x: x["index"])
    for col_header_conf in col_header_confs:
        ds_conf = col_header_conf["source"]
        with open(os.path.normpath(os.path.join(sql_dir, ds_conf["sql"])), "r") as l_sqlfile:
            l_sql = l_sqlfile.read()
        l_table = db_client.execute(l_sql.format(p_yyyymm))
        l_table = sorted(l_table, key=lambda x: x[ds_conf["order"]])
        l_col_h = {
            size: len(l_table),
            indexes: {x[ds_conf["data"]]: x[ds_conf["order"]] for x in l_table}
        }
        col_headers.append(l_col_h)
    for i in range(len(col_headers)):
        col_headers[i]["span"] = calc_header_span(i, col_headers)

    # draw column headers
    for i in range(len(col_headers)):
        l_span = col_headers[i+1]["size"] if i+1 < len(col_headers) else 1
        l_col_i = col_i
        for h_txt in col_headers[i].keys():
            # ↑のkeys() 3.6以降は順序が保存されるらしいのでそれ前提
            ws.cell(row_i, l_col_i, value=h_txt)
            l_col_i += l_span
            # col_header_confs["cell_style"]
        # col_header_confs["row_style"]
        # col_header_confs["last_col_style"] # colnum: l_col_i -1, right side
        row_i++

    # query bodies
    bodies = []
    body_confs = sorted(sheet["bodies"], key=lambda x: x["index"])
    for body_conf in body_confs:
        ds_conf = body_conf["source"]
        rh_ds_conf = ds_conf["group"]["row_header"]["source"]
        with open(os.path.normpath(os.path.join(sql_dir, rh_ds_conf["sql"])), "r") as l_sqlfile:
            l_sql = l_sqlfile.read()
        l_rh_table = db_client.execute(l_sql.format(p_yyyymm))
        l_rh_table = sorted(l_rh_table, key=lambda x: x[rh_ds_conf["order"]])
        l_rh_indexes = {x[rh_ds_conf["data"]]: x[rh_ds_conf["order"]] for x in l_rh_table}

        with open(os.path.normpath(os.path.join(sql_dir, ds_conf["sql"])), "r") as l_sqlfile:
            l_sql = l_sqlfile.read()
        l_table = db_client.execute(l_sql.format(p_yyyymm))
        l_ch_columns = ds_conf["group"]["col_headers"]
        l_ch_columns = sorted(l_ch_columns, key=lambda x: x["header_index"])
        l_body = {
            table: l_table,
            rh_indexes: l_rh_indexes,
            rh_column_name: rh_ds_conf = ds_conf["group"]["row_header"]["column_name"]
            ch_column_names: [x["column_name"] for x in l_ch_columns]
        }
        bodies.append(l_boby)

    # draw bodies
    for body in bodies:
        row_header_span = sheet["row_header_span"]
        for l_row in boby["table"]:
            l_row_i = row_i + int(rh_indexes[l_row[body["rh_column_name"]]])
            l_column_i = row_header_span + 1
            for i in range(len(col_headers)):
                l_h_idx = int(col_headers[i]["indexes"][body["table"][body["ch_column_names"][i]]])
                l_column_i += l_h_idx * col_headers[i]["span"]
            ws.cell(l_row_i, l_column_i, value=l_low[ds_conf["data"]])
            # ds_conf["cell_style"]
        # draw row headers
        for h_txt in body["rh_indexes"].keys()
            ws.cell(row_i, row_header_span, value=h_txt)
            # ds_conf["header_style"]
            # ds_conf["row_style"]
            row_i++

wb.save(p_output_file_path)
#Completed.
print("Successfully exported.")
sys.exit(0)
