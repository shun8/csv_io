#!/usr/bin/env python3
import os
import sys
import yaml

import postgresclient
import xlsxgenerator

# 引数判定
if len(sys.argv) != 5:
    print("Usage: python " + os.path.basename(__file__) + " <format_file> <output_file_path> <from_date> <to_date>")
    sys.exit(1)

p_format_file = sys.argv[1]
p_output_file_path = sys.argv[2]
p_from_date = sys.argv[3]
p_to_date = sys.argv[4]

# DB接続情報
l_dirname = os.path.dirname(os.path.abspath(__file__))
with open(os.path.normpath(os.path.join(l_dirname, "../db_connection.yaml")), "r") as db_config:
    l_db_config = yaml.safe_load(db_config)
    db_client = postgresclient.PostgresClient(l_db_config["postgres"])
# SQLファイル配置場所
sql_dir = os.path.normpath(os.path.join(l_dirname, "../sql"))

l_xlsxgenerator = xlsxgenerator.XLSXGenerator(db_client, sql_dir)
wb = l_xlsxgenerator.gen_xlsx_like_csv(
        p_format_file, {"from_date": p_from_date, "to_date": p_to_date})
wb.save(p_output_file_path)

#Completed.
print("Successfully exported.")
sys.exit(0)
