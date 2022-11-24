#!/usr/bin/env python3
import pyodbc

class SQLServerClient:

    def __init__(self, config):
        self.DBNAME = config["dbname"].strip("/")
        self.HOST = config["host"]
        self.PORT = config["port"] 
        self.USER = config["user"]
        self.PASSWORD = config["password"]
        self.CONNECTION_STR = "DRIVER={SQL Server};SERVER=" + self.HOST + ";PORT=" + self.PORT + ";DATABASE=" + self.DBNAME + ";UID=" + self.USER + ";PWD=" + self.PASSWORD

    def execute(self, sql, params={}):
        # connection保持しておかない実装
        conn = pyodbc.connect(self.CONNECTION_STR)
        # 変数名でパラメータ置換(psycopg2との互換性)
        for k, v in params.items():
            sql = sql.replace("%(" + k + ")s", "'" + v + "'")
            sql = sql.replace("%(" + k + ")", v)

        # dict形式で結果を取得
        cur = conn.cursor()
        cur.execute(sql)

        dict_result = [dict(zip([column[0] for column in cur.description], row)) for row in cur.fetchall()]

        cur.close()
        conn.close()
        return dict_result
