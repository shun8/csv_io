#!/usr/bin/env python3
import psycopg2
import psycopg2.extras

class PostgresClient:

    def __init__(self, config):
        self.DBNAME = config["dbname"].strip("/")
        self.HOST = config["host"]
        self.PORT = config["port"] 
        self.USER = config["user"]
        self.PASSWORD = config["password"]

    def execute(self, sql, params={}):
        # connection保持しておかない実装
        conn = psycopg2.connect(
            dbname=self.DBNAME,
            user=self.USER,
            password=self.PASSWORD,
            host=self.HOST,
            port=self.PORT
        )
        # dict形式で結果を取得: https://qiita.com/itoufo/items/7306122497fd4f712bff
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        cur.execute(sql, params)
        results = cur.fetchall()

        dict_result = []
        for row in results:
            dict_result.append(dict(row))

        cur.close()
        conn.close()
        return dict_result
