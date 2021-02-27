#!/usr/bin/env python3
import os
import psycopg2
import yaml

class PostgresClient:

    def __init__(self, config):
        self.DBNAME = config["dbname"].strip("/")
        self.HOST = config["host"]
        self.PORT = config["port"] 
        self.USER = config["user"]
        self.PASSWORD = config["password"]

    def execute(self, sql):
        conn = psycopg2.connect(
            dbname=self.DBNAME,
            user=self.USER,
            password=self.PASSWORD,
            host=self.HOST,
            port=self.PORT
        )
        cur = conn.cursor()

        cur.execute(sql)
        result = cur.fetchall()

        cur.close()
        conn.close()

        return result
