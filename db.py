"""SQLite persistence for verification results."""

import sqlite3
import threading
from datetime import datetime

from config import DB_FILE


def key(o):
    return "|".join([o["PLC"], o["Area"], o["Equipment"]])


class DB:
    def __init__(self):
        self.c = sqlite3.connect(DB_FILE, check_same_thread=False)
        self.c.execute("""CREATE TABLE IF NOT EXISTS verification
        (k TEXT PRIMARY KEY, plc TEXT, area TEXT, equipment TEXT, template TEXT,
         result TEXT, cause TEXT, comment TEXT, tester TEXT, timestamp TEXT)""")
        self.c.commit()
        self.lock = threading.Lock()

    def get(self, k):
        with self.lock:
            r = self.c.execute("SELECT result,cause,comment,tester,timestamp FROM verification WHERE k=?", (k,)).fetchone()
        return r

    def save(self, o, result, cause, comment, tester):
        k = key(o)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.c.execute("""INSERT INTO verification VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(k) DO UPDATE SET result=excluded.result,cause=excluded.cause,
            comment=excluded.comment,tester=excluded.tester,timestamp=excluded.timestamp""",
            (k, o["PLC"], o["Area"], o["Equipment"], o["Template"], result, cause, comment, tester, now))
            self.c.commit()
