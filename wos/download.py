import time
from io import BytesIO

import pandas as pd
import requests

from .config import *


class WOS:
    def __init__(self, sid, qid: str, *, savefile, batch=1000):
        self.qid = qid
        self.sid = sid
        self.batch = batch
        self.savefile = savefile
        self.code = -1  # 0是正常 1是次数过多 2是其它错误 3是网络错误
        self.size = 1
        self.df = []

    def params(self, start, end):
        param = params.copy()
        param["markFrom"] = str(start)
        param["markTo"] = str(end)
        param["parentQid"] = self.qid
        return param

    def __download__(self, param) -> Status:
        response = requests.post(url, json=param, headers={"X-1p-Wos-Sid": self.sid})
        if response.status_code == 500:
            return Status.LIMIT

        if response.status_code != 200:
            return Status.ERROR

        df = pd.read_excel(BytesIO(response.content))
        if self.size == 1:
            self.size = df.shape[0]
        else:
            self.size += df.shape[0]
        self.df.append(df)
        return Status.NORMAL

    def run(self, count: int):
        while self.size < count:
            param = self.params(self.size, self.size + self.batch - 1)
            self.code = self.__download__(param)

            if self.code == Status.LIMIT:
                for _ in range(10):
                    time.sleep(1)
                continue

            if self.code == Status.ERROR:
                break

    def save(self):
        if len(self.df):
            df = pd.concat(self.df, axis=0, ignore_index=True)
            df.to_excel(self.savefile, index=False)
