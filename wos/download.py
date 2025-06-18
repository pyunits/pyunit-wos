import time
from io import BytesIO

import pandas as pd
import requests

url = "https://webofscience.clarivate.cn/api/wosnx/indic/export/saveToFile"

params = {
    "displayTimesCited": "true",
    "displayCitedRefs": "true",
    "product": "UA",
    "colName": "WOS",
    "displayUsageInfo": "true",
    "fileOpt": "xls",
    "action": "saveToExcel",
    "locale": "en_US",
    "view": "summary",
    "filters": "fullRecord",
    "sortBy": "relevance",
    "isRefQuery": "false",
}


class WOS:
    def __init__(self, sid, qid: str, *, batch=1000):
        self.qid = qid
        self.sid = sid
        self.batch = batch
        self.code = -1  # 0是正常 1是次数过多 2是其它错误 3是网络错误
        self.size = 1
        self.df = []

    def params(self, start, end):
        param = params.copy()
        param["markFrom"] = str(start)
        param["markTo"] = str(end)
        param["parentQid"] = self.qid
        return param

    def __download__(self, param) -> int:
        """
        :return: 0是正常 1是次数过多 2是其它错误 3是网络错误
        """
        response = requests.post(url, json=param, headers={"X-1p-Wos-Sid": self.sid})
        if response.status_code == 200:
            df = pd.read_excel(BytesIO(response.content))
            size = df.shape[0]
            if self.size == 1:
                self.size = size
            else:
                self.size += size
            self.df.append(df)
            return 0
        elif response.status_code == 500:
            try:
                err = response.json()["errm"]
                if err.startswith("Verifiable request limit"):
                    return 1
            except Exception as e:
                return 2
        return 3

    def run(self, count: int):
        while self.size < count:
            param = self.params(self.size, self.size + self.batch-1)
            self.code = self.__download__(param)
            if self.code == 1:
                for _ in range(10):
                    time.sleep(1)

    def save(self, file: str):
        df = pd.concat(self.df, axis=0, ignore_index=True)
        df.to_excel(file, index=False)
