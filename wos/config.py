from enum import Enum

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


class Status(Enum):
    NORMAL = 0  # 正常
    LIMIT = 1  # 次数过多
    ERROR = 2  # 其它错误
