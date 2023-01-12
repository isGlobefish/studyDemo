import json
import pandas as pd

data = pd.read_excel('C:/Users/Zeus/Desktop/json123.xlsx', header=0, sheet_name=0)

content = []
for irow in range(data.shape[0]):
    content.append(
        (
            {
                "PARTNER": str(data.loc[irow, '客户ID']),
                "ZQDCY_FROM": str(data.loc[irow, '门店名称']),
                "ZQDCY": "",
            }
        )

    )

# 字典中的数据都是单引号,但是标准的json需要双引号
Json = json.dumps(content, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ':'))
# 可读可写, 如果不存在则创建, 如果有内容则覆盖
with open("C:/Users/Zeus/Desktop/jsonTest.json", "w+", encoding='utf-8') as js:
    js.write(Json)
    js.close()