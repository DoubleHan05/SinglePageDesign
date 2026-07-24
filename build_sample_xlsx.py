"""生成示例 Excel（KM360N 数据）—— sheet1 产品数据 / sheet2 参数表 / sheet3 使用说明"""
from pathlib import Path
import openpyxl

HERE = Path(__file__).parent
out = HERE / "samples" / "KM360N.xlsx"
out.parent.mkdir(parents=True, exist_ok=True)

wb = openpyxl.Workbook()

# Sheet1：产品数据
ws1 = wb.active
ws1.title = "产品数据"
sheet1_rows = [
    ("型号", "KM360N"),
    ("产品图", "assets/printer.svg"),
    ("细节图1", "assets/printer.svg"),
    ("细节图2", "assets/printer.svg"),
    ("特点1", "免安装驱动、异地远程打印、多人跨端打印、支持多个系统"),
    ("特点2", "普通用户：番茄标签 + 快麦云打印机"),
    ("特点3", "第三方软件用户：快麦云打印机无需开发支持所有软件"),
    ("特点4", "企业开发者：支持 SaaS 对接、免费 API 接口"),
    ("场景1", "仓|仓储"),
    ("场景2", "物|物流"),
    ("场景3", "零|零售"),
]
for r in sheet1_rows:
    ws1.append(r)
ws1.column_dimensions["A"].width = 32
ws1.column_dimensions["B"].width = 60

# Sheet2：参数表
ws2 = wb.create_sheet("参数表")
ws2.append(["参数名称", "参数值"])
params = [
    ("电池寿命", ""),
    ("充放电次数", ""),
    ("指令集", "ESC\\POS、TSPL"),
    ("打印内容", "英文、中文简体、中文繁体、图片、字符、线条、符号、一维条码、二维码"),
    ("编码方式", "GBK、BIG5、KSC5601、UTF-8"),
    ("系统支持操作系统", "Windows、iOS、Android、Linux、MacOS"),
    ("对接平台/软件", "番茄标签、快递助手、风火递等"),
    ("固件升级方式", "支持 USB 本地升级、OTA 升级"),
    ("设备状态上报", "支持主动上报打印机状态（在线、离线）"),
    ("作业状态上报", "支持"),
    ("设备日志获取", "不支持"),
    ("切纸方式", "手动撕纸"),
    ("开盖方式", "上翻盖"),
    ("纸仓形式", "有纸仓"),
    ("出纸方式", "前出纸"),
    ("配件", "电源适配器、数据线、说明书（含保修卡）"),
    ("工作温度/湿度", "温度：0-45℃；湿度：20-90%"),
    ("存储温度/湿度", "温度：-10-60℃；湿度：10-90%"),
    ("纸张回退", "支持"),
    ("自动进纸", "支持"),
    ("认证标准", "3C认证"),
    ("目标市场/场景", "连锁餐饮、连锁商超、连锁门店"),
]
for r in params:
    ws2.append(r)
ws2.column_dimensions["A"].width = 32
ws2.column_dimensions["B"].width = 60

wb.save(out)
print("wrote", out)
