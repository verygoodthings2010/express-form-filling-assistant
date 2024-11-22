import xlwings as xw
import PySimpleGUI as sg
import re
from typing import List, Dict

def extract_info(text: str) -> List[Dict[str, str]]:
    # 用于存储最终结果的列表
    final_results = []

    # 正则表达式
    phone_regex = r'^(13[0-9]|14[0-9]|15[0-9]|16[0-9]|17[0-9]|18[0-9]|19[8|9])\d{8}$|^(\d{3,4}-)?\d{7,8}$'  # 中国手机号码的通用格式
    weight_regex = r"([一二三四五六七八九十\d]+)斤"
    box_count_regex = r"([一二三四五六七八九十\d]+)箱"

    # 将文本分割成行
    data = text.strip().split('\n')

    # 处理数据中的每一行
    for entry in data:
        entry = entry.strip()
        if entry:
            # 将行按标点或空格分割
            parts = re.split(r'[，.、|/,\s。：:]+', entry)

            phone = None
            weight = None
            box_count = None
            address = ""
            name = None

            # 逐一分析每个部分的信息
            for part in parts:
                # 提取电话号码
                if re.match(phone_regex, part):
                    phone = part
                # 提取重量
                elif re.match(weight_regex, part):
                    weight = part
                # 提取箱子数量
                elif re.match(box_count_regex, part):
                    box_count = part
                # 提取姓名
                elif len(part) < 10 and not (part.isdigit() or "箱" in part or "斤" in part):
                    name = part
                # 收集地址
                else:
                    address += (part + " ") if address else part

            # 将提取的数据添加到 final_results
            final_results.append({
                'name': name,
                'phone': phone,
                'weight': weight,
                'number': box_count,  # 使用正确的数量键
                'address': address.strip()  # 去掉尾部空格
            })

    return final_results

sg.theme('DarkBlue')
# 创建用户输入界面布局
layout = [
    [sg.Text('选择Excel文件'), sg.FileBrowse(key='file')],
    [sg.Checkbox('从第一行写入', key='s1')],
    [sg.Text('收件人信息（每行一个）'), sg.Multiline(key='text', size=(45, 10))],
    [sg.Button('提交'), sg.Button('退出')]
]

# 创建窗口
window = sg.Window('自动填表程序', layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == '退出':
        break
    if event == '提交':
        text = values['text']
        file = values['file']
        results = extract_info(text)

        # 打开 Excel 工作簿
        wb = xw.Book(file)
        sheet = wb.sheets[0]
        if not values["s1"]:
            # 获取已使用范围的行数
            row_count = sheet.used_range.last_cell.row
            # 从下一行开始填充数据（假设第一行是表头）
            start_row = row_count + 1
        else:
            start_row = 2

        # 读取第一行的表头
        headers = sheet.range('A1:Z1').value  # 读取表头，假定为A-Z列
        header_to_index = {header: index + 1 for index, header in enumerate(headers)}  # 创建索引字典

        # 从第二行开始填充数据
        for row_index, person in enumerate(results, start=start_row):
            if '收件人手机' in header_to_index:
                sheet.range((row_index, header_to_index['收件人手机'])).value = person.get('phone')  # 填充手机号码
            if '卖家备注' in header_to_index:
                sheet.range((row_index, header_to_index['卖家备注'])).value = person.get('weight')  # 填充备注
            if '数量' in header_to_index:
                sheet.range((row_index, header_to_index['数量'])).value = person.get('number')  # 填充数量
            if '收件人地址' in header_to_index:
                sheet.range((row_index, header_to_index['收件人地址'])).value = person.get('address')  # 填充地址
            if '收件人姓名' in header_to_index:
                sheet.range((row_index, header_to_index['收件人姓名'])).value = person.get('name')  # 填充姓名
            elif '收件人电话' in header_to_index:
                sheet.range((row_index, header_to_index['收件人电话'])).value = person.get('phone')  # 填充手机号码

        sg.popup("填充完毕，识别可能有误差，请检查")