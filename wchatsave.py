# -*- coding: utf-8 -*-
# @author: vincentlz
# @file: wchatsc.py
# @time: 2023/1/29 13:30

"""
文件功能&使用说明：

"""
import pandas as pd
import uiautomation as auto
from openpyxl.workbook import Workbook

def collectTOP(collect):
    last = None
    # 点击列表组件左上角空白部分，激活右窗口
    collect.Click(10, 10)
    while 1:
        item = collect.GetFirstChildControl().TextControl()
        print(item.Name)
        item.SendKeys("{PAGEUP 10}")
        if last == item.Name:
            break
        last = item.Name

def read_collect_info(ico_tag):
    # 根据按钮相对位置找到文本节点的位置
    desc_tags = ico_tag.GetNextSiblingControl().GetNextSiblingControl().GetChildren()
    content = "\n".join([desc_tag.Name for desc_tag in desc_tags])
    time_tag, source_tag = ico_tag \
        .GetParentControl().GetParentControl() \
        .GetNextSiblingControl().GetNextSiblingControl() \
        .GetChildren()[:2]
    time, source = time_tag.Name, source_tag.Name
    return content, time, source


def get_meau_item(name):
    menu = wechatWindow.MenuControl()
    menu_items = menu.GetLastChildControl().GetFirstChildControl().GetChildren()
    for menu_item in menu_items:
        if menu_item.ControlTypeName != "MenuItemControl":
            continue
        if menu_item.Name == name:
            return menu_item


wechatWindow = auto.WindowControl(
    searchDepth=1, Name="微信", ClassName='WeChatMainWndForPC')
button = wechatWindow.ButtonControl(Name='收藏')
button.Click()
button = wechatWindow.ButtonControl(Name='链接')
button.Click()
collect = wechatWindow.ListControl(Name='收藏')
# collectTOP(collect)

result = []
last = None
while 1:
    page = []
    text = collect.GetFirstChildControl().TextControl().Name
    if text == last:
        # 首元素无变化说明已经到底部
        break
    last = text
    items = collect.GetChildren()
    for i, item in enumerate(items):
        # 每个UI列表元素内部都必定存在左侧的图标按钮
        ico_tag = item.ButtonControl()
        content, time, source = read_collect_info(ico_tag)
        # 对首尾位置的节点作偏移处理
        item.RightClick(y=-5 if i == 0 else 5 if i == len(items)-1 else None)
        get_meau_item("复制地址").Click()
        url = auto.GetClipboardText()
        row = [content, time, source, url]
        print("\r", row, end=" "*1000)
        page.append(row)
    result.extend(page)
    item.SendKeys("{PAGEDOWN}")

df = pd.DataFrame(result, columns=["文本", "时间", "来源", "url"])
df.drop_duplicates(inplace=True)
print("微信收藏的链接条数：",df.shape[0])
df.head()

df[["标题", "描述"]] = df.文本.str.split(r"\n", 1, True)
df = df[["标题", "描述", "时间", "来源", "url"]]
df.to_excel("微信收藏链接5.xlsx", index=False)
