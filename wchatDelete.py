# -*- coding: utf-8 -*-
# @author: vincentlz
# @file: wchatDelete.py
# @time: 2023/1/29 13:39

"""
文件功能&使用说明：

"""
import uiautomation as auto


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
collect = wechatWindow.ListControl(Name='收藏')

for _ in range(5):
    item = collect.GetFirstChildControl()
    item.RightClick()
    get_meau_item("删除").Click()
    confirmDialog = wechatWindow.WindowControl(
        searchDepth=1, Name="微信", ClassName='ConfirmDialog')
    delete_button = confirmDialog.ButtonControl(Name="确定")
    delete_button.Click()
