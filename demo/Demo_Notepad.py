# -*- coding: utf-8 -*-
# @Name:      Demo_Notepad.py

# 执行此示例前，请先手动启动记事本，如操作异常，并不是库导致，而是Windows系统版本不一致导致的控件属性不同，可通过运行Demo_Record.py录制属于自己的测试脚本
# Before executing this example, please manually start Notepad. If the operation fails and it is not due to the library,
# it is because the Windows system version is inconsistent, resulting in different control properties. You can run Demo_Record.py to record your own test script.

from easy_uiauto import run_action
from easy_uiauto import push_message  # , compile_controls,find_control

# compile_controls(max_depth=1)

TEST_INFO_1 = {'TEST_ID': '1', 'ACTION': '点击', 'LOCATION':
    {'WindowName': '无标题 - 记事本', 'Name': '文本编辑器', 'ClassName': 'Edit', 'ControlType': 'EditControl', 'foundIndex': '', 'AutomationId': '15',
     'Xpath': [
         {'ControlType': 'WindowControl', 'Name': '无标题 - 记事本', 'ClassName': 'Notepad', 'searchDepth': 1},
         {'ControlType': 'EditControl', 'Name': '文本编辑器', 'ClassName': 'Edit', 'AutomationId': '15', 'searchDepth': 2}],
     'Img': '', 'PARAMETERS': {'x': -1, 'y': -1}}}
push_message(run_action(TEST_INFO_1))
# print(find_control(TEST_INFO_1["LOCATION"],debug=True))

TEST_INFO_2 = {'TEST_ID': '2', 'ACTION': '输入文本',
               'LOCATION': {'WindowName': '', 'Name': '', 'ClassName': '', 'ControlType': '', 'foundIndex': '', 'AutomationId': '', 'Xpath': [], 'Img': '',
                            'PARAMETERS': {'输入文本': 'test'}}}
push_message(run_action(TEST_INFO_2))
# print(find_control(TEST_INFO_2["LOCATION"],debug=True))
