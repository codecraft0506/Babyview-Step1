from pywinauto.application import Application
from pywinauto import Desktop
from dotenv import load_dotenv
import os
import time

load_dotenv()

# 使用 uia 啟動主程式
app = Application(backend="uia").start(r"C:\SYPOS\SYPOS.exe")
dlg = app.window(title="新儀科技資訊管理系統登錄作業")
dlg.wait('visible', timeout=10)

# 輸入帳密
edits = dlg.descendants(control_type="Edit")
edits[1].type_keys(os.getenv("USERNAME"), with_spaces=True)
edits[0].type_keys(os.getenv("PASSWORD"), with_spaces=True)
dlg.child_window(title="確  定(S)", control_type="Button").click()

time.sleep(5)

# 抓主畫面
main_dlg = app.window(title_re=".*寶沛靴資訊管理系統.*")
main_dlg.wait('visible', timeout=10)

main_dlg.menu_select("批發銷貨(D)->批發出貨管理(W)")
# 等新畫面開啟
time.sleep(2)

# 嘗試取得「出貨管理」視窗（根據視窗標題或正則表達式）
shipment_dlg = app.window(title_re=".*批發出貨資料管理作業.*")
shipment_dlg.wait('visible', timeout=10)

# 勾選「單據日期」CheckBox
checkbox = shipment_dlg.child_window(title="單據日期", control_type="CheckBox")
checkbox.wait('enabled', timeout=5)
checkbox.click_input()

# 勾選「全部資料」RadioButton
radio_button = shipment_dlg.child_window(title="全部資料", control_type="RadioButton")
radio_button.wait('enabled', timeout=5)
radio_button.select()

app_win32 = Application(backend="win32").connect(process=app.process)
shipment_dlg = app_win32.window(title_re=".*批發出貨資料管理作業.*")
shipment_dlg.wait('visible', timeout=10)

time.sleep(5)

# 列出所有 Edit 控件來手動觀察順序
dt_pickers = shipment_dlg.descendants(class_name="SysDateTimePick32")

if len(dt_pickers) >= 2:
    start_picker = dt_pickers[0]
    end_picker = dt_pickers[1]

    start_picker.set_focus()
    start_picker.type_keys("2025/04/01", with_spaces=True)

    end_picker.set_focus()
    end_picker.type_keys("2025/04/22", with_spaces=True)
else:
    print("❌ 找不到日期欄位（SysDateTimePick32 控件）")