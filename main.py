from pywinauto.application import Application
from pywinauto import Desktop, mouse, keyboard
import time

USERNAME = '10201'
PASSWORD = '852456'

def fill_date(x, y, date_str):
    year, month, day = date_str.split("/")

    mouse.click(coords=(x, y))
    time.sleep(0.3)

    keyboard.send_keys(month)
    keyboard.send_keys("{RIGHT}")
    keyboard.send_keys(day)
    keyboard.send_keys("{RIGHT}")
    keyboard.send_keys(year)

def export(report_title, save_path):
    print_dlg = app.window(title_re=".*列印選擇視窗.*")
    print_dlg.wait('visible', timeout=10)

    radio = print_dlg.child_window(title=report_title, control_type="RadioButton")
    radio.wait('exists visible ready', timeout=5)
    radio.select()

    save_btn = print_dlg.child_window(title="轉存檔案(S)", control_type="Button")
    save_btn.wait('enabled', timeout=5)
    save_btn.invoke()

    # 抓另存新檔對話框
    save_dlg = Desktop(backend="uia").window(title_re=".*(另存新檔|儲存|Save As).*")
    save_dlg.wait('visible', timeout=10)

    # 清除預設檔名並輸入新的
    filename_edit = save_dlg.child_window(title="檔案名稱(N):", control_type="Edit")
    filename_edit.wait('enabled', timeout=5)
    filename_edit.click_input()
    time.sleep(0.2)
    keyboard.send_keys("{BACKSPACE 30}")
    time.sleep(0.2)
    keyboard.send_keys(save_path)

    # 點擊「存檔」
    save_button = save_dlg.child_window(title="存檔(S)", control_type="Button")
    save_button.wait('enabled', timeout=10)
    save_button.click_input()

    # 如果檔案存在，處理覆蓋對話框
    try:
        replace_dlg = Desktop(backend="uia").window(title_re=".*確認另存新檔.*")
        replace_dlg.wait('visible', timeout=3)
        yes_btn = replace_dlg.child_window(title_re="是.*", control_type="Button")
        yes_btn.wait('enabled', timeout=3)
        yes_btn.click_input()
    except:
        pass  # 若沒出現覆蓋視窗，直接略過
    
def wait_for_export_done_and_exit():
    try:
        # 等待出現「轉存完畢」提示視窗，最長等 600 秒
        done_dlg = Desktop(backend="uia").window(title_re="進銷存管理.*")
        done_dlg.wait('visible', timeout=600)
        
        # 列出所有靜態文字控件內容
        static_texts = done_dlg.descendants()
        for ctrl in static_texts:
            text = ctrl.window_text()
            print(text)
            if "OK" in text:
                print("✅ 偵測到提示訊息：OK")
                break
        else:
            print("⚠️ 找到視窗但沒有『OK』文字")

    except:
        print("⚠️ 未偵測到『轉存完畢』提示，超時略過")

    # 最終關閉主程式
    app.kill()

# 啟動主程式
app = Application(backend="uia").start(r"C:\SYPOS\SYPOS.exe")
dlg = app.window(title="新儀科技資訊管理系統登錄作業")
dlg.wait('visible', timeout=10)

# 登入
edits = dlg.descendants(control_type="Edit")
edits[1].type_keys(USERNAME, with_spaces=True)
edits[0].type_keys(PASSWORD, with_spaces=True)
dlg.child_window(title="確  定(S)", control_type="Button").click()
time.sleep(5)

# 抓主畫面
main_dlg = app.window(title_re=".*寶沛靴資訊管理系統.*")
main_dlg.wait('visible', timeout=10)
main_dlg.menu_select("批發銷貨(D)->批發出貨管理(W)")
time.sleep(2)

# 出貨管理視窗
shipment_dlg = app.window(title_re=".*批發出貨資料管理作業.*")
shipment_dlg.wait('visible', timeout=10)

# 勾選條件
shipment_dlg.child_window(title="單據日期", control_type="CheckBox").click_input()
shipment_dlg.child_window(title="全部資料", control_type="RadioButton").select()

# 輸入日期
time.sleep(1)
fill_date(165, 131, "2025/5/1")
time.sleep(0.5)
fill_date(282, 131, "2025/5/1")

time.sleep(5)

# 點擊「列印(P)」
shipment_dlg.child_window(title="列印(P)", control_type="Button").invoke()

# 選擇報表並輸出
export("出　貨　單", r"C:\Github\Babyview-Step1\出貨單")

# 等待轉存完成
time.sleep(5)
wait_for_export_done_and_exit()
