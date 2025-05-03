from pywinauto.application import Application
from pywinauto import Desktop, mouse, keyboard
from datetime import datetime, timedelta
import time
import json

# 載入設定檔
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

USERNAME = config["username"]
PASSWORD = config["password"]
SYPOS_PATH = config["sypos_path"]
START_DATE_POS = tuple(config["start_date_pos"])
END_DATE_POS = tuple(config["end_date_pos"])
REPORT_TITLE = config["report_title"]
OUTPUT_DIR = config["output_dir"]

# 計算日期
date_mode = config["date_mode"]
advance = config["advance"]

if date_mode == 'relative':
    unit = config.get("period_unit", "days")
    value = config.get("period_value", 1)
    end = datetime.today() - timedelta(days=advance)

    if unit == "days":
        start = end - timedelta(days=value - 1)
    elif unit == "weeks":
        start = end - timedelta(weeks=value) + timedelta(days=1)
    elif unit == "months":
        from dateutil.relativedelta import relativedelta
        start = end - relativedelta(months=value) + timedelta(days=1)
    else:
        raise ValueError("Unsupported period_unit in config.json")

    START_DATE = start.strftime("%Y%m%d")
    END_DATE = end.strftime("%Y%m%d")

elif date_mode == 'fixed':
    START_DATE = config["start_date"]
    END_DATE = config["end_date"]
    
else:
    raise ValueError("Unsupported date mode")

def fill_date(x, y, date_str):
    year, month, day = date_str[0:4], date_str[4:6], date_str[6:8]

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
app = Application(backend="uia").start(SYPOS_PATH)
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
fill_date(*START_DATE_POS, START_DATE)
time.sleep(0.5)
fill_date(*END_DATE_POS, END_DATE)

time.sleep(5)

# 點擊「列印(P)」
shipment_dlg.child_window(title="列印(P)", control_type="Button").invoke()

# 選擇報表並輸出
file_path = f"{OUTPUT_DIR}\\{REPORT_TITLE.replace('　', '').replace(' ', '')}-{START_DATE}-{END_DATE}"
export(REPORT_TITLE, file_path)

# 等待轉存完成
time.sleep(5)
wait_for_export_done_and_exit()
