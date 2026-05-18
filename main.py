import sys
import time
import subprocess
from pywinauto.application import Application


print("Starting the application...")
subprocess.Popen("C:\Program Files (x86)\Clipstamp\Clipstamp.exe")
time.sleep(1)
app = Application(backend="uia").connect(title="クリップスタンプ")
# 開いているウィンドウを全部表示
for w in app.windows():
    print(repr(w.window_text()), w.is_visible(), w.is_enabled())
dlg = app.window(title="クリップスタンプ")
dlg.wait('ready', timeout=10)

print("Performing calculations...")
if dlg.get_show_state() == 2:  # Check if the calculator is minimized
    dlg.restore()  # Restore the window if it is minimized

dlg.set_focus()

dlg.child_window(auto_id="134", control_type="Edit").set_text("久保\n")

dlg.child_window(title="印影保存", control_type="MenuItem").invoke()
time.sleep(1)

save_dlg = dlg.child_window(title="印影保存", control_type="Window")
save_dlg.wait("ready", timeout=10)
if save_dlg.get_show_state() == 2:  # Check if the save dialog is minimized
    save_dlg.restore()  # Restore the window if it is minimized
save_dlg.set_focus()
save_dlg.child_window(auto_id="1148", control_type="Edit").set_text(r"C:\Users\test2\ws\RPA\test.bmp")

save_dlg.child_window(auto_id="1",control_type="Button").click()  # Click the "Save" button
print("Saving the stamp...")

