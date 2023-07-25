import ctypes
import datetime
import os
import shutil
import subprocess
import threading
import tkinter
import tkinter.ttk
import winreg
from _winapi import DETACHED_PROCESS
from time import sleep
from tkinter import filedialog, messagebox

import psutil


def write_file_config(new_path):
    try:
        app_data_path = os.path.expandvars('$APPDATA') + "\\Tencent\\WeChat\\All Users\\config\\3ebffe94.ini"
        f = open(app_data_path, 'w', encoding='utf-8')
        f.write(new_path)
        f.close()
        print("写入配置文件成功!!!")
        return True
    except:
        print("写入配置文件失败!!!")
        return False


# 写入注册表文件
def write_reg_config(reg_root, reg_path, reg_keyname, key_value):
    # 注意打开权限，默认只读
    try:
        key = winreg.OpenKeyEx(reg_root, reg_path, reserved=0, access=winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, reg_keyname, 0, winreg.REG_SZ, key_value)
        winreg.CloseKey(key)
        print("写入注册表配置成功!!!")
        return True
    except:
        print("写入注册表配置失败!!!")
        return False


# 查询硬盘分区容量
def query_disk_freespace(folder):
    free_bytes = ctypes.c_ulonglong(0)
    ctypes.windll.kernel32.GetDiskFreeSpaceExW(ctypes.c_wchar_p(folder), None, None, ctypes.pointer(free_bytes))
    return free_bytes.value / 1024 / 1024 / 1024


# 获取硬盘分区
def getDisklist():
    disklist = []
    for i in psutil.disk_partitions():
        disklist.append(i.device)
    return disklist


# 返回可用的分区目录；分区的容量必须大于微信存储容量20GB。
def get_aralible_disk():
    disklist = getDisklist()
    for i in disklist:
        if i == os.path.expandvars('$SystemDrive') + "\\":
            # print("系统盘，跳过")
            continue
        if query_disk_freespace(i) - wx_file_size >= 20:
            # print("空间可用")
            return i
    return "None"


# 查询指定目录的大小
def query_folder_size(path='.'):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            total_size += os.path.getsize(fp)
    return round(total_size / 1024 / 1024 / 1024, 2)  # 保留2为小数


def wechat_file_size():
    if read_wechat_file_config() == "MyDocument:":
        return query_folder_size(os.path.expandvars('$USERPROFILE') + "\\Documents\\WeChat Files\\")
    else:
        return query_folder_size(read_wechat_file_config())


# 获取微信的聊天记录目录
def wx_old_path():
    old_paht = read_wechat_file_config()
    if old_paht == "MyDocument:":
        return os.path.expandvars('$USERPROFILE') + "\\Documents\\"
    else:
        return old_paht


def read_wechat_file_config():
    app_data_path = os.path.expandvars('$APPDATA') + "\\Tencent\\WeChat\\All Users\\config\\3ebffe94.ini"
    f = open(app_data_path, 'r', encoding='utf-8')
    context = f.read()
    f.close()
    return context


# 迁移文件,Win7及以上，robocopy复制
def move_file(old_file_path, new_file_path):
    # if os.path.exists(new_file_path + "\\WeChat Files\\"):
    if os.path.exists(os.path.join(new_file_path, "WeChat Files")):
        return False
    else:
        # shutil.copytree(old_file_path + "\\WeChat Files\\", new_file_path + "\\WeChat Files\\")
        # shutil.rmtree(wx_old_path() + "\\WeChat Files\\")
        # shutil.copytree 模块复制大量小文件的情况，发生卡死，复制不全，效率低下。
        # shutil.copytree(os.path.join(wx_old_path(), "WeChat Files"), os.path.join(new_file_path, "WeChat Files"))
        # shutil.rmtree(os.path.join(wx_old_path(), "WeChat Files"))
        if os.path.exists(os.path.join(os.path.expandvars('$SystemRoot')) + "/system32/robocopy.exe"):
            # robocopy,移动文件
            commond1 = "robocopy.exe " + "\"" + os.path.join(wx_old_path(),
                                                             "WeChat Files") + "\"" + " " + "\"" + os.path.join(
                new_file_path,
                "WeChat Files") + "\"" + " /S  /MOVE"
            # os.system(commond1)  #会出现命令行窗口
            subprocess.call(commond1, creationflags=DETACHED_PROCESS)
            return True
        else:
            # xcopy 复制文件
            commond2 = "xcopy.exe " + "\"" + os.path.join(wx_old_path(),
                                                          "WeChat Files") + "\\\"" + " " + "\"" + os.path.join(
                new_file_path, "WeChat Files") + "\\\"" + " /C /R /Y /S"
            os.system(commond2)
            shutil.rmtree(os.path.join(wx_old_path(), "WeChat Files"))
            return True


# 存储位置及容量检查,检查后其余按钮可见
def check():
    btn1.pack_forget()
    ps_bar_start()
    # entry1.insert(0, read_wechat_file_config())  # 将获取到的存储文件位置，显示到文本框
    entry1.insert(0, wx_old_path())  # 将获取到的存储文件位置，显示到文本框
    entry2.insert(0, str(wx_file_size) + " GB")

    entry3.insert(0, new_path)
    btn2.pack(side=tkinter.LEFT, padx=2)
    btn3.pack(side=tkinter.LEFT, padx=2)
    btn4.pack(side=tkinter.LEFT, padx=2)
    ps_bar_stop()
    pass


# 迁移聊天记录
def sub_prog():
    ps_bar_start()
    th = threading.Thread(None, target=migrate, name="th", daemon=True)
    th.start()
    pass


def migrate():
    subprocess.call("taskkill /F /IM wechat.exe", creationflags=DETACHED_PROCESS)
    # os.system("taskkill /F /IM wechat.exe")
    btn2.pack_forget()
    sleep(2)
    if ck1.get():
        # 选中
        # 迁移命令
        if move_file(wx_old_path(), new_path):
            # 迁移成功
            write_file_config(new_path)
            messagebox.showinfo("提示", "迁移完毕，新路径：" + new_path)
            pass
        else:
            # 迁移失败
            messagebox.showinfo("提示", "迁移失败：您选择的位置已存在微信目录。")
            pass
        ps_bar_stop()
    else:
        # 未选中
        # 删除原有记录
        shutil.rmtree(wx_old_path() + "\\WeChat Files\\")
        write_file_config(new_path)
        ps_bar_stop()
        messagebox.showinfo("提示", "迁移完毕，历史记录已清空！新路径：" + new_path)
        pass


# 恢复默认
def default():
    subprocess.call("taskkill /F /IM wechat.exe", creationflags=DETACHED_PROCESS)
    write_file_config("MyDocument:")
    write_reg_config(winreg.HKEY_CURRENT_USER, r"SOFTWARE\\Tencent\\WeChat", "FileSavePath", "MyDocument:")
    messagebox.showinfo("提示", "已恢复默认位置：" + "MyDocument:")
    pass


# 新用户预设,
def preset():
    write_reg_config(winreg.HKEY_CURRENT_USER, r"SOFTWARE\\Tencent\\WeChat", "FileSavePath", new_path)
    messagebox.showinfo("提示", "预设完毕，新路径：" + new_path)
    pass


def entry3_click(event):
    entry3.delete(0, "end")
    f = ""
    f = filedialog.askdirectory()
    entry3.insert(0, f)
    global new_path
    new_path = f
    pass


def new_path():
    np = get_aralible_disk()
    if np != "None":
        return np + "WeChat" + str(datetime.datetime.now().date())
    else:
        return ""


def ps_bar_start():
    label1.pack(side=tkinter.LEFT)
    psbar.pack(side=tkinter.LEFT)
    psbar.start()
    pass


def ps_bar_stop():
    psbar.stop()
    label1.pack_forget()
    psbar.pack_forget()
    pass


root = tkinter.Tk()
root.title("微信历史记录迁移程序")
root.geometry("400x180")  # 指定窗口的大小

wx_file_size = wechat_file_size()
new_path = new_path()
ck1 = tkinter.IntVar()
ck1.set(1)

f1 = tkinter.Frame(root)
f1.pack(side=tkinter.TOP)
f2 = tkinter.Frame(root)
f2.pack(side=tkinter.TOP)
f3 = tkinter.Frame(root)
f3.pack(side=tkinter.LEFT)

tkinter.Label(f1, text="当前微信存储位置 ：").grid(row=0, column=0, pady=2)
entry1 = tkinter.Entry(f1, width=20)
entry1.grid(row=0, column=1, padx=2, pady=2)
tkinter.Label(f1, text="聊天记录空间使用：").grid(row=1, column=0, pady=2)
entry2 = tkinter.Entry(f1, width=20)
entry2.grid(row=1, column=1, padx=2, pady=2)
tkinter.Label(f1, text="选 择 新 位 置：").grid(row=2, column=0, pady=2)
entry3 = tkinter.Entry(f1, width=20)  # 路径选择器
entry3.grid(row=2, column=1, padx=2, pady=2)
entry3.bind("<Button-1>", entry3_click)
tkinter.Label(f1, text="是否保留聊天记录：").grid(row=3, column=0, pady=2)
checkbutton1 = tkinter.Checkbutton(f1, width=20, variable=ck1)  # 复选框需要设置variable=ck1，状态放入变量
checkbutton1.grid(row=3, column=1, padx=2, pady=2)

btn1 = tkinter.Button(f2, text=" 检 查 ", command=check)
btn1.pack(side=tkinter.LEFT, padx=2)
btn2 = tkinter.Button(f2, text=" 迁 移 ", command=sub_prog)
btn3 = tkinter.Button(f2, text=" 恢复默认 ", command=default)
btn4 = tkinter.Button(f2, text="新用户预设", command=preset)

label1 = tkinter.Label(f3, text="处理中请稍后:")
psbar = tkinter.ttk.Progressbar(f3, length=300, mode="indeterminate", orient=tkinter.HORIZONTAL)

root.mainloop()
