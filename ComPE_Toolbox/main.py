import pick, zipfile, os, time, psutil, wmi, tqdm,sys,shutil


def select(title, menu):
    selected, option = pick.pick(menu, title, indicator="→")
    return option


def unzip():
    while True:
        dir = input(r"请输入指定目录（如：C:\Users\Username\)：")
        if not os.path.isdir(dir):
            print("目录不存在！请输入正确的目录!")
            continue
        else:
            print("开始复制镜像...")
            zip = zipfile.ZipFile(sys.argv[0])
            zip.extract(zip.namelist()[2],dir)
            zip.extract(zip.namelist()[3], dir)
            zip.close()
            print("解压完成！感谢您使用ComPE！程序将在5秒后退出...")
            time.sleep(5)
            os._exit(0)


def par_get_disk(partition_letter):  # 盘符取磁盘序号
    c = wmi.WMI()
    for disk_drive in c.Win32_DiskDrive():
        for partition in disk_drive.associators("Win32_DiskDriveToDiskPartition"):
            for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                if logical_disk.Caption == partition_letter:
                    return int(disk_drive.DeviceID[-1])


def getpars():  # 取可移动磁盘
    disklist = []
    for disk in psutil.disk_partitions():
        if par_get_disk(disk.device[0] + ":") == par_get_disk(systempar) or par_get_disk(disk.device[0] + ":") is None:
            pass
        else:
            disklist.append(disk.device + "            分 区 总 空 间 ：" + str(
                int(psutil.disk_usage(disk[0]).total / 1024 / 1024 / 1024)) + " GB")
    return disklist

def getsystempars():  # 取可移动磁盘
    disklist = []
    for disk in psutil.disk_partitions():
        if not par_get_disk(disk.device[0] + ":") == par_get_disk(systempar):
            pass
        else:
            disklist.append(disk.device + "            分 区 总 空 间 ：" + str(
                int(psutil.disk_usage(disk[0]).total / 1024 / 1024 / 1024)) + " GB")
    return disklist


def getparnum(par):  # 取分区序号
    parlist = []
    for the in getpars():
        if par_get_disk(the[0:2]) == par_get_disk(par):
            parlist.append(the[0:2])
    return parlist.index(par) + 1


def install_to_movable_disk():  # 安装到可移动磁盘
    print("正在获取可移动磁盘分区列表...")
    disks = getpars()
    disk = select("您 的 计 算 机 含 有 以 下 可 移 动 磁 盘 分 区 ，\n请 选 择 安 装 到 的 磁 盘 位 置 ：", disks)
    print("您选择了", disks[disk][0:2])
    print("警告！本操作会清楚您选择的磁盘数据，请再三确认备份好后，按下任意键继续...")
    os.system("pause>nul")
    print("准备开始写入...")
    for cmd in tqdm.tqdm(range(5)):
        if cmd == 0:
            with open("./movabledisk.txt", "w") as f:
                f.write("SELECT DISK " + str(par_get_disk(disks[disk][0:2])) +
                        "\nSELECT PARTITION " + str(getparnum(disks[disk][0:2])) +
                        "\nFORMAT FS=FAT32 QUICK" +
                        "\nACTIVE")
        elif cmd == 1:
            init_disk()
        elif cmd == 2:
            zip = zipfile.ZipFile(sys.argv[0])
            for name in zip.namelist():
                zip.extract(name, "./runtime")
            zip.close()
        elif cmd == 3:
            scode = os.system(".\\runtime\\7z.exe x -o"+disks[disk][0:3]+" "+".\\runtime\\ISO\\ComPE_Release.iso>nul")
            if not scode == 0:
                print("运行失败！请确认程序是否完整，以及分区是否存在。请按任意键退出程序...")
                os.remove("./movabledisk.txt")
                os.removedirs("./runtime")
                os.system("pause")
                os._exit(1)
        elif cmd == 4:
            os.remove("./movabledisk.txt")
            os.system("del /s /f /q .\\runtime>nul")
    print("运行完成！感谢您使用ComPE！程序将在5秒后退出...")
    time.sleep(5)
    os._exit(0)



with os.popen("cmd /c echo %windir%") as acs:
    systempar = acs.readlines()[0][0:2]


def init_disk():
    status_code = os.system("diskpart -s .\\movabledisk.txt>nul")
    if not status_code == 0:
        print("分区初始化失败！请确认分区是否仍然存在！请按任意键退出程序...")
        os.remove("./movabledisk.txt")
        os.system("del /s /f /q .\\runtime>nul")
        os.system("pause")
        os._exit(1)

def install_to_BCD():
    print("正在获取系统磁盘分区列表...")
    disks = getsystempars()
    disk = select("您 的 计 算 机 含 有 以 下 系 统 磁 盘 分 区 ，\n请 选 择 安 装 到 的 磁 盘 位 置 ：", disks)
    print("您选择了", disks[disk][0:2])
    print("由于程序会修改启动项，本程序可能会被安全软件误识别为高危程序，建议退出所有安全软件后，按下任意键继续...")
    os.system("pause")
    guid1="{BEAAB3B9-80C3-13A8-6CF2-F5CC87F4006E}"
    guid2="{40D2B668-D94D-7EE2-0C07-7A796BB7A1D1}"
    print("配置临时文件...")
    zip = zipfile.ZipFile(sys.argv[0])
    for name in zip.namelist():
        zip.extract(name, "./runtime")
    zip.close()
    print
    cmds = [".\\runtime\\7z.exe e .\\runtime\\ISO\\ComPE_Release.iso \"sources\\boot.wim\" -o"+disks[disk][0:3]+"sources",
            ".\\runtime\\7z.exe e .\\runtime\\ISO\\ComPE_Release.iso \"boot\\boot.sdi\" -o" + disks[disk][0:3] + "sources",
            "bcdedit /create "+guid1+" /d \"进入ComPE维护系统\" /application osloader",
            "bcdedit /create "+guid2+" /device",
            "bcdedit /set "+guid2+" ramdisksdidevice partition=\""+disks[disk][0:2]+"\"",
            "bcdedit /set "+guid2+" ramdisksdipath \\sources\\boot.sdi",
            "bcdedit /set "+guid1+" device ramdisk=\"["+disks[disk][0:2]+"]\\sources\\boot.wim,"+guid2,
            "bcdedit /set "+guid1+" osdevice ramdisk=\"["+disks[disk][0:2]+"]\\sources\\boot.wim,"+guid2,
            "bcdedit /set "+guid1+" path \\windows\\system32\\boot\\winload.exe",
            "bcdedit /set "+guid1+" systemroot \\windows",
            "bcdedit /set "+guid1+" detecthal yes",
            "bcdedit /set "+guid1+" winpe yes",
            "bcdedit /displayorder "+guid1+" /addlast"
            ]
    print("正在写入系统...")
    print("============================================================")
    for cmd in cmds:
        scode=os.system(cmd)
        if not scode == 0:
            os.system("del /s /f /q .\\runtime>nul")
            print("程序运行失败！请确认程序是否完整，以及是否被安全软件误处理。请按任意键退出本程序...")
            os._exit(1)
    shutil.copyfile(".\\runtime\\uninstall.exe",disks[disk][0:3]+"sources\\uninstall.exe")
    os.system("mshta VBScript:Execute(\"Set a=CreateObject(\"\"WScript.Shell\"\"):Set b=a.CreateShortcut(a.SpecialFolders(\"\"Desktop\"\") & \"\"\卸载ComPE.lnk\"\"):b.TargetPath=\"\""+disks[disk][0:3]+"sources\\uninstall.exe"+"\"\":b.WorkingDirectory=\"\"%~dp0\"\":b.Save:close\")")
    print("============================================================")
    print("写入完成！感谢您使用ComPE！程序将在5秒后自动退出...")
    os.system("del /s /f /q .\\runtime>nul")
    time.sleep(5)
    os.exit(0)


if __name__ == '__main__':
    while True:
        os.system("cls")
        mode = select(
            "欢 迎 使 用 ComPE工 具 箱\n该 工 具 箱 可 以 帮 助 您 将 ComPE安 装 到 指 定 位 置 \n请 使 用 方 向 键 ↑ ↓ 选 择 安 装 方 式 ：",
            ["1.复 制 ISO镜 像 文 件 到 指 定 位 置", "2.安 装 到 可 移 动 磁 盘", "3.安 装 ComPE到 系 统"])
        if mode == 0:
            print("您选择了复制ISO镜像文件到指定位置。")
            unzip()
        elif mode == 1:
            print("您选择了安装ComPE到可移动磁盘。")
            install_to_movable_disk()
        elif mode == 2:
            print("您选择了安装ComPE到系统。")
            install_to_BCD()