import pick,os,time,sys

def GetDesktopPath():
    return os.path.join(os.path.expanduser("~"), 'Desktop')

if __name__ == "__main__":
    select,option=pick.pick(["1.是 ， 我 要 立 刻 卸 载 ComPE","2.否 ， 退 出 本 程 序"],"确 定 要 卸 载 ComPE吗 ？","→")
    if option == 0:
        try:
            os.system("bcdedit /delete {BEAAB3B9-80C3-13A8-6CF2-F5CC87F4006E} /cleanup")
            os.system("bcdedit /delete {40D2B668-D94D-7EE2-0C07-7A796BB7A1D1}")
            os.remove(".\\boot.sdi")
            os.remove(".\\boot.wim")
            os.remove(GetDesktopPath()+"\\卸载ComPE.lnk")
            print("卸载完成！感谢您使用ComPE。程序将在5秒后退出...")
            time.sleep(5)
            os.remove(".\\sys.argv[0]")
            os._exit(0)
        except:
            print("卸载失败！可能已经卸载完毕，如果已经卸载完毕请手动删除本程序，按任意键退出本程序...")
            os.system("pause>nul")
            os._exit(1)
    elif option == 1:
        os._exit(0)