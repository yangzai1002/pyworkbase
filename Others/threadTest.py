import os
from time import sleep,ctime
import _thread as thread
def fun():
    print("开始运性")
    sleep(2)
    print("fun2 yun xing jieshu:",ctime())
def main():
    thread.start_new(fun,())
    thread.join()
    print("-----over---")
if __name__=="__main__":
    main()