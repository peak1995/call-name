import tkinter as tk
import xlrd
import time
import random
import datetime
class call():
    # 初始化
    def __init__(self):
        # 第1步，建立窗口window
        self.window = tk.Tk()
        # 第2步，给窗口的可视化起名字
        self.window.title('班级点名器程序')
        # 第3步，设定窗口的大小(长＊宽)
        self.window.geometry('500x400')
        self.text = tk.StringVar()  # 创建str类型
        self.count = tk.StringVar()
        self.data = self.read_data()
        # 获取星期几
        d = datetime.datetime.now()
        self.day = d.weekday() + 1
    def read_data(self):
        '''
        数据读取
        :return:
        '''
        workbook = xlrd.open_workbook('学生名单.xlsx')
        sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
        data = list(sheet1.col_values(0))  # 读取第一列内容
        data.pop(0)  # 把姓名 去掉
        return data

    def take(self):
        '''
        负责随机抽取同学提问
        :return:
        '''
        for s in range(50):
            '''
            后几秒慢点，制造紧张气氛
            '''
            desc = ''
            if s == 47:
                time.sleep(0.5)
            elif s == 48:
                time.sleep(0.6)
            elif s == 48:
                time.sleep(0.7)
            elif s == 49:
                time.sleep(0.9)
            else:
                time.sleep(0.1)

            classes = random.sample(self.data, 2)
            desc += "随机点中了:%s\n" % classes[0]
            desc += "还有点中的:%s\n" % classes[1]

            if s == 49:
                self.savedesc(desc)  # 写入日志
            self.text.set(desc)  # 设置内容
            self.window.update()  # 屏幕更新

    def kill(self):
        if self.day == 1:
            count = random.randint(50, 100)
            kill_desc = "奖励了你们%d遍" % (count)

        elif self.day == 2:
            count = random.randint(50, 120)
            kill_desc = "奖励了你们%d遍" % (count)
            self.count.set(kill_desc)
        elif self.day == 3:
            count = random.randint(50, 140)
            kill_desc = "奖励了你们%d遍" % (count)
        elif self.day == 4:
            count = random.randint(50, 160)
            kill_desc = "奖励了你们%d遍" % (count)
            self.count.set(kill_desc)
        elif self.day == 5:
            count = random.randint(50, 180)
            kill_desc = "奖励了你们%d遍" % (count)
        else:
            count = random.randint(50, 180)
            kill_desc = "奖励了你们%d遍" % (count)
            # kill_desc = '周末就别提问了'

        self.count.set(kill_desc)  # 设置内容
        self.window.update()  # 屏幕更新

        self.savecount(kill_desc)  # 写入日志


    def gettime(self):
        if self.day==1:
            day="一"
        elif self.day==2:
            day="二"
        elif self.day==3:
            day='三'
        elif self.day==4:
            day='四'
        elif self.day==5:
            day='五'
        elif self.day==6:
            day='六'
        else:
            day='日'
        return time.strftime('%Y-%m-%d', time.localtime(time.time())) + "  星期" + day
        # str(self.day)


    def savedesc(self, desc):
        '''
        负责把选中的人写入到log里面
        :param desc:
        :return:
        '''
        with open('log.txt', 'a', encoding='utf-8') as f:
            f.write(self.gettime() + "\n" + desc)

    def savecount(self, count):
        '''
        负责把被罚写的遍数写入到log里面
        :param count:
        :return:
        '''
        with open('log.txt', 'a', encoding='utf-8') as f:
            f.write(str(count) + '\n')
            f.write('--------------------------------\n')

    def main(self):
        '''
        主函数负责绘制
        :return:
        '''
        # 绘制日期、班级总人数等
        # now = time.strftime('%Y-%m-%d', time.localtime(time.time())) + "  星期" + str(self.day)
        now = self.gettime()
        now += "\n班级总人数:%s人" % str(len(self.data))
        now += "\n正在随机抽取中\n"

        l1 = tk.Label(self.window, fg='blue', text=now, width=200, height=5)
        l1.config(font='Helvetica -%d bold' % 15)
        l1.pack()  # 安置标签

        # 绘制筛选信息
        l2 = tk.Label(self.window, fg='red', textvariable=self.text, width=400, height=3)
        l2.config(font='Helvetica -%d bold' % 30)
        l2.pack()

        # 绘制惩罚信息
        l3 = tk.Label(self.window, fg='red', textvariable=self.count, width=400, height=3)
        l3.config(font='Helvetica -%d bold' % 20)
        l3.pack()

        # 绘制筛选按钮
        btntake = tk.Button(self.window, text="筛选", width=15, height=2, command=self.take)
        btntake.pack()

        # 绘制惩罚按钮
        btnkill = tk.Button(self.window, text="惩罚", width=15, height=2, command=self.kill)
        btnkill.pack()

        # 进入循环
        self.window.mainloop()


if __name__ == '__main__':
    call = call()
    call.main()