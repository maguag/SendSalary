import xlrd
import xlwt
import smtplib
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import time
import os
import tkinter as tk


# 拆分功能
def split_excel(yname, sa_sheet_add, fail_sheet_add):
    xlsx = xlrd.open_workbook(sa_sheet_add)
    table = xlsx.sheet_by_index(0)
    xlsx2 = xlwt.Workbook()
    sheetq = xlsx2.add_sheet('salary')
    for i in range(1, table.col_values(0).__len__()):
        if table.cell_value(i, 0) == yname:
            for m in range(table.row_values(0).__len__()):
                sheetq.write(0, m, table.cell_value(0, m))
                sheetq.write(1, m, table.cell_value(i, m))
                xlsx2.save('{}{}.xls'.format(fail_sheet_add, table.cell_value(i, 0)))


# 获取邮箱功能
def find_mailadd(yname, mail_sheet_add):
    mailadd = xlrd.open_workbook(mail_sheet_add)
    mail = mailadd.sheet_by_index(0)
    for i in range(0, mail.col_values(0).__len__()):
        if str(mail.cell_value(i, 0)) == yname:
            return mail.cell_value(i, 1)


def send_mail(yname, found_add, smtp_choose, mail_add, fail_sheet_add, mail_pwd):
    mail_title = '{}{}月工资条'.format(yname, time.strftime("%m", time.localtime()))
    htmlwrite = '''
    <p>尊敬的{}:</p>
    <p>    感谢您对公司的无私贡献，您本月的工资表已发送至附件，请下载附件自行查看</p>
    <p>时间：{}</p>
    '''.format(yname, time.strftime("%Y-%m-%d %H:%M", time.localtime()))

    msg = MIMEMultipart()
    msg["Subject"] = Header(mail_title, 'utf-8')
    msg["From"] = mail_add
    msg["To"] = Header('尊敬的{}'.format(yname), 'utf-8')

    msg.attach(MIMEText(htmlwrite, 'html', 'utf-8'))
    attachment = MIMEApplication(open('{}{}.xls'.format(fail_sheet_add, yname), 'rb').read())
    attachment.add_header('content-Disposition', 'attachment', filename=('{}.xls'.format(yname)))
    msg.attach(attachment)

    try:
        smtp2 = SMTP_SSL(smtp_choose)
        # smtp2.set_debuglevel(1)
        smtp2.ehlo(smtp_choose)
        smtp2.login(mail_add, str(mail_pwd))
        smtp2.sendmail(mail_add, found_add, msg.as_string())
        smtp2.quit()
        print('{}已经发送'.format(yname))
        os.remove('{}{}.xls'.format(fail_sheet_add, yname))
    except smtplib.SMTPException:
        print('{}发送失败'.format(yname), "邮件已存至发送失败文件夹，请手动发送")


# 图形界面
def gui():
    window = tk.Tk()
    window.title('发送工资条---MaguaG')
    window.geometry('400x400')

    r1 = tk.StringVar()
    r2 = tk.StringVar()
    r3 = tk.StringVar()
    r4 = tk.StringVar()
    r5 = tk.StringVar()
    r6 = tk.StringVar()
    r7 = tk.StringVar()
    global sa_sheet_add
    global mail_sheet_add
    global fail_sheet_add
    global smtp_choose
    global mail_add
    global mail_pwd

    def write_data():
        global sa_sheet_add
        global mail_sheet_add
        global fail_sheet_add
        global smtp_choose
        global mail_add
        global mail_pwd
        sa_sheet_add = r1.get()
        mail_sheet_add = r2.get()
        fail_sheet_add = r3.get()
        smtp_choose = r4.get()
        mail_add = r5.get()
        mail_pwd = r6.get()
        # a = [sa_sheet_add, mail_sheet_add, fail_sheet_add, smtp_choose, mail_add, mail_pwd]
        # print(a)
        # return sa_sheet_add, mail_sheet_add, fail_sheet_add, smpt_choose, mail_add, mail_pwd
        # tk.messagebox.showinfo('提示','导入成功，请继续点击关闭窗口开始发送数据')
        print('----------------------------')
        print('-------已使用输入配置---------')

    def window_quit():
        print('----------------------------')
        print('-------执行发送邮件操作-------')
        window.quit()

    def moren():

        global sa_sheet_add
        global mail_sheet_add
        global fail_sheet_add
        global smtp_choose
        global mail_add
        global mail_pwd
        with open(r7.get(), 'r') as f:
            sa_sheet_add = f.readline().strip()
            mail_sheet_add = f.readline().strip()
            fail_sheet_add = f.readline().strip()
            smtp_choose = f.readline().strip()
            mail_add = f.readline().strip()
            mail_pwd = f.readline().strip()
        print('----------------------------')
        print('------已使用默认配置----------')

    l0 = tk.Label(window, text='填入表的路径', bg='#87cefa', font=('Arial', 15), width=400).pack()

    fram1 = tk.Frame()
    l1 = tk.Label(fram1, text='工资表路径：', bg='pink', font=('Arial', 12), width=12)
    e1 = tk.Entry(fram1, textvariable=r1)
    l1.grid(row=1, column=1)
    e1.grid(row=1, column=2)

    l2 = tk.Label(fram1, text='邮箱表路径：', bg='pink', font=('Arial', 12), width=12)
    e2 = tk.Entry(fram1, textvariable=r2)
    l2.grid(row=2, column=1)
    e2.grid(row=2, column=2)

    l3 = tk.Label(fram1, text='失败表路径：', bg='pink', font=('Arial', 12), width=12)
    e3 = tk.Entry(fram1, textvariable=r3)
    l3.grid(row=3, column=1)
    e3.grid(row=3, column=2)
    fram1.pack()

    l4 = tk.Label(window, text='选择邮箱服务', bg='#87cefa', font=('Arial', 15), width=400).pack()

    fram2 = tk.Frame()
    rb1 = tk.Radiobutton(fram2, text='QQ_邮箱', value='smtp.qq.com', variable=r4).grid(row=1, column=1)
    rb2 = tk.Radiobutton(fram2, text='126邮箱', value='smtp.126.com', variable=r4).grid(row=1, column=2)
    rb3 = tk.Radiobutton(fram2, text='163邮箱', value='smtp.163.com', variable=r4).grid(row=2, column=1)
    rb4 = tk.Radiobutton(fram2, text='新浪邮箱', value='smtp.sina.com', variable=r4).grid(row=2, column=2)
    fram2.pack()

    l5 = tk.Label(window, text='登录邮箱', bg='#87cefa', font=('Arial', 15), width=400).pack()

    fram3 = tk.Frame()
    l8 = tk.Label(fram3, text='邮箱用户名：', bg='#c1cdc1', font=('Arial', 12), width=12)
    e8 = tk.Entry(fram3, textvariable=r5)
    l8.grid(row=3, column=1)
    e8.grid(row=3, column=2)
    l9 = tk.Label(fram3, text='邮箱密码：', bg='#c1cdc1', font=('Arial', 12), width=12)
    e9 = tk.Entry(fram3, show='*', textvariable=r6)
    l9.grid(row=4, column=1)
    e9.grid(row=4, column=2)
    fram3.pack()
    l6 = tk.Label(window, text='配置文件地址', bg='#87cefa', font=('Arial', 15), width=400).pack()
    fram5 = tk.Frame()
    l1 = tk.Label(fram5, text='配置文件路径：', bg='grey', font=('Arial', 12), width=12)
    e1 = tk.Entry(fram5, textvariable=r7)
    l1.grid(row=1, column=1)
    e1.grid(row=1, column=2)
    fram5.pack()

    l5 = tk.Label(window, text='执行操作', bg='#87cefa', font=('Arial', 15), width=400).pack()
    fram4 = tk.Frame()
    submit0 = tk.Button(fram4, text='使用默认', width=15, command=moren).grid(row=1, column=0)
    submit1 = tk.Button(fram4, text='上传数据', width=15, command=write_data).grid(row=1, column=1)
    submit2 = tk.Button(fram4, text='点击发送', width=15, command=window_quit).grid(row=1, column=2)
    fram4.pack()

    window.mainloop()
    print('----------------------------')
    return sa_sheet_add, mail_sheet_add, fail_sheet_add, smtp_choose, mail_add, mail_pwd


if __name__ == '__main__':
    a = gui()
    sa_sheet_add = a[0]
    mail_sheet_add = a[1]
    fail_sheet_add = a[2]
    smtp_choose = a[3]
    mail_add = a[4]
    mail_pwd = a[5]
    rereadexcel = xlrd.open_workbook(sa_sheet_add)
    shet = rereadexcel.sheet_by_index(0)
    for i in range(1, shet.col_values(0).__len__()):
        split_excel(shet.cell_value(i, 0), sa_sheet_add, fail_sheet_add)
        a = str(find_mailadd(shet.cell_value(i, 0), mail_sheet_add))
        send_mail(shet.cell_value(i, 0), a, smtp_choose, mail_add, fail_sheet_add, mail_pwd)
        time.sleep(2)
