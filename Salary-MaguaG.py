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
import tkinter.filedialog
import tkinter.font as tf

window = tk.Tk()
window.title('发送工资条---MaguaG')
window.geometry('430x850')
window.configure(bg='#7F7F7F')
window.resizable(0,0)

r4 = tk.StringVar()
r5 = tk.StringVar()
r6 = tk.StringVar()
sa_sheet_add = ''
mail_sheet_add = ''
fail_sheet_add = ''
smtp_choose = ''
mail_add = ''
mail_pwd = ''


# 拆分功能
def split_excel(sa_sheet_add, fail_sheet_add):
    global nc
    global nr
    nc = 0
    nr = 0
    xlsx = xlrd.open_workbook(sa_sheet_add)
    table = xlsx.sheet_by_index(0)
    for p in range(table.col_values(0).__len__()):
        for q in range(table.row_values(0).__len__()):
            if table.cell_value(p, q) == '姓名' or table.cell_value(p,q) == '名字':
                nr = p
                nc = q
    for i in range(nr+1, table.col_values(0).__len__()):
        xlsx2 = xlwt.Workbook()
        sheetq = xlsx2.add_sheet('salary')
        # if table.cell_value(i, nc) == yname:
        for m in range(table.row_values(0).__len__()):
            for k in range(nr + 1):
                sheetq.write(k, m, table.cell_value(k, m))
            sheetq.write(nr+1, m, table.cell_value(i, m))
            xlsx2.save('{}{}.xls'.format(fail_sheet_add, table.cell_value(i, nc)))



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
        # print('{}已经发送'.format(yname))
        text.insert('end', '{}发送成功\r\n'.format(yname),'tagw')
        os.remove('{}{}.xls'.format(fail_sheet_add, yname))
    except smtplib.SMTPException:       
        text.insert('end', '{}发送失败,附件已保存，请检查邮箱表邮件地址\r\n'.format(yname),'tagr')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()    


def load1():
    global sa_sheet_add
    sa_sheet_add = tk.filedialog.askopenfilename()
    with open(sa_sheet_add,'r'):
        pass
    text.insert('end', '------已选择工资表路径--------\r\n','tagw')
    text.insert('end', '---------------------------------------------------\r\n','tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def load2():
    global mail_sheet_add
    mail_sheet_add = tk.filedialog.askopenfilename()
    with open(mail_sheet_add,'r'):
        pass
    text.insert('end', '------已选择邮件表路径--------\r\n','tagw')
    text.insert('end', '---------------------------------------------------\r\n','tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def load3():
    global fail_sheet_add
    fail_sheet_add = tk.filedialog.askdirectory() + '\\'
    text.insert('end', '------已经选择发送失败路径-----\r\n','tagw')
    text.insert('end', '---------------------------------------------------\r\n','tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def load4():
    global r7
    r7 = tk.filedialog.askopenfilename() +'\\'
    text.insert('end', '----------默认配置路径--------\r\n','tagw')
    text.insert('end', '---------------------------------------------------\r\n','tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def denglu():
    global smtp_choose
    global mail_add
    global mail_pwd
    smtp_choose = r4.get()
    mail_add = r5.get()
    mail_pwd = r6.get()

    if smtp_choose == '':
        text.insert('end', '**********请选择邮箱*********\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_add == '':
        text.insert('end', '********请输入邮箱用户名******\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_pwd == '':
        text.insert('end', '********请输入邮箱授权码*******\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    else:
        text.insert('end', '-----手动输入邮箱信息载入成功-----\r\n','tagw')
        text.insert('end', '---------------------------------------------------\r\n','tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()


def moren():
    global smtp_choose
    global mail_add
    global mail_pwd
    r7 = tk.filedialog.askopenfilename()
    with open(r7, 'r') as f:
        smtp_choose = f.readline().strip()
        mail_add = f.readline().strip()
        mail_pwd = f.readline().strip()
    text.insert('end', '-----已使用默认邮箱配置-------\r\n','tagw')
    text.insert('end', '---------------------------------------------------\r\n','tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def showins():
    text.insert('end', '--------------------使用说明----------------------\r\n','tagg')
    text.insert('end', '1.-所有表的必须都在Excel文件的第一个sheet内，请自行调整\r\n','tagg')
    text.insert('end', '2.-工资表表头必须包含“姓名”或者“员工姓名”\r\n','tagg')
    text.insert('end', '----如果是“名字”或者其它，请改成“姓名”或者“员工姓名”\r\n','tagg')
    text.insert('end', '3.-保存员工邮箱的表必须第一列为姓名，第二列为邮箱地址\r\n','tagg')
    text.insert('end', '4.-点击选择邮箱点击前面小圈\r\n','tagg')
    text.insert('end', '----邮箱密码为授权码而非通常密码，请自行百度申请\r\n','tagg')
    text.insert('end', '5.-登录邮箱默配置，请新建txt文本文件保存\r\n','tagg')
    text.insert('end', '*****第1行为邮箱服务\r\n','tagg')
    text.insert('end', '-----第2行为邮箱账号\r\n','tagg')
    text.insert('end', '*****第3行为申请的邮箱第三方登录授权码\r\n','tagg')
    text.insert('end', '-----关于怎么获取邮箱授权码请自行百度\r\n','tagg')
    text.insert('end', '-----关于邮箱服务请百度smtp\r\n','tagg')
    text.insert('end', '-----示例163：smtp.163.com\r\n','tagg')
    text.insert('end', '-----示例 QQ：smtp.qq.com\r\n','tagg')
    text.insert('end', '---------------------------------------------------\r\n','tagg')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def send_atart():
    if sa_sheet_add == '':
        text.insert('end', '**********请选工资表*********\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_sheet_add == '':
        text.insert('end', '**********请选择邮件表********\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif fail_sheet_add == '':
        text.insert('end', '******请选择发送失败保存路径***\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif smtp_choose == '':
        text.insert('end', '**********请勾选邮箱*********\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_add == '':
        text.insert('end', '********请输入邮箱用户名******\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_pwd == '':
        text.insert('end', '********请输入邮箱授权码******\r\n','tagy')
        text.insert('end', '---------------------------------------------------\r\n','tagy')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    else:
        text.insert('end', '--------发送中请稍等片刻------\r\n','tagw')
        text.insert('end', '---------------------------------------------------\r\n','tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
        rereadexcel = xlrd.open_workbook(sa_sheet_add)
        shet = rereadexcel.sheet_by_index(0)
        split_excel(sa_sheet_add, fail_sheet_add)
        for i in range(nr+1, shet.col_values(0).__len__()):
            a = str(find_mailadd(shet.cell_value(i, nc), mail_sheet_add))
            send_mail(shet.cell_value(i, nc), a, smtp_choose, mail_add, fail_sheet_add, mail_pwd)
            time.sleep(1)  
        text.insert("end", '---------------------------------------------------\r\n','tagw')
        text.insert("end", '-------------------------运行结束------------------\r\n','tagw')
        text.insert('end', '---------------------------------------------------\r\n','tagg')
        text.insert('end', '********************发送失败建议：*****************\r\n','tagg')
        text.insert('end', '---------------------------------------------------\r\n','tagg')
        text.insert('end', '方法1：将表手动发送给发送给失败的员工\r\n','tagg')
        text.insert('end', '-------失败的表都存放在失败路径里面\r\n','tagg')
        text.insert('end', '---------------------------------------------------\r\n','tagg')
        text.insert('end', '方法2：将发送失败的人员邮箱更正后\r\n','tagg')
        text.insert('end', '-------新建一个邮箱表录入发送失败员工邮箱信息\r\n','tagg')
        text.insert('end', '-------然后点击选择邮箱表选择新建的表\r\n','tagg')
        text.insert('end', '-------再次点击开始发送\r\n','tagg')
        text.insert('end', '-------会显示名单不在邮箱表内的人员发送失败，不用管\r\n','tagg')
        text.insert('end', '-------只用检差新表内人员是发送情况\r\n','tagg')
        text.insert('end', '---------------------------------------------------\r\n','tagg')
        text.insert('end', '*******完成之后手动清空失败表保存路径里面的表******\r\n','tagg')
        text.insert('end', '---------------------------------------------------\r\n','tagg')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()

# gui界面

f0 = tk.Label(window, text='Designed by MaguaG (-_0) E-mail:MaguaG9494@gmail.com', bg='black', font=('等线', 12), fg='white',width='200').pack()


i4 = tk.Label(window, text='-----------------------------------------------', bg='#7F7F7F', font=('等线', 12),fg='white', width=400).pack()
f1 = l4 = tk.Label(window, text='选择文件', bg='#7F7F7F', font=('等线', 20), fg='white',width='200').pack()

fram1 = tk.Frame()
f2 = tk.Label(fram1, text='请选择工资表：', bg='#BFBFBF', font=('等线', 15),fg='white', width=20).grid(row=1, column=1)
s1 = tk.Button(fram1, text='选择文件', bg='#3A414B',font=('等线', 11),fg='white',width=10, command=load1).grid(row=1, column=2)
fram1.pack()
i1 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 5), width=400).pack()

frame2 = tk.Frame()
l2 = tk.Label(frame2, text='请选择邮箱表：', bg='#BFBFBF', font=('等线', 15),fg='white', width=20).grid(row=2, column=1)
s2 = tk.Button(frame2, text='选择文件',bg='#3A414B',font=('等线', 11),fg='white',width=10, command=load2).grid(row=2, column=2)
frame2.pack()
i2 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 5), width=400).pack()

frame3 = tk.Frame()
l3 = tk.Label(frame3, text='发送失败保存路径：', bg='#BFBFBF', font=('等线', 15),fg='white', width=20).grid(row=3, column=1)
s3 = tk.Button(frame3, text='选择文件', bg='#3A414B',font=('等线', 11),fg='white',width=10, command=load3).grid(row=3, column=2)
frame3.pack()
i5 = tk.Label(window, text='-----------------------------------------------', bg='#7F7F7F', font=('等线', 12),fg='white', width=400).pack()
f3 = tk.Label(window, text='登录邮箱', bg='#7F7F7F', font=('等线', 22),fg='white', width=400).pack()

fram2 = tk.Frame(bg='#7F7F7F')
rb1 = tk.Radiobutton(fram2, text='QQ_邮箱', value='smtp.qq.com', variable=r4,bg='#7F7F7F',fg='#1C1C1C').grid(row=1, column=1)
rb2 = tk.Radiobutton(fram2, text='126邮箱', value='smtp.126.com', variable=r4,bg='#7F7F7F',fg='#1C1C1C').grid(row=1, column=2)
rb3 = tk.Radiobutton(fram2, text='163邮箱', value='smtp.163.com', variable=r4,bg='#7F7F7F',fg='#1C1C1C').grid(row=1, column=3)
rb4 = tk.Radiobutton(fram2, text='新浪邮箱', value='smtp.sina.com', variable=r4,bg='#7F7F7F',fg='#1C1C1C').grid(row=1, column=4)
fram2.pack()

fram4 = tk.Frame()
l8 = tk.Label(fram4, text='用户名：', font=('等线', 15), width=7,bg='#7F7F7F',fg='white').grid(row=1, column=1)
e8 = tk.Entry(fram4, textvariable=r5, width=28).grid(row=1, column=2)
submit4 = tk.Button(fram4, text='默认配置', bg='#3A414B',fg='white', font=('等线', 11),  command=moren).grid(row=1, column=3)
fram4.pack()

i9 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 8),fg='white', width=400).pack()
fram3 = tk.Frame()
l9 = tk.Label(fram3, text='授权码：', font=('等线', 15), width=7,bg='#7F7F7F',fg='white').grid(row=1, column=1)
e9 = tk.Entry(fram3, show='*', textvariable=r6, width=28).grid(row=1, column=2)
submit3 = tk.Button(fram3, text='输入登录', bg='#3A414B',fg='white', font=('等线', 11), command=denglu).grid(row=1, column=3)
fram3.pack()


i6 = tk.Label(window, text='-----------------------------------------------', bg='#7F7F7F', font=('等线', 12),fg='white', width=400).pack()
fram4 = tk.Frame()
submit2 = tk.Button(fram4, text='开始发送', bg='#3A414B',fg='white', font=('等线', 12), width=15, height=2,
                    command=send_atart).grid(row=1,column=3)
f6 = tk.Label(fram4,text='', font=('Arial', 30), width=2,bg='#7F7F7F').grid(row=1, column=2)
submit2 = tk.Button(fram4, text='查看说明', bg='#3A414B', fg='white',font=('等线', 12), width=15, height=2,
                    command=showins).grid(row=1,column=1)
fram4.pack()

f8 = tk.Label(window, text=' ', bg='#7F7F7F', font=('等线', 5), width=400).pack()

fram7 = tk.Frame()
text = tk.Text(fram7, width=300, height=100, borderwidth=5, bg='black')
text.pack()
text.insert('end', '---欢迎使用MaguaG工资条项目------\r\n','tagg')
text.insert('end', '---------------------------------------------------\r\n','tagg')
text.see(tk.END)
text.update()
fram7.pack()

# 设置text字体
ft = tf.Font(family='微软雅黑',size=10)
text.tag_add('tagg','end')
text.tag_config('tagg',foreground = 'green',font = ft)
text.tag_add('tagr','end')
text.tag_config('tagr',foreground = 'red',font = ft)
text.tag_add('tagy','end')
text.tag_config('tagy',foreground = '#FFFF00',font = ft)
text.tag_add('tagw','end')
text.tag_config('tagw',foreground = 'white',font = ft)



window.mainloop()
