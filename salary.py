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
window.title('MaguaG(-_0)')
window.geometry('430x850')
window.configure(bg='#7F7F7F')
window.resizable(0, 0)
sa_sheet_add = ''
mail_sheet_add = ''
global fail_sheet_add
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
    style_1 = xlwt.XFStyle()

    ft = xlwt.Font()
    ft.name = '等线'
    ft.bold = True
    ft.height = 200
    style_1.font = ft

    bo = xlwt.Borders()
    bo.top = xlwt.Borders.THIN
    bo.bottom = xlwt.Borders.THIN
    bo.left = xlwt.Borders.THIN
    bo.right = xlwt.Borders.THIN
    style_1.borders = bo

    ali = xlwt.Alignment()
    ali.horz = xlwt.Alignment.HORZ_CENTER
    ali.vert = xlwt.Alignment.VERT_CENTER
    style_1.alignment = ali

    xlsx = xlrd.open_workbook(sa_sheet_add)
    table = xlsx.sheet_by_index(0)
    for p in range(table.col_values(0).__len__()):
        for q in range(table.row_values(0).__len__()):
            if table.cell_value(p, q) == '姓名':
                nr = p
                nc = q

    for i in range(nr + 1, table.col_values(0).__len__()):
        xlsx2 = xlwt.Workbook()
        sheetq = xlsx2.add_sheet(r'salary')
        for k in range(nr+1):
            for m in range(table.row_values(0).__len__()):
                sheetq.write(k, m, table.cell_value(k, m), style_1)
        for m in range(table.row_values(0).__len__()):
            sheetq.write(nr+1, m, table.cell_value(i, m), style_1)
        try:
            xlsx2.save(r'{}{}.xls'.format(fail_sheet_add, table.cell_value(i, nc)))
        except Exception:
            text.insert('end', '********请关闭所有Excel表*******\r\n', 'tagy')
            text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text.see(tk.END)  # 一直显示最新的一行
            text.update()


# 获取邮箱功能
def find_mailadd(yname, mail_sheet_add):
    mailadd = xlrd.open_workbook(mail_sheet_add)
    mail = mailadd.sheet_by_index(0)
    for i in range(0, mail.col_values(0).__len__()):
        if str(mail.cell_value(i, 0)) == yname:
            return mail.cell_value(i, 1)


def send_mail(yname, found_add, smtp_choose, mail_add, fail_sheet_add, mail_pwd):
    global checkadd

    mail_title = '{}{}月工资条'.format(yname, time.strftime("%m", time.localtime()))
    htmlwrite = '''
    <p>尊敬的{}：</p>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;感谢您对公司的无私贡献，您本月的工资表已发送至附件，请下载附件自行查看。</p>    
    <p>时间：{}</p>
    '''.format(yname,time.strftime("%Y-%m-%d %H:%M", time.localtime()))
    msg = MIMEMultipart()
    msg["Subject"] = Header(mail_title, 'utf-8')
    msg["From"] = mail_add
    msg["To"] = Header('尊敬的{}'.format(yname), 'utf-8')
    msg.attach(MIMEText(htmlwrite, 'html', 'utf-8'))
    attachment = MIMEApplication(open(r'{}{}.xls'.format(fail_sheet_add, str(yname)), 'rb').read())
    attachment.add_header('content-Disposition', 'attachment', filename=(r'{}.xls'.format(str(yname))))
    msg.attach(attachment)
    try:
        smtp2 = SMTP_SSL(smtp_choose)
        smtp2.ehlo(smtp_choose)
        smtp2.login(mail_add, str(mail_pwd))
        smtp2.sendmail(mail_add, found_add, msg.as_string())
        smtp2.quit()
        text.insert('end', '{}发送成功\r\n'.format(yname), 'tagw')
        os.remove('{}{}.xls'.format(fail_sheet_add, yname))
    except smtplib.SMTPException:
        text.insert('end', '{}发送失败,附件已保存，请检查邮箱表邮件地址\r\n'.format(yname), 'tagr')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()


def load1():
    global sa_sheet_add
    sa_sheet_add = tk.filedialog.askopenfilename()
    with open(sa_sheet_add, 'r'):
        pass
    text.insert('end', '------已选择工资表路径--------\r\n', 'tagw')
    text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()


def load2():
    global mail_sheet_add
    mail_sheet_add = tk.filedialog.askopenfilename()
    with open(mail_sheet_add, 'r'):
        pass
    text.insert('end', '------已选择邮件表路径--------\r\n', 'tagw')
    text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()



def load3():
    global  fail_sheet_add
    fail_sheet_add = tk.filedialog.askdirectory() + '\\'
    if fail_sheet_add != '\\':
        text.insert('end', '------已经选择发送失败路径-----\r\n', 'tagw')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
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
    text.insert('end', '-------读取配置，正在检查连通性-------\r\n', 'tagw')
    text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
    text.see(tk.END)  # 一直显示最新的一行
    text.update()
    try:
        smtp4 = SMTP_SSL(smtp_choose)
        smtp4.ehlo(smtp_choose)
        a=smtp4.login(mail_add, str(mail_pwd))
        smtp4.quit()
        if a[1] == b'Authentication successful':
            text.insert('end', '--------服务器连接成功，邮箱信息配置成功-----\r\n', 'tagw')
            text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text.see(tk.END)  # 一直显示最新的一行
            text.update()
    except Exception:
        text.insert('end', '------------服务器连接失败,请重新配置邮箱----------\r\n', 'tagy')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')




def showins():
    text.insert('end', '---------------------------使用说明-----------------------------\r\n', 'tagw')
    text.insert('end', '软件功能：按照姓名拆分表格，批量发送邮件\r\n', 'tagg')
    text.insert('end', '-------------适用于企业发送工资条和学校发送成绩单等场景\r\n', 'tagg')
    text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
    text.insert('end', '首次登录需要配置默认邮箱信息，请点击配置邮箱\r\n', 'tagg')
    text.insert('end', '配置完后，下次使用直接点击默认配置，选择配置好的txt文件\r\n', 'tagg')
    text.insert('end', '****服务器为：SMTP服务器地址\r\n', 'tagg')
    text.insert('end', '-----示例QQ邮箱为：smtp.qq.com\r\n', 'tagg')
    text.insert('end', '-----其他邮箱请百度smtp服务器地址大全\r\n', 'tagg')
    text.insert('end', '****用户名为邮箱账号\r\n', 'tagg')
    text.insert('end', '****授权码为申请的邮箱第三方登录授权码，而非邮箱密码\r\n', 'tagg')
    text.insert('end', '-----关于怎么获取邮箱授权码请自行百度\r\n', 'tagg')
    text.insert('end', '---------------------------注意事项-----------------------------\r\n', 'tagw')
    text.insert('end', '1.-所有表的必须都在Excel文件的第一个sheet，请自行调整\r\n', 'tagg')
    text.insert('end', '2.-表头不能带有合并单元格，必须包含“姓名”一栏\r\n', 'tagg')
    text.insert('end', '-----如果是“名字”或者其它表示，请改成“姓名”\r\n', 'tagg')
    text.insert('end', '-----如果有重名请在名字后加个数字区别开来\r\n', 'tagg')
    text.insert('end', '-----注意与拆分表姓名与邮箱表对应\r\n', 'tagg')
    text.insert('end', '3.-邮箱的表必须第一列为姓名，第二列为邮箱地址\r\n', 'tagg')
    text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
    text.see(tk.END)
    text.update()

def peizhi():
    s200=tk.Toplevel()
    s200.title('配置邮箱信息')
    s200.geometry('450x400')
    s200.configure(bg='#7F7F7F')
    s200.resizable(0, 0)
    s200.focus_force()
    r1 = tk.StringVar()
    r2 = tk.StringVar()
    r3 = tk.StringVar()
    global a
    a = False

    def loadd():
        if a==True:
            n8 = tk.filedialog.askdirectory()
            s9 = open('{}\\默认配置.txt'.format(n8), 'w', encoding='utf-8')
            s9.write('{}\n{}\n{}\n'.format(smtp_choose, mail_add, mail_pwd))
            s9.close()
            text.insert('end', '********保存配置成功***文件名:\"默认配置.txt\"****\r\n', 'tagw')
            text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text.see(tk.END)  # 一直显示最新的一行
            text.update()
            s200.destroy()
        else:
            text2.insert('end', '**********请检查连通性*********\r\n', 'tagy')
            text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text2.see(tk.END)  # 一直显示最新的一行
            text2.update()


    def check():
        global smtp_choose
        global mail_add
        global mail_pwd
        global a
        smtp_choose = r1.get()
        mail_add = r2.get()
        mail_pwd = r3.get()
        if smtp_choose == '':
            text2.insert('end', '********请输入smtp服务器*********\r\n', 'tagy')
            text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text2.see(tk.END)  # 一直显示最新的一行
            text2.update()
        elif mail_add == '':
            text2.insert('end', '********请输入邮箱用户名******\r\n', 'tagy')
            text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text2.see(tk.END)  # 一直显示最新的一行
            text2.update()
        elif mail_pwd == '':
            text2.insert('end', '********请输入邮箱授权码*******\r\n', 'tagy')
            text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text2.see(tk.END)  # 一直显示最新的一行
            text2.update()
        else:
            try:
                smtp3 = SMTP_SSL(smtp_choose)
                smtp3.ehlo(smtp_choose)
                b=smtp3.login(mail_add, str(mail_pwd))
                smtp3.quit()
                if b[1] == b'Authentication successful':
                    text2.insert('end', '--------连接成功，请点击保存配置-----\r\n', 'tagw')
                    text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
                    text2.see(tk.END)  # 一直显示最新的一行
                    text2.update()
                    a = True
            except Exception:
                text2.insert('end', '--------连接失败，请检查服务器，用户名，授权码是否输入有误-----\r\n', 'tagy')
                text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')

    o1 = tk.Label(s200, text='-----------------------------------------------------', bg='#7F7F7F', font=('等线', 10),fg='white',width=400).pack()
    o2 = tk.Label(s200, text='配置默认邮箱服务', bg='#7F7F7F', font=('等线', 15), fg='white', width=400).pack()
    i10 = tk.Label(s200, text='', bg='#7F7F7F', font=('等线', 3), fg='white', width=400).pack()
    fram00 = tk.Frame(s200,bg='#7F7F7F')
    l8 = tk.Label(fram00, text='服务器:', font=('等线', 12), width=9, bg='#7F7F7F', fg='white').grid(row=1, column=1)
    e8 = tk.Entry(fram00, textvariable=r1, width=30).grid(row=1, column=2)
    fram00.pack()
    i11 = tk.Label(s200, text='', bg='#7F7F7F', font=('等线', 3), fg='white', width=400).pack()
    fram4 = tk.Frame(s200,bg='#7F7F7F')
    l8 = tk.Label(fram4, text='用户名:', font=('等线', 12), width=9, bg='#7F7F7F', fg='white').grid(row=1, column=1)
    e8 = tk.Entry(fram4, textvariable=r2, width=30).grid(row=1, column=2)
    fram4.pack()

    i9 = tk.Label(s200, text='', bg='#7F7F7F', font=('等线', 3), fg='white', width=400).pack()

    fram3 = tk.Frame(s200,bg='#7F7F7F')
    l9 = tk.Label(fram3, text='授权码:', font=('等线', 12), width=9, bg='#7F7F7F', fg='white').grid(row=1, column=1)
    e9 = tk.Entry(fram3, show='*', textvariable=r3, width=30).grid(row=1, column=2)
    fram3.pack()
    f88 = tk.Label(s200, text=' ', bg='#7F7F7F', font=('等线', 3), width=400).pack()
    framm8 = tk.Frame(s200,bg='#7F7F7F')
    submit3 = tk.Button(framm8, text='检查连通性',bg='#3A414B', fg='white', font=('等线', 12), command=check,width=18).grid(row=1,column=1)
    lll =tk.Label(framm8,bg='#7F7F7F',width=2).grid(row=1,column=2)
    submit4=tk.Button(framm8, text='保存配置',bg='#3A414B', fg='white', font=('等线', 12), command=loadd,width=18).grid(row=1,column=3)
    framm8.pack()
    f88 = tk.Label(s200, text=' ', bg='#7F7F7F', font=('等线', 3), width=400).pack()

    framm7 = tk.Frame(s200, bg='#7F7F7F')
    text2 = tk.Text(framm7, width=300, height=30, borderwidth=0, bg='black')
    text2.pack()
    text2.insert('end', '--------开始配置邮箱------\r\n', 'tagg')
    text2.insert('end', '-----------------------------------------------------------------\r\n', 'tagg')
    text2.see(tk.END)
    text2.update()
    framm7.pack()
    ft2 = tf.Font(family='微软雅黑', size=10)
    text2.tag_add('tagg', 'end')
    text2.tag_config('tagg', foreground='green', font=ft)
    text2.tag_add('tagr', 'end')
    text2.tag_config('tagr', foreground='red', font=ft)
    text2.tag_add('tagy', 'end')
    text2.tag_config('tagy', foreground='#FFFF00', font=ft)
    text2.tag_add('tagw', 'end')
    text2.tag_config('tagw', foreground='white', font=ft)




def send_atart():
    if   mail_add == ''or smtp_choose == ''or mail_pwd == '':
        text.insert('end', '*********请配置默认邮箱**********\r\n', 'tagy')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif sa_sheet_add == '':
        text.insert('end', '**********请选择待拆分的表*********\r\n', 'tagy')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif mail_sheet_add == '':
        text.insert('end', '**********请选择邮件表********\r\n', 'tagy')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    elif fail_sheet_add == ''or fail_sheet_add=='\\':
        text.insert('end', '******请选择存放发送失败文件夹***\r\n', 'tagy')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
    else:
        text.insert('end', '--------发送中请稍等片刻------\r\n', 'tagw')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()
        rereadexcel = xlrd.open_workbook(sa_sheet_add)
        shet = rereadexcel.sheet_by_index(0)
        mm = False
        for p in range(shet.col_values(0).__len__()):
            for q in range(shet.row_values(0).__len__()):
                if shet.cell_value(p, q) == '姓名':
                    mm=True
        if mm == True:
            split_excel(sa_sheet_add, fail_sheet_add)
        else:
            text.insert('end', '******请认真查看说明检查表头是否含有\'姓名\'********\r\n', 'tagy')
            text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
            text.see(tk.END)  # 一直显示最新的一行
            text.update()
        for i in range(nr + 1, shet.col_values(0).__len__()):
            a = str(find_mailadd(shet.cell_value(i, nc), mail_sheet_add))
            send_mail(shet.cell_value(i, nc), a, smtp_choose, mail_add, fail_sheet_add, mail_pwd)
            time.sleep(1)
        text.insert("end", '---------------------------运行结束-----------------------------\r\n', 'tagw')
        text.insert('end', '*********发送失败建议***************\r\n', 'tagg')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.insert('end', '方法1：适用于失败人数较少情况，更正邮箱\r\n', 'tagg')
        text.insert('end', '-------将表手动发送给发送给失败的人员\r\n', 'tagg')
        text.insert('end', '-------失败的表都存放在失败路径里面\r\n', 'tagg')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.insert('end', '方法2：发送失败多为邮箱表记录错误，将发送失败的人员邮箱更正后\r\n', 'tagg')
        text.insert('end', '-------新建一个工资表，将失败人员信息整合到表内\r\n', 'tagg')
        text.insert('end', '-------再次操作软件，选择待拆分表的时候选择刚刚新建的表\r\n', 'tagg')
        text.insert('end', '-------点击开始发送即可\r\n', 'tagg')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.insert('end', '*******完成之后手动清空失败表保存路径里面的表******\r\n', 'tagg')
        text.insert('end', '*******大批发送失败的情况，请重置授权码，使用新的授权码登录**\r\n', 'tagg')
        text.insert('end', '-----------------------------------------------------------------\r\n', 'tagw')
        text.see(tk.END)  # 一直显示最新的一行
        text.update()


def qingping():
    text.delete('1.0', 'end')


# gui界面

f0 = tk.Label(window, text='Designed by MaguaG (-_0) E-mail:MaguaG9494@gmail.com', bg='black', font=('等线', 12),
              fg='white', width='200').pack()

i5 = tk.Label(window, text='-----------------------------------------------------', bg='#7F7F7F', font=('等线', 10), fg='white',
              width=400).pack()
f3 = tk.Label(window, text='配置邮箱', bg='#7F7F7F', font=('等线', 20), fg='white', width=400).pack()

framm1=tk.Frame(window,bg='#7F7F7F')
lab1 = tk.Label(framm1, text='首次使用请配置默认邮箱', bg='#BFBFBF', font=('等线', 15), fg='white', width=20).grid(row=1,column=2)
butt1 = tk.Button(framm1, text='配置邮箱', bg='#3A414B', font=('等线', 11), fg='white', width=10, command=peizhi).grid(row=1,column=1)
framm1.pack()
i6 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 5), width=400).pack()

framm2=tk.Frame(window,bg='#7F7F7F')
lab1 = tk.Label(framm2, text='选择已有的邮箱配置', bg='#BFBFBF', font=('等线', 15), fg='white', width=20).grid(row=1,column=1)
butt1 = tk.Button(framm2, text='默认配置', bg='#3A414B', font=('等线', 11), fg='white', width=10, command=moren).grid(row=1,column=2)
framm2.pack()



i4 = tk.Label(window, text='-----------------------------------------------------', bg='#7F7F7F', font=('等线', 10), fg='white',
              width=400).pack()
f1 = l4 = tk.Label(window, text='选择文件', bg='#7F7F7F', font=('等线', 20), fg='white', width='200').pack()

fram1 = tk.Frame(window,bg='#7F7F7F')
f2 = tk.Label(fram1, text='请选择待拆分的表', bg='#BFBFBF', font=('等线', 15), fg='white', width=20).grid(row=1, column=1)
s1 = tk.Button(fram1, text='选择文件', bg='#3A414B', font=('等线', 11), fg='white', width=10, command=load1).grid(row=1,
                                                                                                            column=2)
fram1.pack()
i1 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 5), width=400).pack()

frame8 = tk.Frame(window,bg='#7F7F7F')
l2 = tk.Label(frame8, text='选择邮箱表', bg='#BFBFBF', font=('等线', 15), fg='white', width=20).grid(row=2, column=1)
s2 = tk.Button(frame8, text='选择文件', bg='#3A414B', font=('等线', 11), fg='white', width=10, command=load2).grid(row=2,
                                                                                                             column=2)
frame8.pack()
i2 = tk.Label(window, text='', bg='#7F7F7F', font=('等线', 5), width=400).pack()

frame3 = tk.Frame(window)
l3 = tk.Label(frame3, text='存放发送失败表文件夹', bg='#BFBFBF', font=('等线', 15), fg='white', width=20).grid(row=3, column=1)
s3 = tk.Button(frame3, text='选择文件', bg='#3A414B', font=('等线', 11), fg='white', width=10, command=load3).grid(row=3,
                                                                                                             column=2)
frame3.pack()


i6 = tk.Label(window, text='-----------------------------------------------------', bg='#7F7F7F', font=('等线', 10), fg='white',
              width=400).pack()

fram4 = tk.Frame(window,bg='#7F7F7F')
submit2 = tk.Button(fram4, text='开始发送', bg='#3A414B', fg='white', font=('等线', 12), width=15, height=2,
                    command=send_atart).grid(row=1, column=3)
f6 = tk.Label(fram4, text='', font=('Arial', 30), width=2, bg='#7F7F7F').grid(row=1, column=2)
submit2 = tk.Button(fram4, text='查看说明', bg='#3A414B', fg='white', font=('等线', 12), width=15, height=2,
                    command=showins).grid(row=1, column=1)
fram4.pack()

f8 = tk.Label(window, text=' ', bg='#7F7F7F', font=('等线', 3), width=400).pack()

fram7 = tk.Frame(window,bg='#7F7F7F')
text = tk.Text(fram7, width=300, height=32, borderwidth=0, bg='black')
text.pack()
text.insert('end', '---欢迎使用，请点击查看说明！！！-----{}-------\r\n'.format(time.strftime("%Y-%m-%d %H:%M", time.localtime())), 'tagg')
text.see(tk.END)
text.update()
fram7.pack()

fram9 = tk.Frame(window,bg='#7F7F7F')
submit22 = tk.Button(fram9, text='清屏', bg='#3A414B', fg='white', font=('等线', 12), width=15, height=2,
                     command=qingping).pack()
fram9.pack()

# 设置text字体
ft = tf.Font(family='微软雅黑', size=10)
text.tag_add('tagg', 'end')
text.tag_config('tagg', foreground='green', font=ft)
text.tag_add('tagr', 'end')
text.tag_config('tagr', foreground='red', font=ft)
text.tag_add('tagy', 'end')
text.tag_config('tagy', foreground='#FFFF00', font=ft)
text.tag_add('tagw', 'end')
text.tag_config('tagw', foreground='white', font=ft)

window.mainloop()