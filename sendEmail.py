import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import zipfile
import os
import config
import generalMysqlDao
import mysqlDao
import main

def send_email_with_excel_and_image_attachment(email_sender, email_password, email_receiver, subject, body, excel_file_path):
    """
    发送带有Excel附件的邮件
    :param email_sender: 发件人邮箱地址
    :param email_password: 发件人邮箱密码或授权码
    :param email_receiver: 收件人邮箱地址
    :param subject: 邮件主题
    :param body: 邮件正文
    :param excel_file_path: Excel文件路径
    """
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_receiver
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    imgs = mysqlDao.select_invoice_of_email()

    with open(excel_file_path, 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='xlsx')
        attach.add_header('Content-Disposition', 'attachment', filename=excel_file_path)
        msg.attach(attach)

    if imgs:
        for img in imgs:
            img = list(img)[0]
            img.replace("'", '')
            picPath = main.bFilePath + '/' + img + '.jpg'
            attach = MIMEApplication(open(picPath, 'rb').read())
            attach.add_header('Content-Disposition', 'attachment', filename=img + '.jpg')
            msg.attach(attach)

    server = smtplib.SMTP('smtp.qq.com', 587)
    server.starttls()
    server.login(email_sender, email_password)
    text = msg.as_string()
    server.sendmail(email_sender, email_receiver, text)
    server.quit()



def send_email_with_excel_and_image_attachment_general(email_sender, email_password, email_receiver, subject, body, excel_file_path):
    """
    发送带有Excel附件的邮件
    :param email_sender: 发件人邮箱地址
    :param email_password: 发件人邮箱密码或授权码
    :param email_receiver: 收件人邮箱地址
    :param subject: 邮件主题
    :param body: 邮件正文
    :param excel_file_path: Excel文件路径
    """
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_receiver
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    imgs = generalMysqlDao.select_invoice_of_email()
    with open(excel_file_path, 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='xlsx')
        attach.add_header('Content-Disposition', 'attachment', filename=excel_file_path)
        msg.attach(attach)

    if imgs:
        for img in imgs:
            img = list(img)[0]
            img.replace("'", '')
            picPath = main.aFilePath + '/' + img + '.jpg'
            attach = MIMEApplication(open(picPath, 'rb').read())
            attach.add_header('Content-Disposition', 'attachment', filename=img + '.jpg')
            msg.attach(attach)



    server = smtplib.SMTP('smtp.qq.com', 587)
    server.starttls()
    server.login(email_sender, email_password)
    text = msg.as_string()
    server.sendmail(email_sender, email_receiver, text)
    server.quit()

# TODO: 将全部图片压缩成zip文件
def zip_dir(dir_path, outFullName):
    """
    压缩指定文件夹
    :param dir_path: 目标文件夹路径
    :param outFullName:  压缩文件保存路径+XXXX.zip
    :return:
    """
    testcase_zip = zipfile.ZipFile(outFullName, 'w', zipfile.ZIP_DEFLATED)
    for path, dir_names, file_names in os.walk(dir_path):
        for filename in file_names:
            testcase_zip.write(os.path.join(path, filename))
    testcase_zip.close()
    print("打包成功")


# test send email with excel attachment
def test_send_general_email_with_excel_attachment():
    email_sender = config.email_sender
    email_password = config.email_password
    email_receiver = config.email_receiver
    subject = config.subject
    body = '发票a处理结果'
    excel_file_path = config.excel_filename_general
    send_email_with_excel_and_image_attachment_general(email_sender, email_password, email_receiver, subject, body, excel_file_path)
    print("发送邮件成功！")


def test_send_email_with_excel_attachment():
    email_sender = config.email_sender
    email_password = config.email_password
    email_receiver = config.email_receiver
    subject = config.subject
    body = '发票b处理结果'
    excel_file_path = config.excel_filename
    send_email_with_excel_and_image_attachment(email_sender, email_password, email_receiver, subject, body, excel_file_path)
    print("发送邮件成功！")


if __name__ == '__main__':
    test_send_email_with_excel_attachment()
