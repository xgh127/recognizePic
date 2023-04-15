import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import zipfile
import os

# from PIL.Image import msg

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
    发送带有Excel附件和图片文件夹附件的邮件
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

    # 添加Excel附件
    with open(excel_file_path, 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='xlsx')
        attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(excel_file_path))
        msg.attach(attach)

    # 添加图片文件夹附件
    imgs = generalMysqlDao.select_invoice_of_email()
    if imgs:
        # 创建一个临时文件夹用于保存图片文件
        temp_folder_path = 'temp_images'
        if not os.path.exists(temp_folder_path):
            os.makedirs(temp_folder_path)
        else:
            # 如果文件夹已存在，则清空文件夹中的文件
            for root, dirs, files in os.walk(temp_folder_path):
                for file in files:
                    os.remove(os.path.join(root, file))

        for img in imgs:
            img = img[0].replace("'", '')
            picPath = os.path.join(main.aFilePath, img + '.jpg')
            # 将图片文件复制到临时文件夹中
            shutil.copy2(picPath, os.path.join(temp_folder_path, img + '.jpg'))

        # 使用zipfile模块将临时文件夹压缩成zip文件
        zip_file_path = 'images.zip'
        with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    # 将文件内容读取为 bytes 类型，并指定编码方式
                    with open(file_path, 'rb') as f:
                        file_content = f.read()
                        zf.writestr(os.path.relpath(file_path, temp_folder_path), file_content)

        # 将压缩文件作为附件添加到邮件中
        with open(zip_file_path, 'rb') as f:
            attach = MIMEApplication(f.read(), _subtype='zip')
            attach.add_header('Content-Disposition', 'attachment', filename='images.zip')
            msg.attach(attach)

        # 删除临时文件夹和压缩文件
        shutil.rmtree(temp_folder_path)
        os.remove(zip_file_path)

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
    body = '这次一共处理' + str(generalMysqlDao.countTable()) + '个发票，通过个数一共' + str(generalMysqlDao.count_pass()) + '个，没通过个数一共' + str(generalMysqlDao.count_not_pass()) + '个,' + '转人工一共' +  str(generalMysqlDao.count_to_human())
    excel_file_path = config.excel_filename_general
    send_email_with_excel_and_image_attachment_general(email_sender, email_password, email_receiver, subject, body, excel_file_path)
    print("发送邮件成功！")


def test_send_email_with_excel_attachment():
    email_sender = config.email_sender
    email_password = config.email_password
    email_receiver = config.email_receiver
    subject = config.subject
    body = '这次一共处理' + str(mysqlDao.count_all()) + '个发票，通过个数一共' + str(mysqlDao.count_pass()) + '个，没通过个数一共' + str(mysqlDao.count_not_pass()) + '个,' + '转人工一共' +  str(mysqlDao.count_to_human())
    excel_file_path = config.excel_filename
    send_email_with_excel_and_image_attachment(email_sender, email_password, email_receiver, subject, body, excel_file_path)
    print("发送邮件成功！")


if __name__ == '__main__':
    test_send_email_with_excel_attachment()
