import time

import RecongnizePicture as rp
import clearAll
import mongoDao
import mysqlDao
import neo4jDao
import sendEmail
import generalMysqlDao
import util 

'''
增值税发票识别
'''

# 定义全局常量，b文件夹下的图片总数
B_TOTAL = 201
TEST_TOTAL = 10

# 传入图片路径list，遍历列表调用get_VAT_invoice_context函数获取图片内容


filePath = 'test'
bFilePath = 'aistudio-发票数据集/b'
aFilePath = 'aistudio-发票数据集/a'


def batch_get_datas_B(folder_path, batch_size):
    batch_list = util.get_batch_list(folder_path, batch_size)
    for batch in batch_list:
        datas = util.get_batch_datas(batch)
        mysqlDao.insert_many(datas)
        print('插入数据成功！')


def batch_get_datas_A(folder_path, batch_size):
    batch_list = util.get_batch_list(folder_path, batch_size)
    for batch in batch_list:
        datas = util.get_batch_datas_general(batch)
        generalMysqlDao.insertInvoiceInfoList(datas)
        print('插入数据成功！')


# 处理B文件夹下的所有图片
def process_B_folder():
    # 清空数据库中所有发票信息

    print('开始执行B！！！')
    # 执行开始时间
    start_time = time.time()
    # 批量读取图片内容，存到mysql数据库中
    batch_get_datas_B(bFilePath, 10)
    # 将图片内容存到excel表格中，分析此次审批的发票是否合规
    mysqlDao.export_to_excel()
    print('导出excel成功！')
    # 识别财务主体关系
    neo4jDao.create_all_transaction()

    # 将原始图片存到mongodb数据库中
    mongoDao.store_img_to_mongodb(bFilePath)

    # 发送邮件
    sendEmail.test_send_email_with_excel_attachment()
    # 执行结束时间
    end_time = time.time()
    print('执行结束！用时：', end_time - start_time)

def process_A_folder():
    # clearAll.clear_all() # 只在测试中使用
    print('开始执行A！！！')
    # 执行开始时间
    start_time = time.time()
    # 批量读取图片内容，存到mysql数据库中
    batch_get_datas_A(aFilePath, 10)
    # 将图片内容存到excel表格中，分析此次审批的发票是否合规
    generalMysqlDao.export_to_excel()
    print('导出excel成功！')
    # 识别财务主体关系
    neo4jDao.create_all_transaction_of_general()

    # 将原始图片存到mongodb数据库中
    mongoDao.store_gengral_img_to_mongodb(aFilePath)

    # 发送邮件
    sendEmail.test_send_general_email_with_excel_attachment()
    # 执行结束时间
    end_time = time.time()
    print('执行结束！用时：', end_time - start_time)



def main():
    # 清空数据库中所有发票信息
    clearAll.clear_all()
    process_B_folder()
    # process_A_folder()


if __name__ == '__main__':
    main()
