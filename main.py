import time

import RecongnizePicture as rp
import mongoDao
import mysqlDao
import neo4jDao
import util

'''
增值税发票识别
'''

# 定义全局常量，b文件夹下的图片总数
B_TOTAL = 201
TEST_TOTAL = 10


# 传入图片路径list，遍历列表调用get_VAT_invoice_context函数获取图片内容
def get_datas(pics):
    datas = []
    for p in pics:
        data = rp.get_VAT_invoice_context(p)
        datas.append(data)
    return datas


filePath = 'aistudio-发票数据集/test'


def main():
    print('开始执行！！！')
    # 执行开始时间
    start_time = time.time()
    # 1.数据采集
    # 从文件夹b中读取图片
    image_path_list = util.get_image_path_by_filepath(filePath)
    # 获取所有图片内容
    datas = get_datas(image_path_list)

    # 2.数据治理
    # 将图片内容存到mysql数据库中
    mysqlDao.insert_many(datas)
    # 分析和应用
    # 将图片内容存到excel表格中，分析此次审批的发票是否合规
    mysqlDao.export_to_excel()
    print('导出excel成功！')
    # 识别财务主体关系
    neo4jDao.create_all_transaction()

    # 数据存储
    # 将原始图片存到mongodb数据库中
    mongoDao.store_img_to_mongodb(filePath)
    # 执行结束时间
    end_time = time.time()
    print('执行结束！用时：', end_time - start_time)


if __name__ == '__main__':
    main()
