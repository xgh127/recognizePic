import os

import pymongo
import util

# 连接数据库
client = pymongo.MongoClient(host='localhost', port=27017)
db = client['RPA']
originalDataB = db['originalDataB']
originalDataA = db['originalDataA']


# 定义发票图片类，属性为发票图片名称和发票图片的base64编码
class InvoiceImg:
    def __init__(self):
        self.name = ""
        self.img_base64 = ""

    # 定义构造函数，传入发票图片名称和发票图片的base64编码
    def __init__(self, name, img_base64):
        self.name = name
        self.img_base64 = img_base64


# 数据存储，将原始图片存到mongodb数据库中
def store_img_to_mongodb(curPath):
    # 定义for循环，将所有b文件夹下的图片存到mongodb数据库中
    for filename in os.listdir(curPath):
        if filename.endswith('jpg') or filename.endswith('png'):
            picPath = curPath + '/' + filename
            invoiceImg = InvoiceImg(filename, util.get_image_base64_from_path(picPath))
            res = originalDataB.insert_one(invoiceImg.__dict__)
            print(res)


# 数据存储，将原始图片存到mongodb数据库中
def store_gengral_img_to_mongodb(curPath):
    # 定义for循环，将所有b文件夹下的图片存到mongodb数据库中
    for filename in os.listdir(curPath):
        if filename.endswith('jpg') or filename.endswith('png'):
            picPath = curPath + '/' + filename
            invoiceImg = InvoiceImg(filename, util.get_image_base64_from_path(picPath))
            res = originalDataA.insert_one(invoiceImg.__dict__)
            print(res)


# 定义一个函数，删除数据库中的所有数据
def delete_all():
    res = originalDataB.delete_many({})
    print(res)



#定义一个函数，删除数据库中的所有数据
def delete_all_general():
    res = originalDataA.delete_many({})
    print(res)


def test_insert():
    invoiceImg = InvoiceImg('b300', 'this is a test')
    # 将发票图片对象插入数据库
    res = originalDataB.insert_one(invoiceImg.__dict__)


# 写一个测试主函数，用来测试各种函数是否正确
def main():
    # insert_one(0)
    # store_img_to_mongodb('aistudio-发票数据集/test')
    # delete_all()
    # test_insert()
    store_gengral_img_to_mongodb('aistudio-发票数据集/test')
    # originalDataA.delete_many({})


# 运行主函数
if __name__ == '__main__':
    main()
