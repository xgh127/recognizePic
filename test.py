import os
import RecongnizePicture as rp
import store_with_excel as swe
import mysqlDao as md

# 从aistudio-发票数据集/test文件夹中读取图片路径
def get_VAT_invoice_img_from_test():
    folder_path = 'aistudio-发票数据集/test/'
    # 读取文件夹下所有以jpg结尾的文件，并将图片名和folder_path拼接成图片路径,然后调用get_VAT_invoice_context函数获取图片内容
    # 生成一个空列表用于存放图片路径
    pics = []
    # 遍历文件夹，找到后缀为jpg和png的文件，整理之后加入列表
    for filename in os.listdir(folder_path):
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = folder_path + '/' + filename
            pics.append(pic)
    print(pics)
    print('图片路径生成成功！')
    return pics


# 定义函数，用来获取图片内容，传入图片路径的列表，遍历列表调用get_VAT_invoice_context函数获取图片内容
def get_datas(pics):
    datas = []
    for p in pics:
        data = rp.get_VAT_invoice_context(p)
        datas.append(data)
    return datas


# 定义主函数，调用data_governance.py中的函数，将图片内容写入excel表格
def main():
    print('开始执行！！！')

    # 将图片内容写入excel表格
    # img_path = 'aistudio-发票数据集/test/b1.jpg'
    # datas = rp.get_VAT_invoice_context(img_path)
    datas = get_datas(get_VAT_invoice_img_from_test())
    print(datas)
    md.insert_many(datas)
    swe.save_vatInvoice_data(datas)


# 运行主函数
if __name__ == '__main__':
    main()
