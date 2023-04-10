import base64
import urllib
import requests
import os
from PIL import Image
import xlwt
import datetime
import util

def pics(path):
    print('正在生成图片路径')
    # 生成一个空列表用于存放图片路径
    pics = []
    # 遍历文件夹，找到后缀为jpg和png的文件，整理之后加入列表
    for filename in os.listdir(path):
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = path + '/' + filename
            pics.append(pic)
    print('图片路径生成成功！')
    return pics

def datas(pics):
    datas = []

    for p in pics:
        data = get_context(p)
        datas.append(data)
    return datas

def get_context(pic):
    print('正在获取图片正文内容！')
    data = {}
    # 二进制
    img_file = pic
    im = Image.open(img_file)
    (x, y) = im.size
    file_name = pic.split('/')[2]  # TODO:注意换成数据集a时要改为2
    file_name = file_name.split('.')[0]
    max_size = 2000


    if max(im.size) > max_size:
        scale_factor = max_size / max(im.size)
        new_size = tuple(int(dim * scale_factor) for dim in im.size)
        im = im.resize(new_size, resample=Image.LANCZOS)
        im.save(pic, format='JPEG', optimize = True, quality=70)
    f = open(pic, 'rb')
    img = base64.b64encode(f.read())
    try:
        request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/taxi_receipt?access_token="
        params = {"image": img}
        access_token = '24.28996d3cf44704a4c04759df066806ed.2592000.1683466812.282335-32066502'
        request_url = request_url + access_token
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        response = requests.post(request_url, data=params, headers=headers)
        if response:
            json1 = response.json()
            data['InvoiceNum'] = json1['words_result']['InvoiceNum']
            data['TaxiNum'] = json1['words_result']['TaxiNum']
            data['PickupTime'] = json1['words_result']['PickupTime']
            data['DropoffTime'] = json1['words_result']['DropoffTime']
            data['TotalFare'] = json1['words_result']['TotalFare']
            data['Province'] = json1['words_result']['Province']
            data['City'] = json1['words_result']['City']
            data['PricePerkm'] = json1['words_result']['PricePerkm']
            data['Distance'] = json1['words_result']['Distance']

        print('正文内容获取成功')
        return data

    except Exception as e:
        print(e)
        print('不能识别')
        im.save("aistudio-发票数据集/c/"+ file_name+ ".jpg", format='JPEG', optimize = True, quality=70)
    return data

def data_save(datas):
    print('正在写入数据！')
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('出租车', cell_overwrite_ok=True)
    # 创建样式对象，初始化样式
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    # 设置水平居中对齐
    alignment.horz = 2
    # 为样式创建字体
    font = xlwt.Font()
    # 设置字体
    font.name = 'Calibri'
    # 字体大小
    font.height = 200
    style.font = font
    style.alignment = alignment
    title = ['发票号码', '出租车号', '上车时间', '下车时间', '总花费', '省份', '城市', '单价', '总里程']
    num = 0
    for i in range(len(title)):
        sheet.col(i).width = 7777
        sheet.write(0, i, title[i],style)
    for d in range(len(datas)):
        if datas[d]:
            print("写入")

            sheet.write(num + 1, 0, datas[d]['InvoiceNum'],style)
            sheet.write(num + 1, 1, datas[d]['TaxiNum'],style)
            sheet.write(num + 1, 2, datas[d]['PickupTime'],style)
            sheet.write(num + 1, 3, datas[d]['DropoffTime'],style)
            sheet.write(num + 1, 4, datas[d]['TotalFare'],style)
            sheet.write(num + 1, 5, datas[d]['Province'], style)
            sheet.write(num + 1, 6, datas[d]['City'], style)
            sheet.write(num + 1, 7, datas[d]['PricePerkm'], style)
            sheet.write(num + 1, 8, datas[d]['Distance'], style)
            num = num + 1
        else:
            continue

    print('数据写入成功！')
    now = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
    book.save(now+'出租车发票.xls')
    return



# 用于测试
def main():
    # path = 'aistudio-发票数据集/a'
    #
    # Pics = pics(path)
    #
    # Datas = datas(Pics)
    #
    # data_save(Datas)
    # storage.storage_invoice_approval(Datas)
    content = get_context('aistudio-发票数据集/a/a0.jpg')
    # data_save(content)
    print(content)

    # storage.storage_pics(path)
    print('执行结束！')

if __name__ == '__main__':
    main()