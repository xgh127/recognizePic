import base64
import os

import requests
import config

# 识别增值税发票的api
def get_VAT_invoice_context(pic):
    print('正在获取图片正文内容！')
    data = {}
    try:
        request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
        # 二进制方式打开图片文件
        # f = open(pic, 'rb')
        # img = base64.b64encode(f.read())
        # 对图片进行压缩，压缩后的图片大小不超过4M
        from io import BytesIO
        from PIL import Image

        # Open the image file
        with open(pic, 'rb') as f:
            img_data = f.read()

        # Open the image using Pillow
        img = Image.open(BytesIO(img_data))

        # Resize the image to have a maximum width or height of 2000 pixels
        max_size = 2000
        if max(img.size) > max_size:
            scale_factor = max_size / max(img.size)
            new_size = tuple(int(dim * scale_factor) for dim in img.size)
            img = img.resize(new_size, resample=Image.LANCZOS)

        # Compress the image and encode it in base64
        output_buffer = BytesIO()
        img.save(output_buffer, format='JPEG', optimize=True, quality=70)
        compressed_img_data = output_buffer.getvalue()

        base64_img = base64.b64encode(compressed_img_data)
        # 输出img的大小,以MB为单位
        print(len(base64_img) / 1024 / 1024)
        params = {"image": base64_img}

        # 这里需要替换成自己的access_token
        access_token = config.vat_access_token

        request_url = request_url + "?access_token=" + access_token
        headers = {'content-type': 'application/x-www-form-urlencoded'}
        response = requests.post(request_url, data=params, headers=headers)
        if response:
            # print("发票识别的所有内容")
            # print(response.json())
            json1 = response.json()
            filename = os.path.basename(pic).split('.')[0]


            data['invoiceName'] = filename
            data['SellerRegisterNum'] = json1['words_result']['SellerRegisterNum']
            data['InvoiceDate'] = json1['words_result']['InvoiceDate']
            data['PurchasserName'] = json1['words_result']['PurchaserName']
            data['SellerName'] = json1['words_result']['SellerName']
            data['AmountInFiguers'] = json1['words_result']['AmountInFiguers']
            # print(data['AmountInFiguers'])
            # 输出获取某个文件内容成功
            print('获取图片正文内容成功！'+filename)
        return data

    except Exception as e:
        # print("exception get")
        print(e)
    return data


# 写一个主函数用于测试
def main():
    pic = 'aistudio-发票数据集/b/b0.jpg'
    data = get_VAT_invoice_context(pic)
    print(data)


# 运行主函数
if __name__ == '__main__':
    main()
