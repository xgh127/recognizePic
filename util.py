# 从文件夹b中读取图片
import base64
import RecongnizePicture as rp
import os


# 根据图片编号，获取图片路径,传入图片编号和文件夹名称
def get_image_path(num, sub_folder):
    folder_path = 'aistudio-发票数据集/' + f'{sub_folder}/'
    file_name = f'b{num}.jpg'
    return folder_path + file_name


# 根据读出来的图片路径，读取图片并返回图片的base64编码
def get_image_base64_from_b(num):
    image_path = get_image_path_by_filepath(num, 'b')
    with open(image_path, 'rb') as f:
        image_base64 = base64.b64encode(f.read())
        return image_base64


# 定义函数，传入图片路径，读取图片并返回图片的base64编码
def get_image_base64_from_path(image_path):
    with open(image_path, 'rb') as f:
        image_base64 = base64.b64encode(f.read())
        return image_base64


# 读取文件夹b中的图片，返回所有的图片路径
def get_image_path_from_b(start, end):
    folder_path = 'aistudio-发票数据集/b/'
    image_path_list = []
    for i in range(start, end):
        image_path_list.append(folder_path + f'b{i}.jpg')
    return image_path_list


# 获取图片名称
def get_image_name_from_b(num):
    return f'b{num}'


# 获取当前目录某个文件夹下的所有图片路径
def get_image_path_by_filepath(curPath):
    pics = []
    for filename in os.listdir(curPath):
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = curPath + '/' + filename
            pics.append(pic)
    print('图片路径生成成功！')
    return pics


def get_batch_list(folder_path, batch_size):
    """
    批量处理指定文件夹下的所有文件，每批次读取batch_size个文件，并返回一个列表，
    其中每个元素为一个批次的文件路径列表。
    """
    file_list = os.listdir(folder_path)
    # 排序
    file_list.sort(key=lambda x: int(''.join(filter(str.isdigit, x))))
    batch_list = []
    pics = []
    for filename in file_list:
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = os.path.join(folder_path, filename).replace('\\', '/')
            pics.append(pic)
        if len(pics) == batch_size:
            batch_list.append(pics)
            pics = []
    if len(pics) > 0:
        batch_list.append(pics)
    return batch_list


def get_batch_datas(picPaths):
    datas = []
    for p in picPaths:
        data = rp.get_VAT_invoice_context(p)
        datas.append(data)
    return datas


# 定义一个测试主函数，用来测试各种函数是否正确
def main():
    # 测试能否读取图片路径
    # print(get_image_path_by_filepath(0, 'b'))
    # # 测试能否读取图片base64编码
    # print(get_image_base64_from_b(0))
    # # 测试能否读取图片名称
    # print(get_image_name_from_b(0))
    # print("test get_image_path_from_b")
    # print(get_image_path_from_b(0, 10))
    batch_files = batch_process_files('aistudio-发票数据集/b', 10)
    print(batch_files)


# 调用函数，传入文件夹路径和批量大小


# 运行主函数
if __name__ == '__main__':
    main()
