# 定义一个写入将数据excel表格的函数
import xlwt


# 定义一个函数将data中的数据写入excel表中，data是保存了所有发票信息的列表，每个元素是一个字典，含有发票的所有信息，
# 包括发票日期，纳税人识别号，购买方名称，卖方名称，购买金额
def save_vatInvoice_data(datas):
    print('正在写入数据！')
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('增值税发票内容登记', cell_overwrite_ok=True)
    # 设置表头，这里需要增加一个审批状态，如果开票日期在2016年6月12日、审批金额在2700元以内的发票视为合规，审批状态为通过，否则为不通过
    # 如果开票日期或金额为空，则审批状态为转人工
    title = ['开票日期', '纳税人识别号', '购买方名称', '卖方名称', '购买金额', '审批状态']
    for i in range(len(title)):
        sheet.write(0, i, title[i])
    for d in range(len(datas)):
        sheet.write(d + 1, 0, datas[d]['InvoiceDate'])
        sheet.write(d + 1, 1, datas[d]['SellerRegisterNum'])
        sheet.write(d + 1, 2, datas[d]['PurchasserName'])
        sheet.write(d + 1, 3, datas[d]['SellerName'])
        sheet.write(d + 1, 4, datas[d]['AmountInFiguers'])
        print(datas[d]['AmountInFiguers'])
        amount = float(datas[d]['AmountInFiguers'])
        if datas[d]['InvoiceDate'] == '' or amount == '':
            sheet.write(d + 1, 5, '转人工')
        elif datas[d]['InvoiceDate'] == '2016年06月12日' and amount <= 2700.00:
            sheet.write(d + 1, 5, '通过')
        else:
            sheet.write(d + 1, 5, '不通过')
    print('数据写入成功！')
    # 保存excel表格,保存到当前目录下
    book.save('增值税发票结果.xls')


