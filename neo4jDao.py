import py2neo
from py2neo import Relationship, NodeMatcher
from py2neo import Node

import generalMysqlDao
import mysqlDao
import config

# 定义全局的Neo4j连接对象
graph = py2neo.Graph(config.neo4j_uri, auth=(config.neo4j_username, config.neo4j_password))


def create_or_update_node(label, primary_key, data):
    """
    根据指定的标签和主键，创建或更新节点，同时设置节点属性。
    如果节点已存在，则更新节点属性；否则创建新节点。
    """
    node = Node(label, **data)
    node.__primarylabel__ = label
    node.__primarykey__ = primary_key
    graph.merge(node, label, primary_key)
    return node


def create_transaction_relation(payer, payee, invoice_id, amount, time):
    """
    创建交易关系，箭头由payer指向payee，含有交易金额和对应的发票图片名称属性
    """
    payer_name = payer['name']
    payee_name = payee['name']
    # 创建关系
    query = "MATCH (p:Company {name: $payer}), (q:Company {name: $payee}) " \
            "CREATE (p)-[r:Trade {amount: $amount, invoice_id: $invoice_id, time: $time}]->(q)"
    graph.run(query, payer=payer_name, payee=payee_name, amount=amount, invoice_id=invoice_id, time=time)


# 清楚数据库中所有节点和关系
def clear_all():
    graph.delete_all()


# 清楚数据库中和浙江大学有关的节点和关系
def clear_zju():
    graph.run("MATCH (n) WHERE n.name CONTAINS '浙江大学' DETACH DELETE n")


# 进一步封装测试函数，传入买方，卖方，发票号，金额，创建节点和关系
def create_transaction(payer, payee, invoice_id, amount, time):
    payer_data = {"name": payer}
    payee_data = {"name": payee}
    payer_node = create_or_update_node("Company", "name", payer_data)
    payee_node = create_or_update_node("Company", "name", payee_data)
    create_transaction_relation(payer_node, payee_node, invoice_id, amount, time)


# 定义函数，传入datas，内保存着发票信息，插入所有的关系
def create_all_transaction():
    # 从mysql中获取所有的发票信息
    datas = mysqlDao.select_all()
    # 遍历datas,创建Transaction关系
    print(len(datas))
    for d in range(len(datas)):
        # 获取发票信息
        payer = datas[d][4]
        payee = datas[d][5]
        invoice_id = datas[d][1]
        time = datas[d][2]
        # 获取的金额需要/100并保留两位小数
        amount = float(datas[d][6]) / 100
        # 保留两位小数
        amount = round(amount, 2)
        # 创建节点和关系
        create_transaction(payer, payee, invoice_id, amount, time)


# 定义一个创建事件的函数
def create_all_transaction_of_general():
    # 从mysql中获取所有的发票信息,要求PurchasserName和SalesName不为空，接status不等于'转人工'
    datas = generalMysqlDao.queryInvoiceInfoOfTransaction()
    # 遍历datas,创建Transaction关系
    # print(len(datas))
    for d in range(len(datas)):
        # 获取发票信息
        payer = datas[d][3]
        payee = datas[d][4]
        invoice_id = datas[d][1]
        time = datas[d][2]
        # 获取的金额需要/100并保留两位小数
        amount = float(datas[d][5]) / 100
        # 保留两位小数
        amount = round(amount, 2)
        # 创建节点和关系
        create_transaction(payer, payee, invoice_id, amount, time)


# 定义主函数，测试Neo4j数据库的连接
def main():
    clear_all()
    #create_transaction('深圳市购机汇网络有限公司', '广州晶东贸易有限公司', 'b1', 3495, '2014-01-01 00:00:00')

    create_all_transaction()
    create_all_transaction_of_general()
    # clear_all()



# 执行主函数
if __name__ == '__main__':
    main()
