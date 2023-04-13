# 定义主函数，做mongoDB,mysql的数据清理
import mongoDao
import mysqlDao
import neo4jDao
import generalMysqlDao


# 定义清除所有数据的函数
def clear_all():
    # 清空数据库中所有发票信息
    mysqlDao.delete_all()
    generalMysqlDao.delete_all()
    print("mysql清空成功")
    # 清空mongodb中所有发票信息
    mongoDao.delete_all()
    mongoDao.delete_all_general()
    print("mongodb清空成功")
    mysqlDao.delete_excel()
    generalMysqlDao.delete_excel()
    print("excel删除成功")
    # 清空neo4j中所有节点和关系
    neo4jDao.clear_all()
    print("neo4j清空成功")
    # 打印清空成功
    print("全部清空成功")


def main():
    # 清空数据库中所有发票信息
    mysqlDao.delete_all()
    print("mysql清空成功")
    # 清空mongodb中所有发票信息
    mongoDao.delete_all()
    print("mongodb清空成功")
    # 删除excel表格
    mysqlDao.delete_excel()
    print("excel删除成功")
    # 清空neo4j中所有节点和关系
    neo4jDao.clear_all()
    print("neo4j清空成功")
    # 打印清空成功
    print("全部清空成功")


# 执行主函数
if __name__ == '__main__':
    main()
