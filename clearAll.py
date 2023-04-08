# 定义主函数，做mongoDB,mysql的数据清理
import mongoDao
import mysqlDao
import neo4jDao


def main():
    # 清空数据库中所有发票信息
    mysqlDao.delete_all()
    # 清空mongodb中所有发票信息
    mongoDao.delete_all()
    # 删除excel表格
    mysqlDao.delete_excel()
    # 清空neo4j中所有节点和关系
    neo4jDao.clear_all()
    # 打印清空成功
    print("清空成功")


# 执行主函数
if __name__ == '__main__':
    main()
