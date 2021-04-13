import pymysql
from docx import Document
import win32com.client as wc
import os
import xlwt


class Method:
    def __init__(self):
        self.info = "success"

    def remove_brackets(self, list_):
        """
        去除括号

        :param list_: 带有括号的list
        :return: 没有括号的list
        :rtype: list
        """

        res = []
        for item in list_:
            res.append(item[0])
        self.info = "Successfully remove brackets! "
        return res

    def de_duplication(self, list_):
        """list去重

        :param list_: 去重前的list
        :return: 去重后的list
        :rtype: list
        """
        self.info = "Successfully de-duplication! "
        return list(set(list_))

    def compare(self, list_a, list_b):
        """
        list比较

        :param list_a: 文件读取的list
        :type list_a: list
        :param list_b: 数据库读取的list
        :type list_b: list
        :return: list_a中有但是list_b中没有的结果, list
        :rtype: list
        """
        self.info = "Successfully compare! "
        return list(set(list_a) - set(list_b))

    def console(self, info):
        self.info = "Successfully print! "
        print(info)

    def output_to_excel(self, list_, path):
        """
        将比较结果输出至excel表格

        :param list_: 结果list
        :type list_: list
        :param path: 输出路径
        :type path: str
        """
        list_.sort()
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('结果')
        longest_ip = max(list_, key=len)

        for i in range(list_.__len__()):
            worksheet.write(i + 1, 0, label=list_[i])
        worksheet.col(0).width = 256 * (longest_ip.__len__() + 1)
        workbook.save(path)
        self.info = "Successfully output! "


class SQL:
    def __init__(self):
        self.select_cyy = "select * from ip_record where area = '产业园'"
        self.select_jt = "select * from ip_record where area = '集团'"
        self.select_dw = "select * from ip_record where area = '代王'"
        self.select_cyy_ip = "select ip from ip_record where area = '产业园'"
        self.select_jt_ip = "select ip from ip_record where area = '集团'"
        self.select_dw_ip = "select ip from ip_record where area = '代王'"

    def search_cyy(self):
        """产业园所有数据的查询语句"""
        return self.select_cyy

    def search_jt(self):
        """集团所有数据的查询语句"""
        return self.select_jt

    def search_dw(self):
        """代王所有数据的查询语句"""
        return self.select_dw

    def search_cyy_ip(self):
        """产业园所有IP的查询语句"""
        return self.select_cyy_ip

    def search_jt_ip(self):
        """集团所有IP的查询语句"""
        return self.select_jt_ip

    def search_dw_ip(self):
        """代王所有IP的查询语句"""
        return self.select_dw_ip


class Docx:
    def __init__(self):
        self.length = 0

    def get_data(self, path):
        """读取文件

        :param path: 文件路径
        :type path: str

        :return: 一个ip的list
        :rtype: list
        """

        res = []
        document = Document(path)
        for p in document.paragraphs:
            if p.text.__len__() != 0:
                res.append(p.text.strip())
        self.length = res.__len__()
        return res

    def get_data_de_mac(self, path):
        """读取文件, 去掉Mac地址

        :param path: 文件路径
        :type path: str

        :return: 一个ip的list
        :rtype: list
        """

        flag = False
        # 如果是.doc文件
        if path[-1] == "c":
            # 获取当前路径
            c_path = os.getcwd()
            # 拼接为绝对路径
            a_path = c_path + "\\" + path.replace("/", "\\")

            word = wc.Dispatch("Word.Application")
            doc = word.Documents.Open(a_path)
            doc.SaveAs(a_path + "x", 12)
            word.Quit
            doc.close
            path = path + 'x'
            flag = True

        res = []
        document = Document(path)
        for p in document.paragraphs:
            if p.text.__len__() != 0:
                res.append(p.text.strip().split(" ")[0])
        self.length = res.__len__()

        if flag:
            os.remove(path)

        return res


class IPCheck:
    def __init__(self):
        self.host = "10.2.0.40"
        self.user = "root"
        self.password = "ipam@zll2020."
        self.database = "ipam"
        self.charset = "utf8"

    def execute_search(self, sql_select):
        """执行MySQL查询语句

        :param sql_select: sql查询语句
        :type sql_select: str

        :return: 返回查询结果的list
        :rtype: list
        """

        result = []
        conn = pymysql.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database,
            charset=self.charset
        )

        cursor = conn.cursor()

        cursor.execute(sql_select)

        res = cursor.fetchall()

        for r in res:
            result.append(r)

        cursor.close()

        conn.close()

        return result


def main():
    sql = SQL()
    docx = Docx()
    method = Method()
    ip_check = IPCheck()
    output_path = "Excel/dw.xls"
    path_ip = "Doc/ip-2021-4-12-dw.docx"

    # 获取文件中的所有ip
    ip_docx_list = method.de_duplication(docx.get_data_de_mac(path_ip))

    # 获取数据库中的所有ip
    ip_db_list = method.de_duplication(method.remove_brackets(ip_check.execute_search(sql.search_dw_ip())))

    method.console("文件: ")
    print(ip_docx_list.__len__(), ip_docx_list)
    method.console("数据库: ")
    print(ip_db_list.__len__(), ip_db_list)

    # 比较文件中的ip和数据库中的ip
    res = method.compare(ip_docx_list, ip_db_list)
    method.console("文件有数据库没有: ")
    print(res.__len__(), res)
    # 将结果输出至excel表格
    method.output_to_excel(res, output_path)

    res_t = method.compare(ip_db_list, ip_docx_list)
    method.console("数据库有文件没有: ")
    print(res_t.__len__(), res_t)


if __name__ == '__main__':
    test_sql = "select * from ip_record where area = '代王' and ip = '10.26.0.8'"

    main()
