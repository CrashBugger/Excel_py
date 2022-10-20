import regex
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook

import xlrd
import xlutils.copy

stus_id_dict = {}  # 二课堂——分数
score_dict = {}  # 蓝色表头
pattern = regex.compile("\d.*?")


def get_score(info):
    if info:
        numbers = pattern.findall(info)
        score = int(numbers[0])
        return score
    else:
        return 2


if __name__ == '__main__':
    scores = xlrd.open_workbook("2022年10月历史文化学院十佳歌手大赛-赋分表.xls").sheets()[0]

    score_dict_sheet: Workbook = load_workbook(r"十佳歌手二课堂.xlsx").get_sheet_by_name("Sheet1")
    stus_info = [(i[1].value, i[3].value, i[2].value) for i in score_dict_sheet.rows]
    stus_info = stus_info[1:]

    for stu in stus_info:
        stus_id_dict[stu[0]] = (get_score(stu[1]), stu[2])

    start = 0
    for i in range(scores.nrows):
        if start == 0:
            pass
        else:
            row = scores.row(i)
            score_dict[int(row[0].value)] = i
        start += 1

    rd = xlrd.open_workbook("2022年10月历史文化学院十佳歌手大赛-赋分表.xls", formatting_info=True)  # 打开文件
    wt = xlutils.copy.copy(rd)  # 复制
    sheets = wt.get_sheet(0)  # 读取第一个工作表
    for id, infos in stus_id_dict.items():
        try:
            stu_row = score_dict[id]
            sheets.write(stu_row, 7, infos[0])
        except:
            print(id, infos[1])
    wt.save("2022年10月历史文化学院十佳歌手大赛-赋分表.xls")
