from openpyxl import Workbook

title = ["院校领导", "院校设置", "实验室",
         "学科", "可授予的学位", "院校人员数量",
         "硕士专业及招生人数", "博士专业及招生人数",
         "分数线"]
labellist = {
    "院校基本信息" : ['序号', '院校名称', '所在地', '院校隶属', '研究生院', '自划线院校', '院校简介', '周边环境', '院校官网地址', '研究生院官网地址'],
    title[0]: ['序号', '院校名称', '党委书记', '党委副书记', '党委常委', '纪委书记', '校长', '副校长', '校长助理', '院长', '副院长'],
    title[1]: ['院校名称', '院系名称', '院系介绍', '院系官网地址'],
    title[2]: ['院校名称','实验室名称','实验室级别','实验室官网地址'],
    title[3]: ['院校名称','一级学科名称','一级学科代码','二级学科名称','二级学科代码','学科级别'],
    title[4]: ['院校名称','学科代码','学科名称','学位类别'],
    title[5]: ['院校名称','专任教师','本科生人数','硕士研究生人数','博士研究生人数','留学生人数'],
    title[6]: ['院校名称','年份','学科大类','专业名称','专业代码','招生人数'],
    title[7]: ['院校名称','年份','学科大类','专业名称','专业代码','招生人数'],
    title[8]: ['院校名称','年份','专业名称','专业代码','总分数线','科目一分数线','科目二分数线','科目三分数线','科目四分数线','国控线']
}

wb = Workbook()
ws = wb.active
ws.title = "院校基本信息"
ws.append(labellist["院校基本信息"])
for i in title:
    wb.create_sheet(title=i)
    ws = wb[i]
    ws.append(labellist[i])

wb.save("四川省考研学校收集表.xlsx")