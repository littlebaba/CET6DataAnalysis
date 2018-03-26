# -*- coding: utf-8 -*-
import xlrd
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签

plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
from collections import defaultdict


class BaseLookCET(object):
    def __init__(self, file_path):
        self.file_path = file_path
        self.book = xlrd.open_workbook(self.file_path)
        self.sheet_num = self.book.nsheets;
        self.sheet_names = self.book.sheet_names()
        print("-----------------------------------------------------------------------------------")
        print("--excel的表格个数为 {0}\n--各表的名称为 {1}\n".format(self.sheet_num, self.sheet_names))

    def base_pie(self, ls,title):
        font = matplotlib.font_manager.FontProperties(fname='C:\Windows\Fonts\msyh.ttc')
        explode = [0, 0.05]
        plt.axes(aspect=1)
        plt.pie(labels=ls[0], x=ls[1], autopct='%.0f%%', explode=explode, startangle=90)
        # plt.legend(prop=font)
        plt.title(title)
        plt.show()




class LookCET6(BaseLookCET):
    def __init__(self, file_path):
        super(LookCET6, self).__init__(file_path)

    def get_gender_pie(self):
        sh = self.book.sheet_by_index(0)
        sexNum = {'man': 0, 'woman': 0, 'all': 0}
        for v in range(1, sh.nrows):
            if sh.cell_value(v, 15) == '男':
                sexNum['man'] += 1
            else:
                sexNum['woman'] += 1
        sexNum['all'] = sexNum['man'] + sexNum['woman']
        print("--男生人数为 {0}  女生人数为 {1}".format(sexNum['man'], sexNum['woman']))

        labels = [u'男', u'女']
        fracs = [sexNum['man'], sexNum['woman']]
        ls = [labels, fracs]
        title="CET-6男女生比例"
        self.base_pie(ls,title)

    def get_student_category_pie(self):
        sh = self.book.sheet_by_index(0)
        stu_cagy = {'b': 0, 'y': 0, 'all': 0}
        for v in range(1, sh.nrows):
            if sh.cell_value(v, sh.ncols - 3) == '本科':
                stu_cagy['b'] = stu_cagy['b'] + 1
            else:
                stu_cagy['y'] = stu_cagy['y'] + 1
        stu_cagy['all'] = stu_cagy['b'] + stu_cagy['y']
        print(stu_cagy)
        labels = ['本科', '硕士']
        fracs = [stu_cagy['b'], stu_cagy['y']]
        ls = [labels, fracs]
        title = "CET-6本硕比例"
        self.base_pie(ls, title)

    def get_l_success(self,str='本科',title = "CET-6本科达线比例"):
        sh = self.book.sheet_by_index(0)
        low_suc = {'suc': 0, 'fail': 0, 'all': 0}
        for v in range(1, sh.nrows):
            if sh.cell_value(v, sh.ncols - 3) == str:
                low_suc['all']+=1
                if int(sh.cell_value(v, 17)) >= 425:
                    low_suc['suc'] += 1
                else:
                    low_suc['fail'] +=1
        print(low_suc)
        label=['达线','未达线']
        fracs=[low_suc['suc'],low_suc['fail']]
        ls=[label,fracs]
        self.base_pie(ls, title)

    def get_h_success(self):
        self.get_l_success('研究生',title = "CET-6本科达线比例")

    def get_per_pro_mw(self,flag=True,title='各省男女CET-6达线预览'):
        sh = self.book.sheet_by_index(0)
        area={"11":"北京","12":"天津","13":"河北","14":"山西","15":"内蒙古","21":"辽宁","22":"吉林","23":"黑龙江","31":"上海","32":"江苏","33":"浙江","34":"安徽","35":"福建","36":"江西","37":"山东","41":"河南","42":"湖北","43":"湖南","44":"广东","45":"广西","46":"海南","50":"重庆","51":"四川","52":"贵州","53":"云南","54":"西藏","61":"陕西","62":"甘肃","63":"青海","64":"宁夏","65":"新疆","71":"台湾","81":"香港","82":"澳门","91":"国外"}
        areaMen=defaultdict(int)
        areaWomen=defaultdict(int)
        for v in range(1, sh.nrows):
            id_num = sh.cell_value(v,16)
            pro_num = id_num[0:2]
            for key,value in area.items():
                if pro_num == key:
                    if sh.cell_value(v,15) == '女':
                        if flag:
                            if int(sh.cell_value(v, 17)) >= 425:
                                areaMen[value]+=1
                        else:
                            areaMen[value] += 1
                    if sh.cell_value(v, 15) == '男':
                        if flag:
                            if int(sh.cell_value(v, 17)) >= 425:
                                areaWomen[value]+=1
                        else:
                            areaWomen[value] += 1


        for k1,v1 in areaMen.items():
            if k1 not in list(areaWomen.keys()):
                areaWomen[k1]=0
        for k2,v2 in areaWomen.items():
            if k2 not in list(areaMen.keys()):
                areaMen[k2]=0

        pro_name=sorted(areaWomen.keys())
        men_s_num=list()
        women_s_num = list()
        for x in pro_name:
            men_s_num.append(areaMen[x])
            women_s_num.append(areaWomen[x])
        ind = np.arange(0,4*pro_name.__len__(),4)  # the x locations for the groups
        # width = 0.35
        width = 0.65
        p1 = plt.bar(ind, men_s_num, width)
        p2 = plt.bar(ind, women_s_num, width,
                     bottom=men_s_num)
        plt.ylabel('达线人数')
        plt.title(title)
        plt.xticks(ind, pro_name, rotation=-30)
        plt.yticks(np.arange(0, 181, 10))
        plt.legend((p1[0], p2[0]), ('男', '女'))

        plt.show()


if __name__ == '__main__':
    file_path = 'CET6DataAnalysis/172ljcj.xls'
    cet6 = LookCET6(file_path)
    cet6.get_gender_pie()
    cet6.get_student_category_pie()
    cet6.get_l_success()
    cet6.get_h_success()
    cet6.get_per_pro_mw()

