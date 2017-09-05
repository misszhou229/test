# test
#open to excel
#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2017/9/5 10:26
# @Author  : LiuLiJun
# @Site    : 
# @File    : main.py
# @Software: PyCharm

import copy
import pandas as pd

class demo:

    def __init__(self,parent=None):

        self.farm_setting()
        self.warning_stas()
        self.model_setting()
        self.area_info()
        self.relate()
        self.export()

    def farm_setting(self):

        self.farm_info = pd.read_excel("./info/风场信息设置.xlsx", sheetname="风场信息")

        self.farm_info.index = self.farm_info['farm_code'].tolist()

        self.wtgs_info = pd.read_excel("./info/风场信息设置.xlsx", sheetname="风机信息")

        self.wtgs_info.index = self.wtgs_info['CODE_'].tolist()

        self.wtgs_info_bias_name=copy.deepcopy(self.wtgs_info)

        self.wtgs_info_bias_name.index = self.wtgs_info['LOCATION_CODE'].tolist()

    def warning_stas(self):

        self.warning_stas = pd.read_excel("./info/新风机预警统计new.xls", sheetname="dv_predict")

    def model_setting(self):

        self.model_info = pd.read_excel("./info/dv_model(模型表).xlsx", sheetname="Sheet1")

        self.model_info.index = self.model_info['NAME(名称)'].tolist()

        self.model_level={'等级一':1,'等级二':2,'等级三':3,'等级四':4,'等级五':5}

    def area_info(self):

        self.area_info = pd.read_excel("./info/dv_area(区域表).xlsx", sheetname="Sheet1")

        self.area_info.index = self.area_info['CODE(编号)'].tolist()

    def relate(self):

        for i in range(len(self.warning_stas)):

            print(i,self.warning_stas['PROJECT(风场code)'].iloc[i],self.farm_info['AREA_code'].loc[self.warning_stas['PROJECT(风场code)'].iloc[i]])

            if self.farm_info['AREA_code'].loc[self.warning_stas['PROJECT(风场code)'].iloc[i]]:

                area_code=int(self.farm_info['AREA_code'].loc[self.warning_stas['PROJECT(风场code)'].iloc[i]])

                self.warning_stas['AREA(区域code)'].iloc[i]=area_code

                self.warning_stas['AREA_NAME(区域名称)'].iloc[i] =self.area_info['NAME(名称)'].loc[area_code]

            tempdf=self.wtgs_info[self.wtgs_info['FARM_CODE']==self.warning_stas['PROJECT(风场code)'].iloc[i]]

            tempdf.index = tempdf['LOCATION_CODE'].tolist()

            self.warning_stas['TURBINE(风机code)'].iloc[i] = str(tempdf['CODE_'].loc[self.warning_stas['LOCATION_CODE(机位号)'].iloc[i]])[0:8]

            self.warning_stas['TURBINE_MODEL(机型)'].iloc[i] = tempdf['MODEL_DETAIL'].loc[self.warning_stas['LOCATION_CODE(机位号)'].iloc[i]][0:5]

            if self.warning_stas['PREDICT_MODEL_NAME(预警模型名称)'].iloc[i] in self.model_info.index:

                self.warning_stas['PREDICT_MODEL(预警模型id)'].iloc[i] = self.model_info['ID'].loc[self.warning_stas['PREDICT_MODEL_NAME(预警模型名称)'].iloc[i]]

                self.warning_stas['COMP_RELATED(关联部件)'].iloc[i] = self.model_info['COMP_RELATED(关联部件)'].loc[self.warning_stas['PREDICT_MODEL_NAME(预警模型名称)'].iloc[i]]

                self.warning_stas['LEVEL(预警等级:1-5)'].iloc[i] = self.model_level[self.model_info['LEVEL(等级:1-5)'].loc[self.warning_stas['PREDICT_MODEL_NAME(预警模型名称)'].iloc[i]]]


    def export(self):

        self.warning_stas.to_csv('./info/新风机预警统计related.csv')


if __name__=="__main__":
    a=demo()
