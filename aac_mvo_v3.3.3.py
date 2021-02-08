#! python3
# -*- coding: utf-8 -*-

import os
import time
import copy
import scipy.optimize as sco
import scipy.stats as sct
import numpy as np
import pandas as pd
import pandas.io.sql as pdsql
from sqlalchemy import create_engine
import re
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.style
import matplotlib as mpl

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side, Color, PatternFill


mpl.style.use('classic')

# font_name = fm.FontProperties(fname="c:/Windows/Fonts/malgun.ttf").get_name()
# mpl.rc('font', family=font_name)

os.environ["NLS_LANG"] = ".AL32UTF8"  # oracle charset utf설정
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

os.environ["NLS_LANG"] = ".AL32UTF8"  # oracle charset utf설정

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

pd.set_option('display.float_format', lambda x: '%.15f' % x)


""" 자산배분 MVO(Mean-Varience Optimization) 모델
    https://m.blog.naver.com/koko8624/221256556961
    
    * 용어설명
    ShortFall - https://blog.naver.com/gis-rh/220938570739
"""


class AssetAllocationMVO(object):

    def __init__(self):
        self.egn_fnfund_dev = create_engine('oracle://uffdba:venus2006@10.10.1.23:1522/FNFUND')  # 펀드 개발DB
        # self.engine_fnfund = create_engine('oracle://uffdba:venus2006@10.10.1.60:1522/FNFUND')  # 펀드 운영DB
        # self.engine_fndb2 = create_engine('oracle://ufngdba:venus2002@10.10.1.50:1521/FNDB2')  # 마켓 운영DB

    def trim_all_columns(self, df):
        """
        Trim whitespace from ends of each value across all series in dataframe
        """
        trimStrings = lambda x: x.strip() if type(x) is str else x
        return df.applymap(trimStrings)

    def df_to_sheet(self, df, ws, p_row, p_column):
        """
        DataFrame 을 엑셀 쉬트로 write
        """
        for k, row in enumerate(dataframe_to_rows(df, index=False, header=True)):
            # print(k, row)
            for i, item in enumerate(row):
                print("row = ", k, "col = ", i, "item = ", item)
                # ws.cell(row=p_row + k, column=p_column + i).value = item

        exit()

    def get_asset_data_input(self, p_dir_work, p_file_input_data, p_sheet_nm):
        """
            엑셀 파일의 쉬트를 읽어서 입력 데이터를 반환한다.
            p_sheet_nm: Input (자산군별 수익률), Condition (집계단위, 조건값 등)
        """

        file_asset_data = os.path.join(p_dir_work, p_file_input_data)
        file_df = pd.read_excel(file_asset_data, sheet_name=p_sheet_nm, header=0)  # read 엑셀파일 -> dataframe
        # print(file_df)

        return file_df

    def calcPortfolioPerf(self, weights, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate=0):
        """
        입력된 자산군 비중에 대해서 포트폴리오의 기대수익률, 변동성(표준편차), 사프 지수를 계산한다.

        INPUT
        weights: 포트폴리오 내의 자산군의 비중 array
        meanReturns: 포트폴리오 자산군의 평균 수익률
        covMatrix: 포트폴리오 자산군의 공분산
        numPeriodsAnnually: 연율화 단위
        riskFreeRate: 샤프지수 산출을 위한 무위험수익률

        OUTPUT
        tuple 포트폴리오 수익률(return), 변동성(volatility), 샤프(shapre)
        '''
        """

        # Calculate return and variance
        # portReturn = np.sum(meanReturns * weights) * numPeriodsAnnually
        portReturn = np.sum(meanReturns * weights)

        # portStdDev = np.sqrt(np.dot(weights.T, np.dot(covMatrix * numPeriodsAnnually, weights)))
        # portStdDev = np.sqrt(np.dot(weights.T, np.dot(covMatrix, weights))) * np.sqrt(numPeriodsAnnually)
        portStdDev = np.sqrt(np.dot(weights.T, np.dot(covMatrix, weights)))

        portSharpe = (portReturn - riskFreeRate) / portStdDev

        return portReturn, portStdDev, portSharpe

    def negSharpeRatio(self, weights, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate=0):
        '''
        입력된 자산군 비중에 대해서 샤프지수를 산출하여 음(-)으로 변환하여 반환한다.

        INPUT
        weights: 포트폴리오 내의 자산군의 비중 array
        meanReturns: 포트폴리오 자산군의 평균 수익률
        covMatrix: 포트폴리오 자산군의 공분산
        riskFreeRate: 샤프지수 산출을 위한 무위험수익률
        numPeriodsAnnually: 연율화 단위

        '''
        p_ret, p_var, psharpe = self.calcPortfolioPerf(weights, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate)

        return psharpe * -1

    def findEfficientMaxSharpe(self, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate, targetReturn, p_bounds):
        """
        목표수익률 수준에 대해 최대 샤프지수를 구성하는 포트폴리오 자산군의 비중을 계산한다.

        INPUT
        meanReturns: 포트폴리오 자산군의 평균 수익률
        covMatrix: 포트폴리오 자산군의 공분산
        riskFreeRate: 샤프지수 산출을 위한 무위험수익률
        targetReturn: 목표수익률(연율화)
        p_bounds: 비중 제약조건

        OUTPUT
        Dictionary of results from optimization
        """

        numAssets = len(meanReturns)

        def getPortfolioReturn(weights):
            return self.calcPortfolioPerf(weights, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate)[0]

        def getPortfolioSharpe(weights):
            return self.negSharpeRatio(weights, meanReturns, covMatrix, numPeriodsAnnually, riskFreeRate)

        constraints = ({'type': 'eq', 'fun': lambda x: getPortfolioReturn(x) - targetReturn},
                       {'type': 'eq', 'fun': lambda x: np.sum(x) - 1})
        # bounds = tuple((0, 1) for x in range(numAssets))
        bounds = p_bounds

        return sco.minimize(getPortfolioSharpe, numAssets * [1. / numAssets, ], method='SLSQP', bounds=bounds, constraints=constraints)

    def get_input_cond(self, df_cond, grp1, grp2="", ):
        """
        입력된 인수의 값을 비교하여 입력 DataFrame의 데이터를 필터링 하여 리턴 한다.
        """

        df_result = pd.DataFrame()
        if len(grp1) > 0 and len(grp2) == 0:
            df_result = df_cond[df_cond["group1"] == grp1]

        elif len(grp1) > 0 and len(grp2) > 0:
            df_result = df_cond[(df_cond["group1"] == grp1) & (df_cond["group2"] == grp2)]

        return df_result

    def make_mvo_optmz(self, p_df_cond, p_df_asset):
        """
        - 평균-분산 최적화(MVO)
          목표수익률 수준에서 최대 샤프비율을 가지는 자산의 최적 비중 산출
        """

        dic_result = {}  # 결과저장 dictionary

        # index 설정
        df_asset = p_df_asset.set_index("기간").copy()

        # -- 자산명칭
        dic_result["자산명칭"] = df_asset.columns.tolist()
        # print(dic_result["자산명칭"])

        # -- 자산_개수
        dic_result["자산_개수"] = len(df_asset.columns)
        # print(dic_result["자산_개수"])

        # -- 환산단위
        cond_rslt = self.get_input_cond(p_df_cond, grp1="설정값", grp2="환산단위")
        dic_result["환산단위"] = cond_rslt.loc[cond_rslt.index[0], "v1"]
        # print(dic_result["환산단위"])

        # -- 무위험수익률
        cond_rslt = self.get_input_cond(p_df_cond, grp1="설정값", grp2="무위험수익률")
        dic_result["무위험수익률"] = cond_rslt.loc[cond_rslt.index[0], "v1"]
        # print(dic_result["무위험수익률"])

        # -- 자산_기대수익률(수익률 평균)
        # meanReturns = df_asset.mean()
        # 자산군별 과거 수익률 평균을 자산군별 기대수익률 값으로 변경 적용
        cond_rslt = self.get_input_cond(p_df_cond, grp1="기대수익률")
        dic_result["자산_기대수익률"] = pd.Series(cond_rslt["v1"].values.tolist(), index=dic_result["자산명칭"])
        # print(dic_result["자산_기대수익률"])

        # -- 자산 공분산 행렬
        dic_result["자산_공분산_행렬"] = df_asset.cov() * dic_result["환산단위"]
        # print(dic_result["자산_공분산_행렬"])

        # -- 자산 상관계수 행렬
        dic_result["자산_상관계수_행렬"] = df_asset.corr(method='pearson')
        # print(dic_result["자산_상관계수_행렬"])

        # -- 목표수익률 구간 ---------------------------------------------------------------------------------
        cond_rslt = self.get_input_cond(p_df_cond, grp1="설정값", grp2="목표수익률구간")

        dic_result["목표수익률_F"] = cond_rslt.loc[cond_rslt.index[0], "v1"]
        dic_result["목표수익률_T"] = cond_rslt.loc[cond_rslt.index[0], "v2"]
        # print(dic_result["목표수익률_F"], dic_result["목표수익률_T"])

        dic_result["목표수익률구간"] = np.arange(dic_result["목표수익률_F"], dic_result["목표수익률_T"], step=0.00001)  # 1.67 ~ 3.99%
        # print(dic_result["목표수익률구간"])

        # -- 자산군 비중 제약조건 처리
        cond_rslt = self.get_input_cond(p_df_cond, grp1="비중제약")

        bounds_lst = []  # 리스트
        for i, row in cond_rslt.iterrows():

            if pd.isna(row["v1"]):
                bound_tmp = tuple((0, 1))
                bounds_lst.append(bound_tmp)

            else:
                bound_tmp = tuple((row["v1"], row["v1"]))
                bounds_lst.append(bound_tmp)
                # bounds_lst.append(tuple((row["수치값"], row["수치값"])))

        # List to Tuple
        bounds_tp = tuple(x for x in bounds_lst)

        dic_result["자산_비중제약"] = bounds_tp
        # print(dic_result["자산_비중제약"])

        # DataFrame 결과 저장 컬럼 정의
        df_col_lst = ['수익률', '위험', '샤프'] + dic_result["자산명칭"]

        ################################################################################################################
        # -- 목표수익률 수준에서 최대 샤프 포트폴리오 비중 계산
        maxSharpe_results = np.zeros((3 + dic_result["자산_개수"], len(dic_result["목표수익률구간"])))  # 최대샤프 포트폴리오 결과저장 매트릭스
        for i, tReturn in enumerate(dic_result["목표수익률구간"]):
            # print(i, tReturn)

            # 목표수익률 수준에서 최대 샤프지수를 구성하는 포트폴리오의 자산군 비중 계산
            res = self.findEfficientMaxSharpe(dic_result["자산_기대수익률"], dic_result["자산_공분산_행렬"], dic_result["환산단위"], dic_result["무위험수익률"], tReturn, dic_result["자산_비중제약"])

            maxSharpeWeights = res['x'].round(4)  # Max 샤프 구성 포트폴리오 비중
            print(maxSharpeWeights)

            # 포트폴리오 기대수익률, 위험(표준편차), 샤프 지수 계산
            pret, pvar, psharpe = self.calcPortfolioPerf(maxSharpeWeights, dic_result["자산_기대수익률"], dic_result["자산_공분산_행렬"], dic_result["환산단위"], dic_result["무위험수익률"])
            print("최적화 - 목표수익률 - 수익률 - 위험 - 샤프: {} - {} - {} - {} - {} ".format(res['success'], tReturn, pret, pvar, psharpe))

            # maxSharpe_results
            maxSharpe_results[0, i] = pret  # 포트폴리오 수익률
            maxSharpe_results[1, i] = pvar  # 포트폴리오 위험(표준편차)
            maxSharpe_results[2, i] = psharpe  # 샤프지수

            # iterate through the weight vector and add data to results array
            for j in range(len(maxSharpeWeights)):
                maxSharpe_results[j + 3, i] = maxSharpeWeights[j]

        # convert results array to Pandas DataFrame
        df_maxSharpe = pd.DataFrame(maxSharpe_results.T, columns=df_col_lst)

        df_maxSharpe["전체비중"] = df_maxSharpe[dic_result["자산명칭"]].sum(axis=1)
        # df_maxSharpe = df_maxSharpe[df_maxSharpe["전체비중"] == 1.0]  # 전체비중이 100%인 항목만 필터링

        df_maxSharpe = df_maxSharpe.sort_values(['수익률'], ascending=[True])
        df_maxSharpe = df_maxSharpe.reset_index(drop=False)
        # print(df_maxSharpe)
        ################################################################################################################

        ################################################################################################################
        # # bulk insert
        # df_maxSharpe.to_sql('max_sharpe_blk', self.egn_fnfund_dev, if_exists='replace', index=False, chunksize=500)
        # exit()

        # sql_txt = """
        #     select a."index", a."수익률", a."위험", a."샤프"
        #          , a."현금성", a."유동성", a."확정금리", a."국내채권", a."국내주식", a."대체투자", a."전체비중"
        #     from max_sharpe_blk a
        # """
        # df_maxSharpe = pdsql.read_sql(sql_txt, self.egn_fnfund_dev)
        # print(df_maxSharpe)
        ################################################################################################################

        dic_result["최적화_비중"] = df_maxSharpe

        return dic_result  # dictionary

    def make_asset_stat(self, p_df_cond, p_df_asset):
        """
        자산군별 기대수익률, 공분산 행렬 산출
        """

        dic_result = {}  # 결과저장 dictionary

        # -- 자산군
        cond_rslt = self.get_input_cond(p_df_cond, grp1="자산군")
        grouped = cond_rslt.groupby(["group2"])

        for k, group in grouped:

            dic_asset_grp = {}  # 자산군별 결과저장 dictionary

            # print(k)
            # Applying
            df_10 = group[["v1"]].copy()
            # print(df_10)

            # 자산군의 하위자산 명칭
            dic_asset_grp["자산명칭"] = df_10["v1"].values.tolist()

            # -- 환산단위
            cond_rslt = self.get_input_cond(p_df_cond, grp1="설정값", grp2="환산단위")
            dic_asset_grp["환산단위"] = cond_rslt.loc[cond_rslt.index[0], "v1"]

            # -- 자산군별 기대수익률(수익률 평균)
            df_er = self.get_input_cond(p_df_cond, grp1="기대수익률")
            df_er = df_er[df_er['group2'].isin(dic_asset_grp["자산명칭"])]
            dic_asset_grp["자산_기대수익률"] = pd.Series(df_er["v1"].values.tolist(), index=dic_asset_grp["자산명칭"])

            # -- 자산 공분산 행렬
            dic_asset_grp["자산_공분산_행렬"] = p_df_asset[dic_asset_grp["자산명칭"]].cov() * dic_asset_grp["환산단위"]

            # -- 자산 상관계수 행렬
            dic_asset_grp["자산_상관계수_행렬"] = p_df_asset[dic_asset_grp["자산명칭"]].corr(method='pearson')

            # -- 허용위험한도
            df_cond = self.get_input_cond(p_df_cond, grp1="허용위험한도")
            dic_asset_grp["허용위험한도"] = df_cond[df_cond["group2"] == k][["v1", "v2"]]

            dic_result[k] = dic_asset_grp

        # # 결과 출력
        # for key1, val1 in dic_result.items():
        #     print("{} = {}".format(key1, val1))
        #     for key2, val2 in val1.items():
        #         print("- {} = {}".format(key2, val2))

        return dic_result

    def make_optmz_shortfall(self, p_dic_asset_stat, p_dic_maxsharpe):
        """
        최적자산 비중에 의한 자산군별 수익률, 위험, Shortfall 산출

        """

        # 결과저장 dictionary
        # 자산군 정보
        dic_result = copy.deepcopy(p_dic_asset_stat)

        # 최적 자산비중 dataframe index 설정
        df_maxsharpe_sf = p_dic_maxsharpe["최적화_비중"].copy()
        # print(df_maxsharpe_sf)

        list_asset_grp_nm = []
        for key, dic_val in dic_result.items():
            # print(key)
            # print(dic_val["자산명칭"])

            list_asset_grp_nm.append(key)  # 자산군

            list_asset_nm = ["index"]
            list_asset_nm = list_asset_nm + dic_val["자산명칭"]

            df_asset = df_maxsharpe_sf[list_asset_nm].copy()

            df_asset["비중합계"] = df_asset[dic_val["자산명칭"]].sum(axis=1)

            for asset_nm in dic_val["자산명칭"]:
                df_asset[asset_nm] = df_asset[asset_nm] / df_asset["비중합계"]

            # 자산군의 자산비중 재계산 후 합계
            df_asset["비중합계"] = df_asset[dic_val["자산명칭"]].sum(axis=1)

            # Grouping Step - Splitting -> Applying -> Combining
            # Splitting
            grouped = df_asset.groupby(["index"])

            df_asset_result = pd.DataFrame()  # 결과 저장 df
            for i, group in grouped:
                # print(k)
                # Applying
                # df_10 = group.sort_values(['기준일자'], ascending=[True])

                df_10 = group[list_asset_nm].copy()

                asset_weights = df_10[dic_val["자산명칭"]].values[0]
                # print(asset_weights)

                # 포트폴리오 수익률, 위험(표준편차) 계산
                asset_pret, asset_pvar = self.calcPortfolioPerf(asset_weights, dic_val["자산_기대수익률"], dic_val["자산_공분산_행렬"], 12)[:2]
                # print(asset_pret, asset_pvar)

                df_10["수익률"] = asset_pret
                df_10["위험"] = asset_pvar

                # df_cond = dic_val["허용위험한도"]
                # risk_limit = df_cond["v1"].iloc[0]
                risk_limit = 0  # 수익률 0
                df_10[key + "_SF"] = sct.norm.cdf(x=risk_limit, loc=asset_pret, scale=asset_pvar)

                df_asset_result = df_asset_result.append(df_10, sort=False)

            # 자산군 SF 결과 저장
            # dic_result[key + "_SF"] = df_asset_result
            dic_result[key][key + "_SF"] = df_asset_result

            # print(df_asset_result)

            df_maxsharpe_sf = pd.merge(df_maxsharpe_sf, df_asset_result[["index", key + "_SF"]], on=["index"], how="left",
                                    suffixes=('_left', '_right'))

        dic_result["자산군"] = list_asset_grp_nm
        dic_result["최적화_SF"] = df_maxsharpe_sf

        return dic_result

    def main(self):

        # 작업폴더 및 입력 데이터 파일 경로 설정
        dir_work = "D:/Dev/fng_dev/자산배분/Data"  # 작업 폴더
        file_input_data = "mvo_input_data.xls"  # 자산군 수익률 데이터

        # -- ============ 조건입력 데이터 로딩 Start. ==================================================================
        # 조건입력 데이터 - 조건값 데이터를 읽어온다.
        sheet_nm = "Condition"
        df_cond_input = self.get_asset_data_input(p_dir_work=dir_work, p_file_input_data=file_input_data, p_sheet_nm=sheet_nm)
        df_cond_input = self.trim_all_columns(df_cond_input)  # 모든 컬럼 trim
        # print(df_cond_input)
        # -- ============ 조건입력 데이터 로딩 End. ====================================================================

        # -- ============ 자산별 월간 수익률 데이터 로딩 Start. ========================================================
        # 입력 데이터 - 자산별 월간 수익률 데이터를 읽어온다.
        sheet_nm = "Input"
        df_asset_input = self.get_asset_data_input(p_dir_work=dir_work, p_file_input_data=file_input_data, p_sheet_nm=sheet_nm)
        df_asset_input = self.trim_all_columns(df_asset_input)  # 모든 컬럼 trim
        # print(df_asset_input)
        # -- ============ 자산별 수익률 데이터 로딩 End. ===============================================================

        # -- ============ 최적 자산비중 산출 Start. ====================================================================
        dic_maxsharpe = self.make_mvo_optmz(df_cond_input, df_asset_input)
        # print(dic_maxsharpe)
        # for key, val in dic_maxsharpe.items():
        #     print("{} = {}".format(key, val))
        # -- ============ 최적 자산비중 산출 End. ======================================================================

        # -- ============ 최적 자산비중에 의한 자산군 수익율, 위험 산출 Start. =========================================
        dic_asset_stat = self.make_asset_stat(df_cond_input, df_asset_input)
        # -- ============ 최적 자산비중에 의한 자산군 수익율, 위험 산출 End. ===========================================

        # -- ============ 최적 자산비중의 자산군 Shortfall 산출 Start. =================================================
        dic_maxsharpe_sf = self.make_optmz_shortfall(dic_asset_stat, dic_maxsharpe)
        # print(dic_maxsharpe_sf)
        # -- ============ 최적 자산비중의 자산군 Shortfall 산출 End. ===================================================

        print("----- Excel 파일 생성 시작 ")

        # 엑셀 파일 저장 위치 지정
        output_filename, file_extension = os.path.splitext(file_input_data)
        output_filename = output_filename + "_최적화_결과.xlsx"
        result_excel_file = os.path.join(dir_work, output_filename)

        excel_writer = pd.ExcelWriter(result_excel_file, engine='xlsxwriter')  # Excel

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = excel_writer.book

        # default cell format to size 10
        workbook.formats[0].set_font_size(10)

        # 자산_수익률 시트 #################################################################################
        df_asset_input.to_excel(excel_writer, index=False, sheet_name="자산_수익률", startrow=0, startcol=0)

        # 최적화_비중 시트 #################################################################################
        tmp = pd.DataFrame(dic_maxsharpe["자산_기대수익률"])
        tmp = tmp.rename(index=str, columns={0: "기대수익률"})
        tmp.to_excel(excel_writer, index=True, sheet_name="최적화_비중", startrow=7, startcol=0)

        # worksheet = workbook.add_worksheet('최적화_비중')  # 시트 추가
        worksheet = excel_writer.sheets['최적화_비중']
        worksheet.write('A7', '자산_기대수익률')
        worksheet.write('A2', '자산명칭')
        worksheet.write('A3', ", ".join(asset_nm for asset_nm in dic_maxsharpe["자산명칭"]))

        worksheet.write('G2', '자산_개수')
        worksheet.write('G3', dic_maxsharpe["자산_개수"])

        worksheet.write('I2', '환산단위')
        worksheet.write('I3', dic_maxsharpe["환산단위"])

        worksheet.write('K2', '무위험수익률')
        worksheet.write('K3', dic_maxsharpe["무위험수익률"])

        worksheet.write('D7', '자산_공분산_행렬')
        df_10 = dic_maxsharpe["자산_공분산_행렬"]
        df_10.to_excel(excel_writer, index=True, sheet_name="최적화_비중", startrow=7, startcol=3)

        worksheet.write('L7', '자산_상관계수_행렬')
        df_10 = dic_maxsharpe["자산_상관계수_행렬"]
        df_10.to_excel(excel_writer, index=True, sheet_name="최적화_비중", startrow=7, startcol=11)

        worksheet.write('A18', '최적_자산비중(최대 샤프비율)')
        df_10 = dic_maxsharpe["최적화_비중"]
        df_10.to_excel(excel_writer, index=False, sheet_name="최적화_비중", startrow=18, startcol=0)

        list_asset_grp_nm = dic_maxsharpe_sf["자산군"]
        dic_risk_limit = {}  # 자산군 허용위험한도
        for item in list_asset_grp_nm:
            # print(item)
            # print(dic_maxsharpe_sf[item])

            dic_risk_limit[item] = dic_maxsharpe_sf[item]["허용위험한도"]  # 자산군별 허용위험한도

            df_10 = dic_maxsharpe_sf[item][item + "_SF"]
            df_10.to_excel(excel_writer, index=False, sheet_name=item, startrow=18, startcol=0)

            worksheet = excel_writer.sheets[item]
            worksheet.write('A1', item)  # 자산군 명칭
            worksheet.write('A3', '자산명칭')
            worksheet.write('A4', ", ".join(asset_nm for asset_nm in dic_maxsharpe_sf[item]["자산명칭"]))
            worksheet.write('A6', '자산_기대수익률')

            tmp = pd.DataFrame(dic_maxsharpe_sf[item]["자산_기대수익률"])
            tmp = tmp.rename(index=str, columns={0: "기대수익률"})
            tmp.to_excel(excel_writer, index=True, sheet_name=item, startrow=7, startcol=0)

            worksheet.write('D7', '자산_공분산_행렬')
            df_10 = dic_maxsharpe_sf[item]["자산_공분산_행렬"]
            df_10.to_excel(excel_writer, index=True, sheet_name=item, startrow=7, startcol=3)

            worksheet.write('L7', '자산_상관계수_행렬')
            df_10 = dic_maxsharpe_sf[item]["자산_상관계수_행렬"]
            df_10.to_excel(excel_writer, index=True, sheet_name=item, startrow=7, startcol=11)

        # 최적화_비중_SF 시트 ###############################################################################
        df_10 = dic_maxsharpe_sf["최적화_SF"]
        df_10.to_excel(excel_writer, index=False, sheet_name="최적화_비중_SF", startrow=11, startcol=0)

        df_20 = df_10.copy()
        # print(df_10)

        # print(dic_risk_limit)
        # 허용위험한도를 만족하고, 비중의 합이 1
        for key, val in dic_risk_limit.items():
            print(key, val)
            df_20 = df_20[df_20["전체비중"] == 1]
            df_20 = df_20[np.round((df_20[key + "_SF"]/100), 2) <= float(val["v1"].iloc[0])]
            # print(df_10)

        df_30 = df_20.describe()
        print(df_30)

        df_30 = df_30.iloc[[3, 7]]  # min, max
        df_30.to_excel(excel_writer, index=True, sheet_name="최적화_비중_SF", startrow=1, startcol=0)

        for key, val in dic_maxsharpe_sf.items():
            print("{} = {}".format(key, val))

        excel_writer.save()  # excel 파일 저장
        print("----- Excel 파일 생성 종료")


if __name__ == '__main__':

    # ===== 처리시작 시간 ==============================================================================================
    start = time.time()
    # =====

    aaMvo = AssetAllocationMVO()
    aaMvo.main()

    # ===== 처리종료 시간 ==============================================================================================
    end = time.time()  # 처리종료 시각
    # =====

    print("처리시간(분) : {0}".format(int((end - start) / 60)))

