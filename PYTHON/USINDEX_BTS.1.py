# %%
import 해외지수_ToExcel
import pandas as pd
import numpy as np
import xlwings as xw
import pandas.io.sql as psql
import os
from datetime import datetime
from tqdm import tqdm
import matplotlib.pyplot as plt
from IPython.display import display, HTML
import warnings
warnings.filterwarnings('ignore')

pd.set_option('display.max_columns', None)
pd.set_option("display.width", 300)
pd.set_option('display.max_rows', 100)



# %%
print(" < US BestSeller Index > ")
print("")

# -- ================================================================================================================ 
# 날짜 정의

# 정기변경일 : 옵션만기일 익영업일 (미국 영업일 기준)
# 종목선정일 : 정기변경일 5 영업일 전 (한국 영업일 기준)
# 비중확정일 : 정기변경일 3 영업일 전 (한국 영업일 기준)

print('1. 날짜 데이터 (종목선정일, 비중확정일, 정기변경일) 생성')
# ================================================================================================================ --


# 1. 정기변경일 (미국영업일 기준 옵션만기일 익영업일)
mssql_sheet = f'''
WITH A AS 
(
	SELECT ISO_CD
          ,DT
          ,DNO_OF_WK
          ,YYMM
	FROM FACTSET..[FZ_DATE]
	WHERE 1=1
	AND ISO_CD = 'US'
	AND DT BETWEEN '2016-01-01' AND '2024-12-31'
), B AS
(
	SELECT ISO_CD
          ,DT AS TRDT
          ,ROW_NUMBER () OVER (ORDER BY DT) AS TRDT_NUM
	FROM FACTSET..[FZ_DATE_RESV]
	WHERE 1=1
	AND ISO_CD = 'US'
	AND DT BETWEEN '2016-01-01' AND '2024-12-31'
)
SELECT A.ISO_CD
      ,A.DT
      ,A.DNO_OF_WK
      ,A.YYMM
      ,B.TRDT
      ,B.TRDT_NUM
      ,CASE WHEN B.TRDT IS NULL THEN 0 ELSE 1 END AS TRADE_YN
FROM A
	LEFT OUTER JOIN B
		ON A.DT = B.TRDT AND A.ISO_CD = B. ISO_CD
'''

# us_dates = psql.read_sql(mssql_sheet, conn_irisv2) 
us_dates = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/us_dates.xlsx')
us_dates['TRDT'] = us_dates['TRDT'].fillna(method='ffill') 

## 1-(1). 세번째 금요일 추출
dates_friday = us_dates[us_dates['DNO_OF_WK'] == 6].reset_index(drop=True)
dates_friday['WEEK_NUM'] = dates_friday.groupby('YYMM')['DT'].rank()
dates_third_friday = dates_friday[dates_friday['WEEK_NUM'] == 3].reset_index(drop=True)

## 1-(2). 미국 옵션만기일 (세번째 금요일이 휴장일이면 전영업일이 옵션만기일)
opt_exp = pd.DataFrame()
opt_exp['OPT_EXP'] = dates_third_friday.apply(lambda row: row['DT'] if row['TRADE_YN'] == 1 else row['TRDT'], axis=1) 
opt_exp_date = opt_exp['OPT_EXP']

## 1-(3). 정기변경일 (옵션만기일 익영업일)
merge_subset = us_dates[['DT', 'TRDT_NUM']]
merge_subset = merge_subset.rename(columns={'DT' : 'OPT_EXP'})

opt_exp = pd.merge(opt_exp, merge_subset, how='left', on='OPT_EXP')
opt_exp = opt_exp.rename(columns={'TRDT_NUM' : 'OPT_EXP_NUM'})
opt_exp['REB_DT_NUM'] = opt_exp['OPT_EXP_NUM'] + 1

reb_merge_subset = us_dates[['DT', 'TRDT_NUM']]
reb_merge_subset = reb_merge_subset.rename(columns={'TRDT_NUM' : 'REB_DT_NUM'})

reb_date = pd.merge(opt_exp, reb_merge_subset, how='left', on='REB_DT_NUM')
reb_date = reb_date.rename(columns={'DT' : 'REB_DT'})
reb_date = reb_date['REB_DT']


# 2. 종목선정일, 비중확정일
mssql_sheet = f''' 
WITH A AS 
(
	SELECT DT
          ,DNO_OF_WK
          ,YYMM
	FROM WISE..[TZ_DATE]
	WHERE 1=1
	AND DT >= '2016-01-01'
	AND YEAR <= '2024'
), B AS
(
	SELECT DT AS TRDT
          ,ROW_NUMBER () OVER (ORDER BY DT) AS TRDT_NUM
	FROM WISE..[TZ_DATE_RESV]
	WHERE 1=1
	AND DT >= '2016-01-01' 
)
SELECT A.DT
      ,A.DNO_OF_WK
      ,A.YYMM
      ,B.TRDT
      ,B.TRDT_NUM
      ,CASE WHEN B.TRDT IS NULL THEN 0 ELSE 1 END AS TRADE_YN
FROM A
	LEFT OUTER JOIN B
		ON A.DT = B.TRDT 
'''

# kr_dates = psql.read_sql(mssql_sheet, conn_irisv2)
kr_dates = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/kr_dates.xlsx')
kr_dates['TRDT'] = kr_dates['TRDT'].fillna(method='ffill') 

# 2-(1). 정기변경일 한국영업일 여부 확인
reb_date_subset = pd.DataFrame(reb_date).copy()
reb_date_subset = reb_date_subset.rename(columns={'REB_DT' : 'DT'})
reb_date_subset['REB_DT_CHECK'] = 'Y'

kr_reb_dt_check = pd.merge(kr_dates, reb_date_subset, how='left', on='DT')
kr_reb_dt_check = kr_reb_dt_check[kr_reb_dt_check['REB_DT_CHECK'] == 'Y']

# 2-(2). 정기변경일 한국영업일 맵핑 후 종목선정일, 비중확정일 산출
merge_subset = kr_dates[['DT', 'TRDT_NUM']]
merge_subset = merge_subset.rename(columns={'DT' : 'TRDT', 'TRDT_NUM': 'TRDT_NUM_NEW'})
kr_reb_dt_check = pd.merge(kr_reb_dt_check, merge_subset, left_on='TRDT', right_on='TRDT', how='left')

dates_eval_iif = kr_reb_dt_check[['TRDT', 'TRDT_NUM_NEW', 'TRADE_YN']]
dates_eval_iif['EVAL_DT_2_NUM'] = np.where(dates_eval_iif['TRADE_YN'] == 1, dates_eval_iif['TRDT_NUM_NEW'] - 7, dates_eval_iif['TRDT_NUM_NEW'] - 6)
dates_eval_iif['EVAL_DT_1_NUM'] = np.where(dates_eval_iif['TRADE_YN'] == 1, dates_eval_iif['TRDT_NUM_NEW'] - 6, dates_eval_iif['TRDT_NUM_NEW'] - 5)
dates_eval_iif['EVAL_DT_NUM'] = np.where(dates_eval_iif['TRADE_YN'] == 1, dates_eval_iif['TRDT_NUM_NEW'] - 5, dates_eval_iif['TRDT_NUM_NEW'] - 4)
dates_eval_iif['IIF_DT_NUM'] = np.where(dates_eval_iif['TRADE_YN'] == 1, dates_eval_iif['TRDT_NUM_NEW'] - 3, dates_eval_iif['TRDT_NUM_NEW'] - 2)

dates_subset_eval_2 = kr_dates[['DT', 'TRDT_NUM']]
dates_subset_eval_2 = dates_subset_eval_2.rename(columns={'DT' : 'EVAL_DT_2', 'TRDT_NUM' : 'EVAL_DT_2_NUM'})
dates_eval_iif = pd.merge(dates_eval_iif, dates_subset_eval_2, how='left', on='EVAL_DT_2_NUM')

dates_subset_eval_1 = kr_dates[['DT', 'TRDT_NUM']]
dates_subset_eval_1 = dates_subset_eval_1.rename(columns={'DT' : 'EVAL_DT_1', 'TRDT_NUM' : 'EVAL_DT_1_NUM'})
dates_eval_iif = pd.merge(dates_eval_iif, dates_subset_eval_1, how='left', on='EVAL_DT_1_NUM')

dates_subset_eval = kr_dates[['DT', 'TRDT_NUM']]
dates_subset_eval = dates_subset_eval.rename(columns={'DT' : 'EVAL_DT', 'TRDT_NUM' : 'EVAL_DT_NUM'})
dates_eval_iif = pd.merge(dates_eval_iif, dates_subset_eval, how='left', on='EVAL_DT_NUM')

dates_subset_iif = kr_dates[['DT', 'TRDT_NUM']]
dates_subset_iif = dates_subset_iif.rename(columns={'DT' : 'IIF_DT', 'TRDT_NUM' : 'IIF_DT_NUM'})
dates_eval_iif = pd.merge(dates_eval_iif, dates_subset_iif, how='left', on='IIF_DT_NUM')

eval_date_2 = dates_eval_iif['EVAL_DT_2']
eval_date_1 = dates_eval_iif['EVAL_DT_1']
eval_date = dates_eval_iif['EVAL_DT']
iif_date = dates_eval_iif['IIF_DT']

dates = pd.concat([eval_date_2, eval_date_1, eval_date, iif_date, opt_exp_date, reb_date], axis=1)



# -- ================================================================================================================ 
# 예탁결제원 데이터 가공 및 취합 정의

# 순매수결제금액 (3개월 TOP 50, 종목선정일 기준)
# 총결제금액 (3개월 TOP 50, 종목선정일 기준)
# 보관금액 (TOP 50, 데이터 입수 가능 시점 고려하여 종목선정일 2 영업일 전 기준)

print("2. 예탁결제원 데이터 (순매수결제금액, 총결제금액, 보관금액) 취합")
# ================================================================================================================ --


# 1. 순매수결제금액(3m) TOP50 
def seibro_netbuy(date):
    mssql_sheet = f'''
    SELECT COUNTRY
          ,START_DT
          ,END_DT
          ,RANK
          ,CODE AS ISIN
          ,NAME
          ,NET_BUY
    FROM factset.dbo.FW_STK_SETTLEMENT_3M 
    WHERE 1=1
    AND COUNTRY = 'US'
    AND [TYPE] = 4
    AND COUNTRY_CD = 1
    AND END_DT IN {date}
    ORDER BY END_DT ASC, NET_BUY DESC
    '''
    
    netbuy = psql.read_sql(mssql_sheet, conn_irisv2)
    return netbuy

eval_date_yyyymmdd = tuple(eval_date.dt.strftime("%Y%m%d"))
# netbuy = seibro_netbuy(eval_date_yyyymmdd)
netbuy = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/netbuy.xlsx')
netbuy['END_DT'] = netbuy['END_DT'].astype(int).astype(str)


# 2. 총결제금액(3m) TOP50 
def seibro_buysellsum(date):
    mssql_sheet = f'''
    SELECT COUNTRY
          ,START_DT
          ,END_DT
          ,RANK
          ,CODE AS ISIN
          ,NAME
          ,TOT AS BUY_SELL_SUM
    FROM factset.dbo.FW_STK_SETTLEMENT_3M 
    WHERE 1=1
    AND COUNTRY = 'US'
    AND [TYPE] = 3
    AND COUNTRY_CD = 1
    AND END_DT IN {date}
    ORDER BY END_DT ASC, TOT DESC
    '''

    buysellsum = psql.read_sql(mssql_sheet, conn_irisv2)
    return buysellsum

eval_date_yyyymmdd = tuple(eval_date.dt.strftime("%Y%m%d"))
# buysellsum = seibro_buysellsum(eval_date_yyyymmdd)
buysellsum = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/buysellsum.xlsx')
buysellsum['END_DT'] = buysellsum['END_DT'].astype(int).astype(str)


# 3. 보관금액 TOP50 
def seibro_deposit(date):
    mssql_sheet = f'''
    SELECT COUNTRY
          ,BASE_DT AS END_DT_ORIGINAL
          ,RANK
          ,CODE AS ISIN
          ,NAME
          ,DEPOSIT
    FROM factset.dbo.FW_STK_DEPOSIT
    WHERE 1=1
    AND COUNTRY = 'US'
    AND COUNTRY_CD = 1
    AND BASE_DT IN {date}
    ORDER BY BASE_DT ASC, DEPOSIT DESC
    '''

    deposit = psql.read_sql(mssql_sheet, conn_irisv2)
    
    dates_update = pd.DataFrame()
    dates_update['END_DT_ORIGINAL'] = dates['EVAL_DT_2'].dt.strftime("%Y%m%d")
    dates_update['END_DT'] = dates['EVAL_DT'].dt.strftime("%Y%m%d")
    deposit_update = pd.merge(deposit, dates_update, how='left', on='END_DT_ORIGINAL') 

    return deposit_update

eval_date_2_yyyymmdd = tuple(eval_date_2.dt.strftime("%Y%m%d"))
# deposit_update = seibro_deposit(eval_date_2_yyyymmdd)
deposit_update = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/deposit_update.xlsx')
deposit_update['END_DT'] = deposit_update['END_DT'].astype(int).astype(str)


# 4. 취합 (합집합)
df_seibro = pd.DataFrame()

for i in range(len(eval_date_yyyymmdd)):
    eval_date_subset = eval_date_yyyymmdd[i]
    
    netbuy_subset = netbuy[netbuy['END_DT'] == eval_date_subset]
    netbuy_subset_isin = netbuy_subset[['ISIN', 'NAME', 'END_DT']]
    
    buysellsum_subset = buysellsum[buysellsum['END_DT'] == eval_date_subset]
    buysellsum_subset_isin = buysellsum_subset[['ISIN', 'NAME', 'END_DT']]
    
    deposit_update_subset = deposit_update[deposit_update['END_DT'] == eval_date_subset]
    deposit_update_subset_isin = deposit_update_subset[['ISIN', 'NAME', 'END_DT']]
    
    # 순매수결제금액, 총결제금액, 보관금액 토탈 데이터셋 생성(ISIN 기준)
    subset_total = pd.concat([netbuy_subset_isin, buysellsum_subset_isin, deposit_update_subset_isin], axis=0)     
    subset_total = subset_total.drop_duplicates(subset='ISIN')
    subset_total = subset_total.reset_index(drop=True)

    # 토탈 데이터셋(ISIN)에 순매수결제금액 데이터 결합
    netbuy_subset = netbuy_subset[['ISIN', 'NET_BUY']]
    subset_total = pd.merge(subset_total, netbuy_subset, how='left', on='ISIN')
    
    # 토탈 데이터셋(ISIN)에 총결제금액 데이터 결합
    buysellsum_subset = buysellsum_subset[['ISIN', 'BUY_SELL_SUM']]
    subset_total = pd.merge(subset_total, buysellsum_subset, how='left', on='ISIN')

    # 토탈 데이터셋(ISIN)에 보관금액 데이터 결합
    deposit_update_subset = deposit_update_subset[['ISIN', 'DEPOSIT']]
    subset_total = pd.merge(subset_total, deposit_update_subset, how='left', on='ISIN')

    df_seibro = pd.concat([df_seibro, subset_total], axis=0)



# -- ================================================================================================================ 
# 예탁결제원 데이터 부가정보 취합 (code, market value, financial, sector) 후 별도처리

# 1. CODE(TICKER)
# 합병 혹은 상장폐지 이벤트 발생 시, FZ_ENTITY 테이블에서 해당 종목 정보 제거되기 때문에 FZ_ENTITY_DELETE 테이블 사용
# FZ_ENTITY_DELETE 는 내부코드 히스토리가 들어가 있는 테이블로써 여러 개의 하나의 TICKER에 여러 개의 내부코드 존재
# FS_STK_DATA와의 내부 조인을 통해 실제 사용되고 있는 내부코드를 포착, DISTINCT를 걸어서 중복값을 제거한 고유한 내부코드 추출

# 2. 시가총액 
# 종목선정일이 미국 휴장일인 경우 가장 최근 영업일 시가총액을 사용

# 3. 재무정보 (3년치 EPS)
# 재무정보 입수 가능 시점을 고려해서 LAG 적용
# 상장폐지 종목의 CODE가 신규 ETF에 부여된 경우에 한해 별도처리

# 4. 섹터정보 
# REFINITIV 섹터분류체계 사용
# 상장폐지 종목의 CODE가 신규 ETF에 부여된 경우에 한해 별도처리

# 5. 최종 데이터 별도 처리
# BRK 종목의 경우 class 주식을 .A 나 /A 로 표기하는 경우가 있어 별도 처리
# 최종 데이터셋 TICKER에서 -US 제거

print("3. 예탁결제원 데이터 부가정보(TICKER, 자산구분, 시가총액, 재무정보, 섹터정보) 추가")
# ================================================================================================================ --


# 1. 부가정보 데이터(국가, 자산종류, CIK, 내부코드) 호출
tuple_isin =  tuple(df_seibro['ISIN'].drop_duplicates())

mssql_sheet = f""" 
SELECT DISTINCT A.ISIN
			   ,B.TICKER
			   ,B.COUNTRY_ISO
			   ,B.SEC_TYPE
			   ,C.CMP_CIK
			   ,D.STK_CD
FROM factset.dbo.FZ_ISIN_TICKER_MAP A
	LEFT OUTER JOIN factset.dbo.FZ_ENTITY_DELETE B 
		ON A.TICKER = B.TICKER
	LEFT OUTER JOIN DISCLOSURE.dbo.FA_SEC_CIK C
		ON A.TICKER = CONCAT(C.TICKER, '-US')
	INNER JOIN factset.dbo.FS_STK_DATA D
		ON B.STK_CD = D.STK_CD
WHERE 1=1
AND A.ISIN IN {tuple_isin}
"""

# entity_info = psql.read_sql(mssql_sheet, conn_irisv2)
entity_info = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/entity_info.xlsx')
df_entity_subset = pd.merge(df_seibro, entity_info, how='left', on='ISIN')
df_entity_subset = df_entity_subset.dropna(subset=['TICKER']).reset_index(drop=True) # 상장폐지 혹은 TICKER 변경으로 인한 종목 제거


# 2-(1). 부가정보 데이터(시가총액) 호출
stk_cd = tuple(entity_info['STK_CD'].drop_duplicates())

mssql_sheet = f""" 
SELECT TRD_DT AS END_DT
	  ,STK_CD
      ,MKT_VAL
FROM factset.dbo.FS_STK_DATA
WHERE 1=1
AND STK_CD IN {stk_cd}
AND TRD_DT >= '2016-01-01'
"""

# entity_price = psql.read_sql(mssql_sheet, conn_irisv2)
entity_price = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/entity_price.xlsx')
entity_price['END_DT'] = entity_price['END_DT'].astype(int).astype(str)
df_entity = pd.merge(df_entity_subset, entity_price, how='left', on=['STK_CD', 'END_DT'])

# 2-(2). 시가총액 null값 처리 (미국휴장일)
mssql_sheet = f""" 
SELECT DT
      ,TRDT_1
      ,TRDT
      ,TRADE_YN
      ,YMD as YMD_DT
FROM FACTSET..[FZ_DATE]
WHERE 1=1
AND ISO_CD = 'US'
AND DT IN {tuple(dates['EVAL_DT'].astype(str))}
AND DT <= GETDATE()
"""

# date_subset = psql.read_sql(mssql_sheet, conn_irisv2)
date_subset = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/date_subset.xlsx')
date_subset['YMD_DT'] = date_subset['YMD_DT'].astype(str)

us_date = date_subset[date_subset['TRDT_1'].isna()]
us_date['YMD_TRDT'] = us_date['TRDT'].dt.strftime("%Y%m%d")

us_holiday = us_date[['YMD_DT']].reset_index(drop=True)
us_holiday = tuple(us_holiday['YMD_DT'])
# print("한국 기준 종목선정일이나 미국 기준 휴장일인 날짜")
# print(us_holiday)
# print("")

us_holiday_yester =  us_date[['YMD_TRDT']].reset_index(drop=True)
us_holiday_yester = tuple(us_holiday_yester['YMD_TRDT'])
# print("미국 휴장일 전 가장 최근 영업일")
# print(us_holiday_yester)
# print("")

for i in range(len(us_holiday)):
    us_hday = us_holiday[i]
    us_hday_yester = us_holiday_yester[i]
    
    df_update_subset = entity_price[entity_price['END_DT'] == us_hday_yester]
    df_update_subset['END_DT'] = us_hday
    df_update_subset = df_update_subset[['END_DT', 'STK_CD', 'MKT_VAL']]
    
    df_entity = df_entity.merge(df_update_subset, how='left', on=['END_DT', 'STK_CD'], suffixes=('', '_y'))
    df_entity['MKT_VAL'] = df_entity['MKT_VAL_y'].combine_first(df_entity['MKT_VAL'])
    df_entity.drop(columns=['MKT_VAL_y'], inplace=True)


# 3-(1). 재무정보 - 각 종목선정일 시점 구성종목 후보
EVAL_LIST={}

for i in df_entity['END_DT'][~df_entity['END_DT'].duplicated()]:
  EVAL_LIST[i]=[df_entity.iloc[x, 10] for x in range(len(df_entity)) if df_entity.iloc[x, 2]==i]
  EVAL_LIST[i]=[x for x in EVAL_LIST[i] if pd.isnull(x)==False]
  
# 3-(2). df_entity 데이터프레임에 재무데이터 취합하여 df_final 데이터프레임 생성
for eval_dt in EVAL_LIST.keys():
  
  eval_ym = pd.to_datetime(eval_dt, format='%Y%m%d').strftime('%Y%m')

  for stk_cd in EVAL_LIST[eval_dt]:
    
    mssql_sheet=f"""
    SELECT *
    FROM factset.dbo.FF_CMP_FINDATA
    WHERE 1=1
    AND CMP_CD = '{stk_cd}'
    AND ITEM_CD = '612271050' -- 당기순이익(지배)
    AND TERM_TYP = 1 -- 개별 누적 (연간)
    AND CONVERT(nvarchar(6), DATEADD(MONTH, 4, CONVERT(date, YYMM+'01')), 112) < '{eval_ym}' -- 보고서 제출기간 및 데이터 수집 기간 고려 (4개월 Lag로 지칭)
    ORDER BY YYMM DESC
    """

    # df_ENS = psql.read_sql(mssql_sheet, conn_irisv2)

    try:
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '1Y_결산년월'] = df_ENS.iloc[0,1]
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '1Y'] = df_ENS.iloc[0,5]
      
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '2Y_결산년월'] = df_ENS.iloc[1,1]
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '2Y'] = df_ENS.iloc[1,5]
      
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '3Y_결산년월'] = df_ENS.iloc[2,1]
      df_entity.loc[(df_entity.STK_CD==stk_cd) & (df_entity.END_DT==eval_dt), '3Y'] = df_ENS.iloc[2,5]
    
    except:
      continue

# 3-(3). 과거 상장폐지된 주식의 CODE가 신규 ETF CODE로 설정되어 잘못 병합되는 CASE 별도 처리
df_entity.loc[df_entity['SEC_TYPE'] == 'ETF_ETF', ['CMP_CIK', '1Y_결산년월', '1Y', '2Y_결산년월', '2Y', '3Y_결산년월', '3Y']] = None
df_financial = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/df_financial.xlsx')
df_financial['END_DT'] = df_financial['END_DT'].astype(int).astype(str)
df_financial['CMP_CIK'] = df_financial['CMP_CIK'].fillna(0).astype(int).astype(str)


# 4-(1). 섹터정보 호출
mssql_sheet = f"""
SELECT A.TICKER
	  ,A.STK_CD
	  ,A.SEC_TYPE 
	  ,LEFT(A.SECTOR_CODE, 2) AS SECTOR_CODE_BIG, D.SEC_NM AS ECONOMIC_SECTOR
	  ,LEFT(A.SECTOR_CODE, 4) AS SECTOR_CODE_MID, E.SEC_NM AS BUSINESS_SECTOR
	  ,LEFT(A.SECTOR_CODE, 6) AS SECTOR_CODE_SMALL, F.SEC_NM AS INDUSTRY_GROUP
	  ,LEFT(A.SECTOR_CODE, 8) AS SECTOR_CODE_DETAIL, G.SEC_NM AS INDUSTRY
	  ,A.SECTOR_CODE, C.SEC_NM AS ACTIVITY
FROM (
		SELECT DISTINCT A.TICKER ,A.STK_CD, A.SEC_TYPE, A.SECTOR_CODE 
		FROM factset.dbo.FZ_ENTITY_DELETE A
			INNER JOIN factset.dbo.FS_STK_DATA B
				ON A.STK_CD = B.STK_CD 
		WHERE 1=1
		AND A.COUNTRY_ISO = 'US'
     ) A
	LEFT OUTER JOIN FACTSET.dbo.FA_SECTOR_TRBC C	
		ON A.SECTOR_CODE = C.SEC_CD
	LEFT OUTER JOIN FACTSET.dbo.FA_SECTOR_TRBC D
	 	ON LEFT(A.SECTOR_CODE, 2) = D.SEC_CD
	LEFT OUTER JOIN FACTSET.dbo.FA_SECTOR_TRBC E
 		ON LEFT(A.SECTOR_CODE, 4) = E.SEC_CD
	LEFT OUTER JOIN FACTSET.dbo.FA_SECTOR_TRBC F
	 	ON LEFT(A.SECTOR_CODE, 6) = F.SEC_CD
	LEFT OUTER JOIN FACTSET.dbo.FA_SECTOR_TRBC G
	 	ON LEFT(A.SECTOR_CODE, 8) = G.SEC_CD
"""

# sector_info = psql.read_sql(mssql_sheet, conn_irisv2)
sector_info = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/sector_info.xlsx')
sector_info = sector_info[['TICKER', 'ECONOMIC_SECTOR', 'BUSINESS_SECTOR', 'INDUSTRY_GROUP', 'INDUSTRY', 'ACTIVITY']]

df_final = pd.merge(df_financial, sector_info, how='left', on='TICKER')

# 4-(2). 과거 상장폐지된 주식의 CODE가 신규 ETF CODE로 설정되어 잘못 병합되는 CASE 별도 처리
df_final.loc[df_final['SEC_TYPE'] == 'ETF_ETF', ['ECONOMIC_SECTOR', 'BUSINESS_SECTOR', 'INDUSTRY_GROUP', 'INDUSTRY', 'ACTIVITY']] = None


# 5-(1). BRK 종목 CIK 별도 처리
search_string = 'BRK.'
column_to_change = 'CMP_CIK'
new_value = '1067983'
df_final.loc[df_final['TICKER'].str.contains(search_string), column_to_change] = new_value

# 5-(2). 최종 데이터셋 TICKER에서 -US 제거
df_final['TICKER'] = df_final['TICKER'].str[:-3] 

# 5-(3). 결과물 출력 별도 처리 (SHARE, DR 만 출력)
all = df_final[df_final['END_DT'] == df_final['END_DT'].drop_duplicates().iloc[-1]]
except_etf = all[all['SEC_TYPE'] != 'ETF_ETF'].reset_index(drop=True)



# -- ================================================================================================================ 
# 지수 구성종목 선정

# 1. 기초 유니버스 셋업
# 2. 주식, DR 추출
# 3. 시가총액 50억 달러 이상 종목
# 4. 3년 연속 당기순이익 적자이면서 3년 연속 적자폭이 축소되지 않은 기업, 3년치 데이터가 없는 기업 제거
# 5. 순매수결제금액, 총결제금액, 보관금액 스코어링
# 6. multiple class 종목은 스코어가 높은 단일 종목만 편입
# 7. 암호화폐 관련 종목 제거
# 8. 상위 10 종목 추출
# 9. 유니버스 선정  

print("4. 지수구성종목 선정")
# ================================================================================================================ --


financial_data = pd.DataFrame()
raw_universe_df = pd.DataFrame()
universe_df = pd.DataFrame()

for i in range(len(eval_date_yyyymmdd)):

    evaluation_date = eval_date_yyyymmdd[i]
    today = datetime.today().strftime('%Y%m%d')
    
    # if evaluation_date >= today:
    if evaluation_date >= '20231231':
        break

    
    # 1. 기초 유니버스 셋업
    globals()[f"df_{evaluation_date}"] = df_final[df_final['END_DT'] == evaluation_date]
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)
    
        
    # 2. 주식, DR 추출
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][(globals()[f"df_{evaluation_date}"]['SEC_TYPE'] == 'SHARE') | (globals()[f"df_{evaluation_date}"]['SEC_TYPE'] == 'DR')].reset_index(drop=True)
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    
    # 3. 시가총액 스크리닝 (50억달러 초과 종목 추출)
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"]['MKT_VAL'] > 5000000000].reset_index(drop=True)
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    
    # 4. 재무건전성 스크리닝 (3년치의 충분한 당기순이익 데이터가 존재하지 않거나, 3년 연속 당기순이익 적자이면서 3년 연속 적자폭이 축소되지 않은 기업 제거)
    
    ## 4-(1). 3년 중 NaN값이 하나라도 있으면 NAN_CHECK 컬럼을 Y로 적재
    condition = (globals()[f"df_{evaluation_date}"]['1Y'].isnull()) | (globals()[f"df_{evaluation_date}"]['2Y'].isnull()) | (globals()[f"df_{evaluation_date}"]['3Y'].isnull())
    globals()[f"df_{evaluation_date}"].loc[condition, 'NAN_CHECK'] = 'Y'    
    
    ## 4-(2). 3년 연속 당기순이익 < 0 인 종목 3Y_DEFICIT 컬럼을 Y로 적재
    condition = (globals()[f"df_{evaluation_date}"]['1Y'] < 0) & (globals()[f"df_{evaluation_date}"]['2Y'] < 0) & (globals()[f"df_{evaluation_date}"]['3Y'] < 0)
    globals()[f"df_{evaluation_date}"].loc[condition, '3Y_DEFICIT'] = 'Y'
    
    ## 4-(3). 과거에서 최근으로 올수록 당기순이익 증가하면 3Y_INCREASE 컬럼을 Y로 적재
    condition = (globals()[f"df_{evaluation_date}"]['1Y'] > globals()[f"df_{evaluation_date}"]['2Y']) & (globals()[f"df_{evaluation_date}"]['2Y'] > globals()[f"df_{evaluation_date}"]['3Y']) 
    globals()[f"df_{evaluation_date}"].loc[condition, '3Y_INCREASE'] = 'Y'
    
    ## 4-(4). 제거 요건 만족 시에 ELIMINATE_REQ 컬럼을 Y로 적재
    condition = (globals()[f"df_{evaluation_date}"]['3Y_DEFICIT'] == 'Y') & (globals()[f"df_{evaluation_date}"]['3Y_INCREASE'] != 'Y') | (globals()[f"df_{evaluation_date}"]['NAN_CHECK'] == 'Y')
    globals()[f"df_{evaluation_date}"].loc[condition, 'ELIMINATE_REQ'] = 'Y'
    
    ## 4-(5). 재무 스크리닝 raw data 별도 저장
    financial_data = pd.concat([financial_data, globals()[f"df_{evaluation_date}"]], axis=0) 
    
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"]['ELIMINATE_REQ'] != 'Y'].reset_index(drop=True)
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    
    # 5. 순매수결제금액, 총결제금액, 보관금액 스코어링
    
    ## 5-(1). 순매수결제금액 스코어링
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].sort_values(by='NET_BUY', ascending=False)
    globals()[f"df_{evaluation_date}"]['NET_BUY'] = globals()[f"df_{evaluation_date}"]['NET_BUY'].fillna(0)
    globals()[f"df_{evaluation_date}"]['NET_BUY_RANK'] = globals()[f"df_{evaluation_date}"].NET_BUY.rank(method='min', ascending=False)
    
    ## 5-(2). 총결제금액 스코어링
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].sort_values(by='BUY_SELL_SUM', ascending=False)
    globals()[f"df_{evaluation_date}"]['BUY_SELL_SUM'] = globals()[f"df_{evaluation_date}"]['BUY_SELL_SUM'].fillna(0)
    globals()[f"df_{evaluation_date}"]['BUY_SELL_SUM_RANK'] = globals()[f"df_{evaluation_date}"].BUY_SELL_SUM.rank(method='min', ascending=False)
    
    ## 5-(3). 보관금액 스코어링
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].sort_values(by='DEPOSIT', ascending=False)
    globals()[f"df_{evaluation_date}"]['DEPOSIT'] = globals()[f"df_{evaluation_date}"]['DEPOSIT'].fillna(0)
    globals()[f"df_{evaluation_date}"]['DEPOSIT_RANK'] = globals()[f"df_{evaluation_date}"].DEPOSIT.rank(method='min', ascending=False)
    
    ## 5-(4). 합산순위 스코어링
    globals()[f"df_{evaluation_date}"]['RANK_AVERAGE'] = sum([globals()[f"df_{evaluation_date}"]['NET_BUY_RANK'], globals()[f"df_{evaluation_date}"]['BUY_SELL_SUM_RANK'], globals()[f"df_{evaluation_date}"]['DEPOSIT_RANK']]) / 3
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].sort_values(by=['RANK_AVERAGE', 'NET_BUY'], ascending=[True, False]) # 합산 순위 동차일시, 순매수결제금액 순으로 순위 부여
    globals()[f"df_{evaluation_date}"]['FINAL_RANK'] = globals()[f"df_{evaluation_date}"].RANK_AVERAGE.rank(method='first', ascending=True)
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].reset_index(drop=True)
    
    ## 5-(5). Seibro 스코어링 raw data 별도 저장
    globals()[f"df_raw_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['END_DT', 'ISIN', 'TICKER', 'NAME', 'SEC_TYPE', 'MKT_VAL', 'CMP_CIK', 'NET_BUY', 'NET_BUY_RANK', 'BUY_SELL_SUM', 'BUY_SELL_SUM_RANK', 'DEPOSIT', 'DEPOSIT_RANK', 'RANK_AVERAGE', 'FINAL_RANK']]
    raw_universe_df = pd.concat([raw_universe_df, globals()[f"df_raw_{evaluation_date}"]], axis=0)
    
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    
    # 6. multiple class 종목은 스코어링 상위 종목만 편입
    
    ## 6-(1). multiple class 종목 가운데 스코어링 하위 종목 ISIN 추출 후 데이터셋에서 제거 
    class_df = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"].duplicated(subset='CMP_CIK', keep=False)]
    multiple_class = class_df['CMP_CIK'].drop_duplicates().reset_index(drop=True) 

    for j in range(len(multiple_class)):
        
        cik = str(multiple_class[j])
        df_multiple_class = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"]['CMP_CIK'] == cik]
        multiple_class_isin = df_multiple_class['ISIN'].values[-1]
        
        globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][~globals()[f"df_{evaluation_date}"]['ISIN'].str.contains(multiple_class_isin)]
    
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    
    # 7. 특별 조항 (암호화폐 관련 종목 제거)
    
    ## 7-(1). 암호화폐 관련 종목 제거
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"]['TICKER'] != 'COIN'].reset_index(drop=True)
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']].reset_index(drop=True)

    ## 7-(2). REIT 종목 제거 
    eval_date_datetime = datetime.strptime(evaluation_date, '%Y%m%d')
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][globals()[f"df_{evaluation_date}"]['INDUSTRY_GROUP'] != 'Residential and commercial REITs']
    

    # 8. 종목 개수 10개로 제한
    globals()[f"df_{evaluation_date}"] = globals()[f"df_{evaluation_date}"].iloc[0:10]
    globals()[f"df_ticker_{evaluation_date}"] = globals()[f"df_{evaluation_date}"][['TICKER']]
    
    
    # 9. 유니버스 데이터프레임으로 취합
    globals()[f"univ_{evaluation_date}"] = globals()[f"df_ticker_{evaluation_date}"].rename(columns={'TICKER' : evaluation_date})
    globals()[f"univ_{evaluation_date}"] = globals()[f"univ_{evaluation_date}"][evaluation_date].reset_index(drop=True)
    
    universe_df = pd.concat([universe_df, globals()[f"univ_{evaluation_date}"]], axis=1)
    
    

# -- ================================================================================================================ 
# 구성 종목 비중 결정 및 재수채용데이터 생성

# 스코어 상위 5 종목 순차적으로 다음 비중 부여 (20%, 18%, 16%, 14%, 12%)
# 스코어 하위 5 종목 다음 고정 비중 부여 (4%)
    
print("5. 구성종목 비중 결정 및 지수채용데이터 생성")
# ================================================================================================================ --    


raw_data = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/data.xlsx')
raw_data['BASE_DT'] = raw_data['BASE_DT'].astype(str)

shr_update_data = pd.read_excel('C:/Users/tjgus/Desktop/US.BTS.1/DATA/shr_update.xlsx')
shr_update_data['TRD_DT'] = shr_update_data['TRD_DT'].astype(str)
shr_update_data['TICKER'] = shr_update_data['TICKER'].str[:-3] 

full_expectation = pd.DataFrame()   # 예상안
full_confirmation = pd.DataFrame()  # 확정안
total_data = pd.DataFrame()         # 실제 지수채용데이터

for i in tqdm(range(len(universe_df.columns))):
# for i in tqdm(range(1, 2)):  

    ticker = universe_df.iloc[:, i].dropna()
    ticker_tuple = tuple(ticker)
    
    if i < len(universe_df.columns)-1:

        eval_dt_1 = eval_date_1[i].strftime('%Y%m%d')   # 정기변경일 6영업일 전 (종목선정일이 미국휴장일인 경우 처리를 위해)
        eval_dt = eval_date[i].strftime('%Y%m%d')       # 정기변경일 5영업일 전
        iif_dt = iif_date[i].strftime('%Y%m%d')         # 정기변경일 3영업일 전
        reb_start_dt = reb_date[i].strftime('%Y%m%d')   # 정기변경일 
        reb_end_dt = reb_date[i+1].strftime('%Y%m%d')   # 다음 정기변경일 
    
    elif i < len(universe_df.columns): # 가장 최근 리밸런싱 별도처리

        eval_dt_1 = eval_date_1[i].strftime('%Y%m%d')   
        eval_dt = eval_date[i].strftime('%Y%m%d') 
        iif_dt = iif_date[i].strftime('%Y%m%d') 
        reb_start_dt = reb_date[i].strftime('%Y%m%d') 
        reb_end_dt = datetime.today().strftime('%Y%m%d') 
        
    else:
        pass

    
    # 1. 가격, 주식수, 시가총액 데이터프레임 호출
    mssql_sheet = f""" 
    SELECT A.BASE_DT
          ,A.TRDT
          ,A.TICKER
          ,B.SHORT_NAME
          ,STD_PRC
          ,CLS_PRC
          ,SHR_CNT
          ,NULL AS SHR_CNT_UPDATE
          ,STD_MKT_CAP
          ,CLS_MKT_CAP
    FROM FNINDEX.dbo.FI_IDX_TICKER_HIST A
        LEFT OUTER JOIN SVC.FACTSET.dbo.FZ_ENTITY B
            ON CONCAT(A.TICKER, '-US') = B.TICKER 
    WHERE 1=1
    AND A.TICKER IN {ticker_tuple}
    AND A.BASE_DT >= '{eval_dt_1}' 
    AND A.BASE_DT < '{reb_end_dt}'
    ORDER BY A.BASE_DT ASC
    """
    
    # data = psql.read_sql(mssql_sheet, conn_lilacv2) 
    data = raw_data[raw_data['TICKER'].isin(ticker_tuple)]
    data = data[(data['BASE_DT'] >= eval_dt_1) & (data['BASE_DT'] < reb_end_dt)]                       # 종목선정일 전영업일 ~ 다음 정기변경 전영업일
    
    eval_data = data[data['BASE_DT'] == eval_dt]                                                        # 예상안 작성용 (종목 fix)
    if len(eval_data) == 0:                                                                             # 종목선정일이 미국 휴장일인 경우 전영업일 데이터를 호출
        eval_data = data[data['BASE_DT'] == eval_dt_1] 
    iif_data = data[data['BASE_DT'] == iif_dt]                                                          # 확정안 작성용 (iif fix)
    idx_data = data[((data['BASE_DT'] >= reb_start_dt) & (data['BASE_DT'] <= reb_end_dt))]              # 정기변경일 ~ 다음 정기변경일 전영업일 (지수채용주식수 산출을 위해 산출)

    if (datetime.today().strftime('%Y%m%d') > iif_dt) & (len(data[data['BASE_DT'] == iif_dt]) < 10):    # 오류 체크
        print(ticker_tuple)
        print("! IIF 산출용 데이터셋 10 종목 미만")
        print(iif_dt)
        print(data[data['BASE_DT'] == iif_dt])
        print(i)


    # 2. 주식수 정기변경 (3, 6, 9, 12월) -> SHR_CNT_UPDATE
    
    ## 2-(1). 주식수 참조일 (전월 마지막 영업일) 주식수 호출 후 병합
    if (eval_dt[4:6] == '03' or eval_dt[4:6] == '06' or eval_dt[4:6] == '09' or eval_dt[4:6] == '12'):
        ticker_tuple_us = tuple(item + "-US" for item in ticker_tuple)


        mssql_sheet = f""" 
        SELECT CONVERT(NVARCHAR(8), B.TRD_DT, 112) TRD_DT
            ,A.SHORT_NAME
            ,A.TICKER
            ,B.LIST_STK_CNT
            ,B.CLOSE_PRC/ISNULL(C.AADJ,1.) AS ADJ_PRC
            ,B.CLOSE_PRC, C.ADJ, c.AADJ
            ,B. MOD_TM 
        FROM FACTSET.dbo.FZ_ENTITY A
            INNER JOIN FACTSET.dbo.FS_STK_DATA B
                ON A.STK_CD = B.STK_CD 
            LEFT OUTER JOIN FACTSET.dbo.FS_STK_ADJ_FACTOR C
                ON B.STK_CD = C.STK_CD AND B.TRD_DT BETWEEN C.START_DT AND C.TRD_DT 
        WHERE 1=1
        AND B.TRD_DT >= (SELECT TOP 1 DT
                        FROM FACTSET.DBO.FZ_DATE 
                        WHERE 1=1
                        AND ISO_CD = 'US'
                        AND DT < '{eval_dt}'
                        AND MN_END_YN = 1
                        ORDER BY DT DESC)
        AND B.TRD_DT <= '{iif_dt}'
        AND A.TICKER IN {ticker_tuple_us}
        ORDER BY TRD_DT ASC
        """
        
        # shr_update_data = psql.read_sql(mssql_sheet, conn_irisv2) # 주식수 참조일 ~ 비중확정일
        ref_dt = shr_update_data['TRD_DT'].drop_duplicates().reset_index(drop=True)[0]
        shr_update_data_for_merge = shr_update_data[shr_update_data['TRD_DT'] == ref_dt] # 주식수 참조일 기준 데이터
        
        eval_data = pd.merge(eval_data, shr_update_data_for_merge[['TICKER', 'TRD_DT', 'LIST_STK_CNT']], how='left', on='TICKER')
        eval_data['SHR_CNT_UPDATE'] = eval_data['LIST_STK_CNT']
        eval_data = eval_data.rename(columns={'TRD_DT' : 'REF_DT', 'LIST_STK_CNT' : 'REF_SHR_CNT'})
        
        iif_data = pd.merge(iif_data, shr_update_data_for_merge[['TICKER', 'TRD_DT', 'LIST_STK_CNT']], how='left', on='TICKER')
        iif_data['SHR_CNT_UPDATE'] = iif_data['LIST_STK_CNT']
        iif_data = iif_data.rename(columns={'TRD_DT' : 'REF_DT', 'LIST_STK_CNT' : 'REF_SHR_CNT'})
        
        
        ## 2-(2). 주식수 참조일과 종목선정일 사이에 액면분할과 같은 이벤트가 발생하면 수정계수가 적용된 날짜의 주식수를 호출
        adj_change_df = pd.DataFrame()
        for i in range(len(ticker_tuple)):
            
            subset = shr_update_data[shr_update_data['TICKER'] == ticker_tuple[i]]
            adj_change = subset[subset['ADJ'] != subset['ADJ'].shift(fill_value=subset['ADJ'].iloc[0])].reset_index(drop=True) # ADJ 비교 시에 shift로 인한 null값 처리
            
            if len(adj_change) > 0:
                adj_change_df = pd.concat([adj_change_df, adj_change])
                eval_data.loc[eval_data['TICKER'] == ticker_tuple[i], 'SHR_CNT_UPDATE'] = adj_change_df['LIST_STK_CNT'].iloc[-1]
                iif_data.loc[iif_data['TICKER'] == ticker_tuple[i], 'SHR_CNT_UPDATE'] = adj_change_df['LIST_STK_CNT'].iloc[-1]

        ## 2-(3). 업데이트된 주식수로 시가총액 산출
        eval_data['STD_MKT_CAP'] = eval_data['STD_PRC'] * eval_data['SHR_CNT_UPDATE'] # 업데이트된 주식수로 기준가시가총액 산출
        eval_data['CLS_MKT_CAP'] = eval_data['CLS_PRC'] * eval_data['SHR_CNT_UPDATE'] # 업데이트된 주식수로 종가시가총액 산출
        iif_data['STD_MKT_CAP'] = iif_data['STD_PRC'] * iif_data['SHR_CNT_UPDATE'] # 업데이트된 주식수로 기준가시가총액 산출
        iif_data['CLS_MKT_CAP'] = iif_data['CLS_PRC'] * iif_data['SHR_CNT_UPDATE'] # 업데이트된 주식수로 종가시가총액 산출


    # 3. 예상안 작성
    if eval_dt < today:
        eval_data['TRDT'] = reb_start_dt
        rank_data = raw_universe_df[raw_universe_df['END_DT'] == eval_dt]
        
        df_expectation = pd.merge(eval_data, rank_data[['TICKER', 'FINAL_RANK']], how='left', on='TICKER')
        df_expectation = df_expectation.sort_values(by=['FINAL_RANK'], ascending=True).reset_index(drop=True)
        df_expectation['WEIGHT_RANK'] = df_expectation.FINAL_RANK.rank(method='min', ascending=True)
        
        ## 3-(1). 합산 순위에 따른 비중 부여
        df_expectation_top1 = df_expectation[df_expectation['WEIGHT_RANK'] == 1] 
        df_expectation_top1['TARGET_WEIGHT'] = 0.2

        df_expectation_top2 = df_expectation[df_expectation['WEIGHT_RANK'] == 2] 
        df_expectation_top2['TARGET_WEIGHT'] = 0.18

        df_expectation_top3 = df_expectation[df_expectation['WEIGHT_RANK'] == 3] 
        df_expectation_top3['TARGET_WEIGHT'] = 0.16   

        df_expectation_top4 = df_expectation[df_expectation['WEIGHT_RANK'] == 4] 
        df_expectation_top4['TARGET_WEIGHT'] = 0.14

        df_expectation_top5 = df_expectation[df_expectation['WEIGHT_RANK'] == 5] 
        df_expectation_top5['TARGET_WEIGHT'] = 0.12

        df_expectation_remain = df_expectation.nlargest(5, 'WEIGHT_RANK')
        df_expectation_remain['TARGET_WEIGHT'] = 0.04

        df_expectation_final = pd.concat([df_expectation_top1, df_expectation_top2, df_expectation_top3, df_expectation_top4, df_expectation_top5, df_expectation_remain], axis=0)
        df_expectation_final = df_expectation_final.sort_values(by='WEIGHT_RANK', ascending=True).reset_index(drop=True)

        df_expectation_final['TOTAL_CLS_MKT_CAP'] = df_expectation_final['CLS_MKT_CAP'].sum()
        df_expectation_final['WEIGHT'] =  df_expectation_final['CLS_MKT_CAP'] / df_expectation_final['TOTAL_CLS_MKT_CAP'] 
        df_expectation_final['IIF'] = df_expectation_final['TARGET_WEIGHT'] / df_expectation_final['WEIGHT']
        
        ## 3-(2). iif 시계열 자료 생성을 위해 별도 저장
        full_expectation = pd.concat([full_expectation, df_expectation_final], axis=0) # 예상안


    # 4. 확정안 작성 및 IIF 확정
    if iif_dt < today:
        iif_data['TRDT'] = reb_start_dt
        rank_data = raw_universe_df[raw_universe_df['END_DT'] == eval_dt] # 종목선정일 DATA (SCORE DATA)
        
        df_confirmation = pd.merge(iif_data, rank_data[['TICKER', 'FINAL_RANK']], how='left', on='TICKER')
        df_confirmation = df_confirmation.sort_values(by=['FINAL_RANK'], ascending=True).reset_index(drop=True)
        df_confirmation['WEIGHT_RANK'] = df_confirmation.FINAL_RANK.rank(method='min', ascending=True)
        
        ## 4-(1). 합산 순위에 따른 비중 부여
        df_confirmation_top1 = df_confirmation[df_confirmation['WEIGHT_RANK'] == 1] 
        df_confirmation_top1['TARGET_WEIGHT'] = 0.2
        
        df_confirmation_top2 = df_confirmation[df_confirmation['WEIGHT_RANK'] == 2] 
        df_confirmation_top2['TARGET_WEIGHT'] = 0.18
        
        df_confirmation_top3 = df_confirmation[df_confirmation['WEIGHT_RANK'] == 3] 
        df_confirmation_top3['TARGET_WEIGHT'] = 0.16   
        
        df_confirmation_top4 = df_confirmation[df_confirmation['WEIGHT_RANK'] == 4] 
        df_confirmation_top4['TARGET_WEIGHT'] = 0.14
        
        df_confirmation_top5 = df_confirmation[df_confirmation['WEIGHT_RANK'] == 5] 
        df_confirmation_top5['TARGET_WEIGHT'] = 0.12
        
        df_confirmation_remain = df_confirmation.nlargest(5, 'WEIGHT_RANK')
        df_confirmation_remain['TARGET_WEIGHT'] = 0.04
        
        df_confirmation_fianl = pd.concat([df_confirmation_top1, df_confirmation_top2, df_confirmation_top3, df_confirmation_top4, df_confirmation_top5, df_confirmation_remain], axis=0)
        df_confirmation_fianl = df_confirmation_fianl.sort_values(by='WEIGHT_RANK', ascending=True).reset_index(drop=True)
        
        df_confirmation_fianl['TOTAL_CLS_MKT_CAP'] = df_confirmation_fianl['CLS_MKT_CAP'].sum()
        df_confirmation_fianl['WEIGHT'] =  df_confirmation_fianl['CLS_MKT_CAP'] / df_confirmation_fianl['TOTAL_CLS_MKT_CAP'] 
        df_confirmation_fianl['IIF'] = df_confirmation_fianl['TARGET_WEIGHT'] / df_confirmation_fianl['WEIGHT']
            
        ## 4-(2). iif 시계열 자료 생성을 위해 별도 저장
        full_confirmation = pd.concat([full_confirmation, df_confirmation_fianl], axis=0) # 확정안


    # 5. 시계열 데이터에 IIF 데이터 매핑
    entire_data = pd.merge(idx_data, df_confirmation_fianl[['TICKER', 'IIF']], how='left', on=['TICKER'])
    
    
    # 6. 전 구간 취합
    total_data = pd.concat([total_data, entire_data], axis=0)
    total_data['IDX_CLS_MKT_CAP'] = total_data['CLS_MKT_CAP'] * total_data['IIF']
    total_data['IDX_STD_MKT_CAP'] = total_data['STD_MKT_CAP'] * total_data['IIF']
    
    
# 7. 지수 구성 종목 비중 산출   
total_data = total_data.set_index('TRDT')
total_data['IDX_CLS_MKT_CAP_SUM'] = total_data.groupby('TRDT')['IDX_CLS_MKT_CAP'].sum()
total_data['WEIGHT'] = total_data['IDX_CLS_MKT_CAP'] / total_data['IDX_CLS_MKT_CAP_SUM']
total_data = total_data.reset_index()
total_data = total_data.sort_values(by=['TRDT', 'WEIGHT'], ascending=[True, False])



# -- ================================================================================================================ 
# 지수값 산출

# 계산 성능 향상을 위해 지수 제반 데이터 벡터화
# 기준시가총액(divisor)을 활용한 지수 계산 방식 채택
    
print("6. 지수값 계산")
# ================================================================================================================ --    


# 1. 지수 산출 데이터(종가, 기준가, 주식수, iif) 벡터화
df = total_data.sort_values(by=['TRDT', 'TICKER'], ascending=True)
df = df.reset_index(drop=True)

df_date = df['TRDT'].drop_duplicates()
df_date = df_date.reset_index(drop=True)

ticker = df['TICKER'].drop_duplicates()
ticker = ticker.sort_values(ascending=True)
ticker = ticker.reset_index(drop=True)

shares = pd.DataFrame()
shares['TRDT'] = df_date

close_prc = pd.DataFrame()
close_prc['TRDT'] = df_date

std_prc = pd.DataFrame()
std_prc['TRDT'] = df_date

iif = pd.DataFrame()
iif['TRDT'] = df_date

for i in range(len(ticker)):
    
    shares_subset = df[df['TICKER'] == ticker[i]]
    shares_subset = shares_subset[['TRDT', 'SHR_CNT']]
    shares_subset = shares_subset.reset_index(drop=True)
    shares = pd.merge(left = shares, right = shares_subset, how='left', on='TRDT')
    
    # 예외처리 (shares)
    shares = shares.set_index('TRDT')
    shares.columns = ['SHR_CNT'] * (i+1)
    shares = shares.reset_index()
    
    close_prc_subset = df[df['TICKER'] == ticker[i]]
    close_prc_subset = close_prc_subset[['TRDT', 'CLS_PRC']]
    close_prc_subset = close_prc_subset.reset_index(drop=True)
    close_prc = pd.merge(left = close_prc, right = close_prc_subset, how='left', on= 'TRDT')
    
    # 예외처리 (close_price)
    close_prc = close_prc.set_index('TRDT')
    close_prc.columns = ['CLS_PRC'] * (i+1)
    close_prc = close_prc.reset_index()   
    
    std_prc_subset = df[df['TICKER'] == ticker[i]]
    std_prc_subset = std_prc_subset[['TRDT', 'STD_PRC']]
    std_prc_subset = std_prc_subset.reset_index(drop=True)
    std_prc = pd.merge(left = std_prc, right = std_prc_subset, how='left', on= 'TRDT')
    
    # 예외처리 (standard_price)
    std_prc = std_prc.set_index('TRDT')
    std_prc.columns = ['STD_PRC'] * (i+1)
    std_prc = std_prc.reset_index()   
    
    iif_subset = df[df['TICKER'] == ticker[i]]
    iif_subset = iif_subset[['TRDT', 'IIF']]
    iif_subset = iif_subset.reset_index(drop=True)
    iif = pd.merge(left = iif, right = iif_subset, how='left', on= 'TRDT')
    
    # 예외처리 (iif)
    iif = iif.set_index('TRDT')
    iif.columns = ['IIF'] * (i+1)
    iif = iif.reset_index()   

shares = shares.set_index('TRDT')
shares.columns = ticker
# shares.to_excel("C:/Users/shd4323/Desktop/shares.xlsx")

close_prc = close_prc.set_index('TRDT')
close_prc.columns = ticker
# close_prc.to_excel("C:/Users/shd4323/Desktop/close_prc.xlsx")

std_prc = std_prc.set_index('TRDT')
std_prc.columns = ticker
# std_prc.to_excel("C:/Users/shd4323/Desktop/std_prc.xlsx")

iif = iif.set_index('TRDT')
iif.columns = ticker
# iif.to_excel("C:/Users/shd4323/Desktop/iif.xlsx")


# 2. divisor 사용 계산
close_mkt_val = (close_prc * shares * iif).sum(axis=1).sort_index()
std_mkt_val = (std_prc * shares * iif).sum(axis=1).sort_index()

divisor = pd.DataFrame(index = close_mkt_val.index)
divisor['divisor'] = ""

for i in range(len(close_mkt_val)):
# for i in range(0, 15):
    if i == 0:
        divisor.iloc[i] = close_mkt_val.iloc[i] / 1000
    else:
        divisor.iloc[i] = divisor.iloc[i-1] * std_mkt_val.iloc[i] / close_mkt_val.iloc[i-1]

index = close_mkt_val / divisor['divisor']
index = pd.DataFrame(index, columns=[f"FnGuide_US_BestSeller_INDEX"])
index.index.name = 'TRDT'

# %%