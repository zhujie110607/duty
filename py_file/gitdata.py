import numpy as np
import re

import pandas as pd
import os
import datetime

from pyecharts.charts import Page
from pyecharts.options import ComponentTitleOpts
import webbrowser
import common
from common import create_folder_if_not_exists
from pyecharts.components import Table


class GitData:
    def __init__(self):
        pass

    def get_data(self):
        self.dfs = pd.DataFrame()
        # 获取子库表数据
        sub_df = pd.read_excel(os.path.join(common.base_path, '子库表.xlsx'), sheet_name='子库表',
                               usecols=['子库编码', '子库类型'])
        if sub_df.shape[0] == 0:
            common.show_message('没有获取到子库表数据，请检查是否存在子库表.xlsx文件)', 0)
            return self.dfs
        # 获取数据源表
        source_df = pd.read_excel(common.path, sheet_name='MinvAnalysisTransaction',
                                  usecols=['物料编码', '交易数量', '子库编码', '货位编码', '对方子库编码',
                                           '对方货位编码', '交易日期', '交易类型', '交易来源单据号', '来源'])
        if source_df.shape[0] == 0:
            common.show_message('没有获取到数据源表数据', 0)
            return self.dfs
        # 把source_df['交易日期']格式化为yyyy-mm-dd
        source_df['交易日期'] = pd.to_datetime(source_df['交易日期']).dt.strftime('%Y-%m-%d')

        # 如果source_df['来源']为NaN,则替换为“无”
        source_df['来源'] = source_df['来源'].fillna('无')
        source_df['计费类型'] = '无'
        source_df['项数'] = 0
        # 1、用“子库编码”去子库表匹配子库类型。
        source_df = source_df.merge(sub_df, on='子库编码', how='left')
        # # 删除“字库类型”值为NaN的数据
        source_df = source_df.dropna(subset=['子库类型'])

        if source_df.shape[0] == 0:
            common.show_message('源表数据与子库表数据匹配后没有数据，请检查', 0)
            return self.dfs
        # 2.把“来源”列中包含“调整”或包含“返帐”或包含“调账”的数据删除
        rep = re.compile(r'调账|返帐|调整')
        source_df = source_df[~source_df['来源'].str.contains(rep)]
        if source_df.shape[0] == 0:
            common.show_message('删除 调账、返帐、调整 后没有数据，请检查', 0)
            return self.dfs

        # SHP数据
        # 3. 筛选“子库类型”列中=“SHP”并且“对方货位编码”不为空白值
        SHP_df = source_df[source_df['子库类型'].isin(['SHP']) & (source_df['对方货位编码'].notnull())]

        if SHP_df.shape[0] > 0:
            # 4、把“对方货位编码”列中删除开头为：HQ0*G、“HQ0G2”的数据，删除包含X、20500、20700、2500、2700的数据，删除结尾是BF、W500的数据
            # 定义过滤条件
            filter_pattern = re.compile(r'^HQ0\*G|^HQ0G2|X|20500|20700|2500|2700|BF$|W500$')
            # 应用过滤条件
            SHP_df = SHP_df[~SHP_df['对方货位编码'].str.contains(filter_pattern)]
            if SHP_df.shape[0] > 0:
                # 5.如果“交易数量”为正数，则计费类型为“出库”，如果是负数，则计费类型为“入库”
                SHP_df['计费类型'] = SHP_df['交易数量'].map(lambda x: '出库' if x > 0 else '入库')

        # WAT数据
        WAT_df = source_df[
            (source_df['子库类型'] == 'WAT') & (~source_df['货位编码'].str.contains('MJ')) & (
                source_df['对方货位编码'].notnull())]
        if WAT_df.shape[0] > 0:
            # 定义过滤条件
            filter_pattern = re.compile(r'20500|20700|2500|2700|BF$|W500$')
            # 应用过滤条件
            WAT_df = WAT_df[~WAT_df['对方货位编码'].str.contains(filter_pattern)]
            if WAT_df.shape[0] > 0:
                # “货位编码”包含“HJ”的数据，如果‘交易数量’为正数，则计费类型为“好转坏验收”，如果是负数，则计费类型为“好转坏入库”。
                # “货位编码”不包含“HJ”的数据，如果‘交易数量’为正数，则计费类型为“验收”，如果是负数，则计费类型为“入库”。
                # 定义条件
                conditions = [
                    (WAT_df['货位编码'].str.contains('HJ')) & (WAT_df['交易数量'] > 0),
                    (WAT_df['货位编码'].str.contains('HJ')) & (WAT_df['交易数量'] < 0),
                    (~WAT_df['货位编码'].str.contains('HJ')) & (WAT_df['交易数量'] > 0),
                    (~WAT_df['货位编码'].str.contains('HJ')) & (WAT_df['交易数量'] < 0)
                ]

                # 定义对应的值
                values = ['好转坏验收', '好转坏入库', '验收', '入库']

                # 使用 numpy.select 设置计费类型
                WAT_df['计费类型'] = np.select(conditions, values, default='未定义')

        # BF数据
        BF_df = source_df[(source_df['子库类型'] == '正常子库') &
                          (source_df['交易类型'] == 'Subinventory Transfer') &
                          (~source_df['交易来源单据号'].str.startswith('ST'))]
        if BF_df.shape[0] > 0:
            rep = re.compile(r'20500|20700|2500|2700|BF$|W500$')
            BF_df = BF_df[BF_df['货位编码'].str.contains(rep)]
            if BF_df.shape[0] > 0:
                BF_df['计费类型'] = BF_df['交易数量'].map(lambda x: '入库' if x > 0 else '出库')
        bf_df1 = source_df[(source_df['子库类型'] == '正常子库') & (
                (source_df['交易类型'] == 'SP_S20') | (source_df['交易类型'] == 'SP_OB005'))]
        bf_df1 = source_df[(source_df['子库类型'] == '正常子库') & (source_df['交易类型'].isin(['SP_S20', 'SP_OB005']))]

        if bf_df1.shape[0] > 0:
            conditions = [(bf_df1['交易数量'] < 0), (bf_df1['交易数量'] > 0)]
            values = ['报废出库', '无']
            bf_df1.loc[:, '计费类型'] = np.select(conditions, values, default='未定义')
        # 合并数据
        self.dfs = pd.concat([BF_df, bf_df1, SHP_df, WAT_df], ignore_index=True)
        return self.dfs

    def summarize(self):
        dfs = self.get_data()
        if dfs.shape[0] == 0:
            return None
        # 把“交易数量”识为绝对值
        dfs['交易数量'] = dfs['交易数量'].abs()
        rep = re.compile(r'入库|出库|好转坏验收|好转坏入库|报废出库')
        t1_df = dfs[dfs['计费类型'].str.contains(rep)]
        if t1_df.shape[0] > 0:
            t1_df.loc[:, :].fillna('', inplace=True)
            # 透视表
            t1 = t1_df.pivot_table(
                index=['物料编码', '子库编码', '货位编码', '对方子库编码', '对方货位编码', '交易日期', '计费类型',
                       '交易来源单据号', '项数'], values='交易数量', aggfunc='sum')
            # 把行索引转为列
            t1.reset_index(inplace=True)
            t1.loc[:, '项数'] = 1
            t1 = t1.pivot_table(index=['交易日期', '计费类型'], values=['项数', '交易数量'], aggfunc='sum')
            t1.reset_index(inplace=True)

        t2_df = dfs[dfs['计费类型'] == '验收']
        t2 = pd.DataFrame()
        if t2_df.shape[0] > 0:
            t2_df.loc[:, :].fillna('', inplace=True)
            t2_df.loc[:, '项数'] = 1
            t2 = t2_df.pivot_table(index=['交易日期', '计费类型'], values=['项数', '交易数量'], aggfunc='sum')
            t2.reset_index(inplace=True)
        ts = pd.concat([t1, t2], ignore_index=True)
        if ts.shape[0] > 0:
            result_folder = os.path.join(common.base_path, '结果')
            excel_folder = os.path.join(result_folder, 'Excel', datetime.date.today().strftime('%Y-%m-%d'))
            html_folder = os.path.join(result_folder, 'Html', datetime.date.today().strftime('%Y-%m-%d'))
            save_path = [excel_folder, html_folder]

            create_folder_if_not_exists(save_path)
            t = datetime.datetime.now().strftime("%H%M%S")
            writes = pd.ExcelWriter(os.path.join(save_path[0], f"结果{t}.xlsx"))
            ts.to_excel(writes, index=False)

        t3_df = dfs[dfs['计费类型'] != '无']
        t3 = t3_df.pivot_table(
            index=['物料编码', '子库编码', '货位编码', '对方子库编码', '对方货位编码', '交易日期', '计费类型',
                   '交易来源单据号',
                   '项数'], values='交易数量', aggfunc='sum')
        if t3_df.shape[0] > 0:
            t3.reset_index(inplace=True)
            t3['项数'] = 1
            t3 = t3.pivot_table(index='子库编码', values=['项数', '交易数量'], aggfunc='sum')
            t3.reset_index(inplace=True)
            t3.to_excel(writes, index=False, startrow=0, startcol=8)
            writes.close()

        # 写入html表格并打开
        page = Page(layout=Page.SimplePageLayout)  # 网页中各子图可拖动。默认的SimplePageLayout不可拖动
        table1 = Table()
        table2 = Table()
        if ts.shape[0] > 0:
            table1.add(ts.columns.tolist(), ts.values.tolist())
            table1.set_global_opts(title_opts=ComponentTitleOpts(title="保税坏件货量计算",
                                                                 subtitle=(datetime.datetime.now().strftime(
                                                                     '%Y-%m-%d %H:%M:%S'))))
        if t3.shape[0] > 0:
            table2.add(t3.columns.tolist(), t3.values.tolist())
            table2.set_global_opts(
                title_opts=ComponentTitleOpts(subtitle=(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))))
        page.add(table1, table2)
        page.render(os.path.join(save_path[1], f"结果{t}.html"))
        webbrowser.open(os.path.join(save_path[1], f"结果{t}.html"))
