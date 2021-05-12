import base64
import streamlit as st
import pandas as pd
import openpyxl
import downloading

region = ['上海一区', '上海二区', '上海三区', '上海四区', '上海五区', '上海六区', '上海七区', '上海八区', '上海九区', '上海十区']

region2 = ['上海川沙绿地中心', '上海浦东成山路中心', '上海浦东川沙现代广场中心', '上海浦东花木中心', '上海浦东第一八佰伴中心',
           '上海浦东蓝村路中心', '上海浦东三林东中心', '上海浦东三林中心', '上海浦江城市广场中心']
job = ['SC', 'CR', 'TR']
job2 = ['SC', 'CR', 'TR', 'VIPTR', '其他']


# 读取不同组织区域在职数据
def report(excel_file, excel_file2, excel_file3, state_selected):
    df = excel_file
    df.loc[df['岗位'].str.contains('教师'), '岗位'] = 'TR'
    df.loc[df['岗位'].str.contains('升学规划师'), '岗位'] = 'SC'
    df.loc[df['岗位'].str.contains('班主任'), '岗位'] = 'CR'

    all = pd.DataFrame(columns=['TR', 'SC', 'CR'], index=[region])
    for each in region:
        for i in job:
            df1 = df.loc[(df['组织区域'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all.loc[each, i] = count

    # 写入组织区域编制数据
    fw = pd.DataFrame(index=['上海一区', '上海二区', '上海七区', '上海八区', '上海九区', '上海十区', 'A大区',
                             '上海三区', '上海四区', '上海五区', '上海六区', 'B大区', '上海'],
                      columns=['SC编制', 'SC在职', '缺编', 'SC满编率', 'CR编制', 'CR在职', '缺编',
                               'CR满编率', 'TR编制', 'TR在职', '缺编', 'TR满编率'])
    fw.iloc[0:6, 0] = [46, 64, 28, 17, 28, 25]  # SCA大区编制
    fw.iloc[7:11, 0] = [48, 44, 29, 42]  # SCB大区编制

    fw.iloc[0:6, 4] = [62, 79, 32, 13, 35, 28]  # CRA大区编制
    fw.iloc[7:11, 4] = [58, 51, 34, 55]  # CRB大区编制

    fw.iloc[0:6, 8] = [337, 433, 181, 82, 181, 160]  # TRA大区编制
    fw.iloc[7:11, 8] = [315, 293, 170, 277]  # TRB大区编制

    for each in region:
        for b in job:
            fw.at[each, b + '在职'] = all.at[each, b][0]

    fw.loc['A大区'] = fw.iloc[0:6:, :].sum()
    fw.loc['B大区'] = fw.iloc[7:11:, :].sum()
    fw.loc['上海'] = fw.loc[['A大区', 'B大区']].sum()

    fw.iloc[0:13, 2] = fw['SC编制'] - fw['SC在职']
    fw.iloc[0:13, 6] = fw['CR编制'] - fw['CR在职']
    fw.iloc[0:13, 10] = fw['TR编制'] - fw['TR在职']

    fw.iloc[0:13, 3] = fw['SC在职'] / fw['SC编制']
    fw.iloc[0:13, 7] = fw['CR在职'] / fw['CR编制']
    fw.iloc[0:13, 11] = fw['TR在职'] / fw['TR编制']

    fw.index.name = '区域'

    # 读取不同校区在职数据

    all2 = pd.DataFrame(columns=['TR', 'SC', 'CR'], index=[region2])
    for each in region2:
        for i in job:
            df1 = df.loc[(df['组织区域'] == '上海三区') & (df['部门\校区'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all2.loc[each, i] = count
    fw2 = pd.DataFrame(index=['上海川沙绿地中心', '上海浦东成山路中心', '上海浦东川沙现代广场中心', '上海浦东花木中心',
                              '上海浦东第一八佰伴中心', '上海浦东蓝村路中心', '上海浦东三林东中心', '上海浦东三林中心', '上海浦江城市广场中心'],
                       columns=['体量', 'SC编制', 'SC在职', 'SC缺编', 'SC满编率', 'CR编制', 'CR在职',
                                'CR缺编', 'CR满编率', 'TR编制', 'TR在职', 'TR缺编', 'TR满编率'])
    fw2.iloc[0:9, 0] = ['一般', '准KA', '微型', '一般', 'SSKA', '微型', '一般', '准KA', '一般']
    fw2.iloc[0:9, 1] = [4, 6, 4, 4, 12, 4, 4, 6, 4]  # SC编制
    fw2.iloc[0:9, 5] = [6, 9, 3, 5, 18, 3, 5, 9, 5]  # CR编制
    fw2.iloc[0:9, 9] = [31, 46, 14, 18, 85, 16, 26, 53, 26]  # TR编制

    for each in region2:
        for b in job:
            fw2.at[each, b + '在职'] = all2.at[each, b][0]

    fw2.loc['上海三区'] = fw2.iloc[:, :].sum()

    fw2['SC缺编'] = fw2['SC编制'] - fw2['SC在职']
    fw2['CR缺编'] = fw2['CR编制'] - fw2['CR在职']
    fw2['TR缺编'] = fw2['TR编制'] - fw2['TR在职']
    fw2['SC满编率'] = fw2['SC在职'] / fw2['SC编制']
    fw2['CR满编率'] = fw2['CR在职'] / fw2['CR编制']
    fw2['TR满编率'] = fw2['TR在职'] / fw2['TR编制']

    fw2.columns = ['体量', 'SC编制', 'SC在职', '缺编', 'SC满编率', 'CR编制', 'CR在职',
                   '缺编', 'CR满编率', 'TR编制', 'TR在职', '缺编', 'TR满编率']
    fw2.at['上海三区', '体量'] = ''
    fw2.index.name = '中心'

    # 读取不同组织区域入职数据
    df2 = excel_file2

    df2.loc[df2['岗位'].str.contains('VIP'), '岗位'] = 'VIPTR'
    df2.loc[df2['岗位'].str.contains('教师'), '岗位'] = 'TR'
    df2.loc[df2['岗位'].str.contains('升学规划师'), '岗位'] = 'SC'
    df2.loc[df2['岗位'].str.contains('班主任'), '岗位'] = 'CR'
    df2.loc[~df2['岗位'].str.contains('VIPTR|TR|SC|CR'), '岗位'] = '其他'

    all3 = pd.DataFrame(columns=['VIPTR', 'TR', 'SC', 'CR', '其他'], index=[region])
    for each in region:
        for i in job2:
            df1 = df2.loc[(df2['组织区域'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all3.loc[each, i] = count

    # 读取不同组织区域离职数据
    df3 = excel_file3

    df3.loc[df3['岗位'].str.contains('VIP'), '岗位'] = 'VIPTR'
    df3.loc[df3['岗位'].str.contains('教师'), '岗位'] = 'TR'
    df3.loc[df3['岗位'].str.contains('升学规划师'), '岗位'] = 'SC'
    df3.loc[df3['岗位'].str.contains('班主任'), '岗位'] = 'CR'
    df3.loc[~df3['岗位'].str.contains('VIPTR|TR|SC|CR'), '岗位'] = '其他'

    all4 = pd.DataFrame(columns=['VIPTR', 'TR', 'SC', 'CR', '其他'], index=[region])
    for each in region:
        for i in job2:
            df1 = df3.loc[(df3['组织区域'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all4.loc[each, i] = count

    # 读取不同校区入职数据
    all5 = pd.DataFrame(columns=['VIPTR', 'TR', 'SC', 'CR', '其他'], index=[region2])
    for each in region2:
        for i in job2:
            df1 = df2.loc[(df2['部门\校区'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all5.loc[each, i] = count

    # 读取不同校区离职数据
    all6 = pd.DataFrame(columns=['VIPTR', 'TR', 'SC', 'CR', '其他'], index=[region2])
    for each in region2:
        for i in job2:
            df1 = df3.loc[(df3['部门\校区'] == each)]
            count = df1['岗位'][df1['岗位'] == i].count()
            all6.loc[each, i] = count

    # 写入组织区域入职离职数据
    fw3 = pd.DataFrame(index=['上海一区', '上海二区', '上海七区', '上海八区', '上海九区', '上海十区', 'A大区',
                              '上海三区', '上海四区', '上海五区', '上海六区', 'B大区', '上海'],
                       columns=['SC入职', 'CR入职', 'TR入职', 'VIPTR入职', 'SDT入职', 'DDT入职', '其他入职', '小计入职',
                                'SC离职', 'CR离职', 'TR离职', 'VIPTR离职', 'SDT离职', 'DDT离职', '其他离职', '小计离职',
                                'SC', 'CR', 'TR', 'VIPTR', 'SDT', 'DDT', '其他', '小计'])
    for each in region:
        for b in job2:
            fw3.at[each, b + '入职'] = all3.at[each, b][0]

    for each in region:
        for b in job2:
            fw3.at[each, b + '离职'] = all4.at[each, b][0]

    fw3.loc['A大区'] = fw3.iloc[0:6:, :].sum()
    fw3.loc['B大区'] = fw3.iloc[7:11:, :].sum()
    fw3.loc['上海'] = fw3.loc[['A大区', 'B大区']].sum()
    fw3['小计入职'] = fw3['SC入职'] + fw3['CR入职'] + fw3['TR入职'] + fw3['VIPTR入职'] + fw3['其他入职']
    fw3['小计离职'] = fw3['SC离职'] + fw3['CR离职'] + fw3['TR离职'] + fw3['VIPTR离职'] + fw3['其他离职']
    fw3['SC'] = fw3['SC入职'] - fw3['SC离职']
    fw3['CR'] = fw3['CR入职'] - fw3['CR离职']
    fw3['TR'] = fw3['TR入职'] - fw3['TR离职']
    fw3['VIPTR'] = fw3['VIPTR入职'] - fw3['VIPTR离职']
    fw3['其他'] = fw3['其他'] - fw3['其他']
    fw3['小计'] = fw3['小计入职'] - fw3['小计离职']

    fw3.columns = ['SC', 'CR', 'TR', 'VIPTR', 'SDT', 'DDT', '其他', '小计',
                   'SC', 'CR', 'TR', 'VIPTR', 'SDT', 'DDT', '其他', '小计',
                   'SC', 'CR', 'TR', 'VIPTR', 'SDT', 'DDT', '其他', '小计']
    fw3.index.name = '大区/区域'

    # 写入校区入职离职数据
    fw4 = pd.DataFrame(index=['上海浦东第一八佰伴中心', '上海浦东成山路中心', '上海浦东三林中心', '上海川沙绿地中心',
                              '上海浦东花木中心', '上海浦江城市广场中心', '上海浦东蓝村路中心', '上海浦东三林东中心', '上海浦东川沙现代广场中心'],
                       columns=['SC入职', 'CR入职', 'TR入职', 'VIPTR入职', 'SDT入职', '其他入职', '小计入职',
                                'SC离职', 'CR离职', 'TR离职', 'VIPTR离职', 'SDT离职', '其他离职', '小计离职',
                                'SC', 'CR', 'TR', 'VIPTR', 'SDT', '其他', '小计'])
    for each in region2:
        for b in job2:
            fw4.at[each, b + '入职'] = all5.at[each, b][0]

    for each in region2:
        for b in job2:
            fw4.at[each, b + '离职'] = all6.at[each, b][0]

    fw4.loc['总计'] = fw4.iloc[:, :].sum()

    fw4['小计入职'] = fw4['SC入职'] + fw4['CR入职'] + fw4['TR入职'] + fw4['VIPTR入职'] + fw4['其他入职']
    fw4['小计离职'] = fw4['SC离职'] + fw4['CR离职'] + fw4['TR离职'] + fw4['VIPTR离职'] + fw4['其他离职']
    fw4['SC'] = fw4['SC入职'] - fw4['SC离职']
    fw4['CR'] = fw4['CR入职'] - fw4['CR离职']
    fw4['TR'] = fw4['TR入职'] - fw4['TR离职']
    fw4['VIPTR'] = fw4['VIPTR入职'] - fw4['VIPTR离职']
    fw4['其他'] = fw4['其他'] - fw4['其他']
    fw4['小计'] = fw4['小计入职'] - fw4['小计离职']

    fw4.columns = ['SC', 'CR', 'TR', 'VIPTR', 'SDT', '其他', '小计',
                   'SC', 'CR', 'TR', 'VIPTR', 'SDT', '其他', '小计',
                   'SC', 'CR', 'TR', 'VIPTR', 'SDT', '其他', '小计']
    fw4.index.name = '校区'

    if state_selected == '组织区域满编率':
        st.table(fw)
    elif state_selected == '校区满编率':
        st.table(fw2)
    elif state_selected == '组织区域入职离职数据':
        st.table(fw3)
    elif state_selected == '校区入职离职数据':
        st.table(fw4)

    st.sidebar.markdown(downloading.get_table_download_link(fw, '组织区域满编率', '组织区域满编率'), unsafe_allow_html=True)
    st.sidebar.markdown(downloading.get_table_download_link(fw2, '校区满编率', '校区满编率'), unsafe_allow_html=True)
    st.sidebar.markdown(downloading.get_table_download_link(fw3, '组织区域入职离职数据', '组织区域入职离职数据'), unsafe_allow_html=True)
    st.sidebar.markdown(downloading.get_table_download_link(fw4, '校区入职离职数据', '校区入职离职数据'), unsafe_allow_html=True)
