import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 设置中文字体
plt.rcParams["font.family"] = ["SimHei"]
plt.rcParams['figure.dpi'] = 300

def load_data(excel_path=None, use_sample_data=True):
    """加载数据并进行预处理，使用指定的初始化数据"""
    if use_sample_data:
        # 表单1数据（使用指定的初始化数据）
        form1_data = {
            '文物编号': ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', 
                       '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                       '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
                       '31', '32', '33', '34', '35', '36', '37', '38', '39', '40',
                       '41', '42', '43', '44', '45', '46', '47', '48', '49', '50',
                       '51', '52', '53', '54', '55', '56', '57', '58'],
            '纹饰': ['C', 'A', 'A', 'A', 'A', 'A', 'B', 'C', 'B', 'B', 
                   'C', 'B', 'C', 'C', 'C', 'C', 'C', 'A', 'A', 'A',
                   'A', 'B', 'A', 'C', 'C', 'C', 'B', 'A', 'A', 'A',
                   'C', 'C', 'C', 'C', 'C', 'C', 'C', 'C', 'C', 'C',
                   'C', 'A', 'C', 'A', 'A', 'A', 'A', 'A', 'A', 'A',
                   'C', 'C', 'A', 'C', 'C', 'C', 'C', 'C'],
            '类型': ['高钾', '铅钡', '高钾', '高钾', '高钾', '高钾', '高钾', '铅钡', '高钾', '高钾',
                   '铅钡', '高钾', '高钾', '高钾', '高钾', '高钾', '高钾', '高钾', '铅钡', '铅钡',
                   '高钾', '高钾', '铅钡', '铅钡', '铅钡', '铅钡', '高钾', '铅钡', '铅钡', '铅钡',
                   '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡',
                   '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡',
                   '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡', '铅钡'],
            '颜色': ['蓝绿', '浅蓝', '蓝绿', '蓝绿', '蓝绿', '蓝绿', '蓝绿', '紫', '蓝绿', '蓝绿',
                   '浅蓝', '蓝绿', '浅蓝', '深绿', '浅蓝', '浅蓝', '浅蓝', '深蓝', np.nan, '浅蓝',
                   '蓝绿', '蓝绿', '蓝绿', '紫', '浅蓝', '紫', '蓝绿', '浅蓝', '浅蓝', '深蓝',
                   '紫', '浅绿', '深绿', '深绿', '浅绿', '深绿', '深绿', '深绿', '深绿', np.nan,
                   '浅绿', '浅蓝', '浅蓝', '浅蓝', '浅蓝', '浅蓝', '浅蓝', np.nan, '黑', '黑',
                   '浅蓝', '浅蓝', '浅蓝', '浅蓝', '绿', '蓝绿', '蓝绿', np.nan],
            '表面风化': ['无风化', '风化', '无风化', '无风化', '无风化', '无风化', '风化', '风化', '风化', '风化',
                      '风化', '风化', '无风化', '无风化', '无风化', '无风化', '无风化', '无风化', '风化', '无风化',
                      '无风化', '风化', '风化', '无风化', '风化', '风化', '风化', '风化', '风化', '无风化',
                      '无风化', '无风化', '无风化', '风化', '无风化', '风化', '无风化', '风化', '风化', '风化',
                      '风化', '风化', '风化', '风化', '无风化', '无风化', '无风化', '风化', '风化', '风化',
                      '风化', '风化', '风化', '风化', '无风化', '风化', '风化', '风化']
        }
        df1 = pd.DataFrame(form1_data)
        df1['文物编号'] = df1['文物编号'].astype(str)
        
        # 表单2数据（使用指定的初始化数据）
        form2_data = {
            '文物采样点': ['01', '02', '03部位1', '03部位2', '04', '05', '06部位1', '06部位2', '07', 
                        '08', '08严重风化点', '09', '10', '11', '12', '13', '14', '15', '16', '17', 
                        '18', '19', '20', '21', '22', '23未风化点', '24', '25未风化点', '26', '26严重风化点', 
                        '27', '28未风化点', '29未风化点', '30部位1', '30部位2', '31', '32', '33', '34', 
                        '35', '36', '37', '38', '39', '40', '41', '42未风化点1', '42未风化点2', '43部位1', 
                        '43部位2', '44未风化点', '45', '46', '47', '48', '49', '49未风化点', '50', '50未风化点', 
                        '51部位1', '51部位2', '52', '53未风化点', '54', '54严重风化点', '55', '56', '57', '58'],
            '二氧化硅(SiO2)': [69.33, 36.28, 87.05, 61.71, 65.88, 61.58, 67.65, 59.81, 92.63, 20.14, 4.61, 
                             95.02, 96.77, 33.59, 94.29, 59.01, 62.47, 61.87, 65.18, 60.71, 79.46, 29.64, 
                             37.36, 76.68, 92.35, 53.79, 31.94, 50.61, 19.79, 3.72, 92.72, 68.08, 63.3, 
                             34.34, 36.93, 65.91, 69.71, 75.51, 35.78, 65.91, 39.57, 60.12, 32.93, 26.25, 
                             16.71, 18.46, 51.26, 51.33, 12.41, 21.7, 60.74, 61.28, 55.21, 51.54, 53.33, 
                             28.79, 54.61, 17.98, 45.02, 24.61, 21.35, 25.74, 63.66, 22.28, 17.11, 49.01, 
                             29.15, 25.42, 30.39],
            '氧化钠(Na2O)': [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, 2.86, 3.38, 3.21, 2.1, 2.12, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, 7.92, np.nan, 2.31, np.nan, np.nan, np.nan, np.nan, 0.92, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 2.22, np.nan, 1.38, np.nan, 
                          np.nan, np.nan, 5.74, 5.68, np.nan, np.nan, 3.06, 2.66, np.nan, 4.66, 0.8, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, 1.22, 3.04, np.nan, np.nan, 2.71, np.nan, 
                          np.nan, np.nan],
            '氧化钾(K2O)': [9.99, 1.05, 5.19, 12.37, 9.67, 10.95, 7.37, 7.68, np.nan, np.nan, np.nan, 0.59, 
                         0.92, 0.21, 1.01, 12.53, 12.28, 7.44, 14.52, 5.71, 9.42, np.nan, 0.71, 4.71, 0.74, 
                         np.nan, np.nan, np.nan, np.nan, 0.4, np.nan, 0.26, 0.3, 1.41, np.nan, 1.6, 0.21, 0.15, 
                         0.25, 0.38, 0.14, 0.23, np.nan, np.nan, np.nan, 0.44, 0.15, 0.35, np.nan, np.nan, 0.2, 
                         0.11, 0.25, 0.29, 0.32, np.nan, 0.3, np.nan, np.nan, np.nan, np.nan, np.nan, 0.11, 
                         0.32, np.nan, np.nan, np.nan, np.nan, 0.34],
            '氧化钙(CaO)': [6.32, 2.34, 2.01, 5.87, 7.12, 7.35, np.nan, 5.41, 1.07, 1.48, 3.19, 0.62, 0.21, 
                         3.51, 0.72, 8.7, 8.23, np.nan, 8.27, np.nan, np.nan, 2.93, np.nan, 1.22, 1.66, 0.5, 
                         0.47, 0.63, 1.44, 3.01, 0.94, 1.34, 2.98, 4.49, 4.24, 0.89, 0.46, 0.64, 0.78, 0.38, 
                         0.37, 0.89, 0.68, 1.11, 1.87, 4.96, 0.79, np.nan, 5.24, 6.4, 2.14, 0.84, np.nan, 0.87, 
                         2.82, 4.58, 2.08, 3.19, 3.12, 3.58, 5.13, 2.27, 0.78, 3.19, np.nan, 1.13, 1.21, 1.31, 
                         3.49],
            '氧化镁(MgO)': [0.87, 1.18, np.nan, 1.11, 1.56, 1.77, 1.98, 1.73, np.nan, np.nan, np.nan, np.nan, 
                         np.nan, 0.71, np.nan, np.nan, 0.66, 1.02, 0.52, 0.85, 1.53, 0.59, 5.45, np.nan, 0.64, 
                         0.71, np.nan, np.nan, np.nan, np.nan, 0.54, 1, 1.49, 0.98, 0.51, 0.89, np.nan, 1, 
                         np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 2.73, 1.09, 1.16, 0.89, 0.95, 
                         np.nan, 0.74, 1.67, 0.61, 1.54, 1.47, 1.2, 0.47, 0.54, 1.19, 1.45, 0.55, 1.14, 1.28, 
                         1.11, np.nan, np.nan, np.nan, 0.79],
            '氧化铝(Al2O3)': [3.93, 5.73, 4.06, 5.5, 6.44, 7.5, 11.15, 10.05, 1.98, 1.34, 1.11, 1.32, 0.81, 
                           2.69, 1.46, 6.16, 9.23, 3.15, 6.18, np.nan, 3.05, 3.57, 1.51, 6.19, 3.5, 1.42, 1.59, 
                           1.9, 0.7, 1.18, 2.51, 4.7, 14.34, 4.35, 3.86, 3.11, 2.36, 2.35, 1.62, 1.44, 1.6, 
                           2.72, 2.57, 0.5, 0.45, 3.33, 3.53, 5.66, 2.25, 3.41, 12.69, 5, 4.79, 3.06, 13.65, 
                           5.38, 6.5, 1.87, 4.16, 5.25, 2.51, 1.16, 6.06, 4.15, 3.65, 1.45, 1.85, 2.18, 3.52],
            '氧化铁(Fe2O3)': [1.74, 1.86, np.nan, 2.16, 2.06, 2.62, 2.39, 6.04, 0.17, np.nan, np.nan, 0.32, 
                           0.26, np.nan, 0.29, 2.88, 0.5, 1.04, 0.42, 1.04, np.nan, 1.33, np.nan, 2.37, 0.35, 
                           np.nan, np.nan, 1.55, np.nan, np.nan, 0.2, 0.41, 0.81, 2.12, 2.74, 4.59, 1, np.nan, 
                           0.47, 0.17, 0.32, np.nan, 0.29, np.nan, 0.19, 1.79, np.nan, np.nan, 0.76, 1.39, 0.77, 
                           np.nan, np.nan, np.nan, 1.03, 2.74, 1.27, 0.33, np.nan, 1.19, 0.42, 0.23, np.nan, 
                           np.nan, np.nan, np.nan, np.nan, np.nan, 0.86],
            '氧化铜(CuO)': [3.87, 0.26, 0.78, 5.09, 2.18, 3.27, 2.51, 2.18, 3.24, 10.41, 3.14, 1.55, 0.84, 
                         4.93, 1.65, 4.73, 0.47, 1.29, 1.07, 1.09, np.nan, 3.51, 4.78, 3.28, 0.55, 2.99, 8.46, 
                         1.12, 10.57, 3.6, 1.54, 0.33, 0.74, np.nan, np.nan, 0.44, 0.11, 0.47, 1.51, 0.16, 0.68, 
                         3.01, 0.73, 0.88, np.nan, 0.19, 2.67, 2.72, 5.35, 1.51, 0.43, 0.53, 0.77, 0.65, np.nan, 
                         0.7, 0.45, 1.13, 0.7, 1.37, 0.75, 0.7, 0.54, 0.83, 1.34, 0.86, 0.79, 1.16, 3.13],
            '氧化铅(PbO)': [np.nan, 47.43, 0.25, 1.41, np.nan, np.nan, 0.2, 0.35, np.nan, 28.68, 32.45, np.nan, 
                         np.nan, 25.39, np.nan, np.nan, 1.62, 0.19, 0.11, 0.19, np.nan, 42.82, 9.3, 1, np.nan, 
                         16.98, 29.14, 31.9, 29.53, 29.92, np.nan, 17.14, 12.31, 39.22, 37.74, 16.55, 19.76, 16.16, 
                         46.55, 22.05, 41.61, 17.24, 49.31, 61.03, 70.21, 44.12, 21.88, 20.12, 59.85, 44.75, 13.61, 
                         15.99, 25.25, 25.4, 15.71, 34.18, 23.02, 44, 30.61, 40.24, 51.34, 47.42, 13.66, 55.46, 58.46, 
                         32.92, 41.25, 45.1, 39.35],
            '氧化钡(BaO)': [np.nan, np.nan, np.nan, 2.86, np.nan, np.nan, 1.38, 0.97, np.nan, 31.23, 30.62, np.nan, 
                         np.nan, 14.61, np.nan, np.nan, np.nan, 0, 0, 0, np.nan, 5.35, 23.55, 1.97, np.nan, 11.86, 
                         26.23, 6.65, 32.25, 35.45, np.nan, 4.04, 2.03, 10.29, 10.35, 3.42, 4.88, 3.55, 10, 5.68, 
                         10.83, 10.34, 9.79, 7.22, 6.69, 9.76, 10.47, 10.88, 7.29, 3.26, 5.22, 10.96, 10.06, 9.23, 
                         7.31, 6.1, 4.19, 14.2, 6.22, 8.94, np.nan, 8.64, 8.99, 7.04, np.nan, 7.95, 15.45, np.nan, 
                         7.66],
            '五氧化二磷(P2O5)': [1.17, 3.57, 0.66, 0.7, 0.79, 0.94, 4.18, 4.5, 0.61, 3.59, 7.56, 0.35, np.nan, 9.38, 
                              0.15, 1.27, 0.16, 0.26, np.nan, 0.18, 1.36, 8.83, 5.75, 1.1, 0.21, np.nan, 0.14, 0.19, 
                              3.13, 6.04, 0.36, 1.04, 0.41, np.nan, 1.41, 1.62, 0.17, 0.13, 0.34, 0.42, 0.07, 1.46, 
                              0.48, 1.16, 1.77, 7.46, 0.08, np.nan, np.nan, 12.83, np.nan, np.nan, 0.2, 0.1, 1.1, 
                              11.1, 4.32, 6.34, 6.34, 8.1, 8.75, 5.71, np.nan, 4.24, 14.13, 0.35, 2.54, np.nan, 8.99],
            '氧化锶(SrO)': [np.nan, 0.19, np.nan, 0.1, np.nan, 0.06, 0.11, 0.12, np.nan, 0.37, 0.53, np.nan, np.nan, 
                         0.37, np.nan, np.nan, np.nan, np.nan, 0.04, np.nan, 0.07, 0.19, np.nan, np.nan, np.nan, 0.33, 
                         0.91, 0.2, 0.45, 0.62, np.nan, 0.12, 0.25, 0.35, 0.48, 0.3, np.nan, np.nan, 0.22, np.nan, 0.22, 
                         0.31, 0.41, 0.61, 0.68, 0.47, 0.35, np.nan, 0.64, 0.47, 0.26, 0.23, 0.43, 0.85, 0.25, 0.46, 
                         0.3, 0.66, 0.23, 0.39, np.nan, 0.44, 0.27, 0.88, 1.12, np.nan, np.nan, np.nan, 0.24],
            '氧化锡(SnO2)': [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 0, 2.36, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 0.23, np.nan, 0.4, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 1.31, np.nan, np.nan, 
                          np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                          np.nan],
            '二氧化硫(SO2)': [0.39, np.nan, np.nan, np.nan, 0.36, 0.47, np.nan, np.nan, np.nan, 2.58, 15.03, np.nan, 
                           np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                           np.nan, np.nan, np.nan, np.nan, np.nan, 1.96, 15.95, np.nan, np.nan, np.nan, np.nan, 0.44, 
                           np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 3.66, np.nan, np.nan, np.nan, np.nan, 
                           np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                           np.nan, np.nan, 0.47, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 
                           np.nan]
        }
        df2 = pd.DataFrame(form2_data)
        
        # 解析采样点风化状态（根据补充说明）
        def parse_weathering_status(sample_point):
            if '未风化点' in sample_point:
                return '无风化'  # 采样点未风化（即使文物整体可能风化）
            elif '严重风化点' in sample_point:
                return '风化'    # 与表1一致
            elif '部位' in sample_point:
                # 从部位采样点提取文物编号，查询表1的风化状态
                base_id = sample_point.split('部位')[0]
                return df1[df1['文物编号'] == base_id]['表面风化'].values[0]
            else:
                # 基础采样点，与表1一致
                return df1[df1['文物编号'] == sample_point]['表面风化'].values[0]
        
        # 添加采样点风化状态列
        df2['采样点风化状态'] = df2['文物采样点'].apply(parse_weathering_status)
        # 提取文物编号（用于关联）
        df2['文物编号'] = df2['文物采样点'].str.extract('(\d+)').astype(str)
        
    else:
        # 从Excel读取数据
        df1 = pd.read_excel(excel_path, sheet_name='表单1', dtype={'文物编号': str})
        df2 = pd.read_excel(excel_path, sheet_name='表单2', dtype={'文物采样点': str})
        
        # 解析采样点风化状态
        def parse_weathering_status(sample_point):
            if '未风化点' in sample_point:
                return '无风化'
            elif '严重风化点' in sample_point:
                return '风化'
            elif '部位' in sample_point:
                base_id = sample_point.split('部位')[0]
                return df1[df1['文物编号'] == base_id]['表面风化'].values[0]
            else:
                return df1[df1['文物编号'] == sample_point]['表面风化'].values[0]
        
        df2['采样点风化状态'] = df2['文物采样点'].apply(parse_weathering_status)
        df2['文物编号'] = df2['文物采样点'].str.extract('(\d+)').astype(str)
    
    return df1, df2

def verify_initial_data(df1, df2):
    """验证初始化数据，重点检查样本数量和分布"""
    print("="*60)
    print("初始化数据验证报告")
    print("="*60)
    
    # 1. 表单1基本信息统计
    print("\n【表单1 样本数量统计】")
    print(f"总文物数量: {len(df1)}个")
    
    # 统计高钾和铅钡玻璃数量
    type_counts = df1['类型'].value_counts()
    high_k_count = type_counts.get('高钾', 0)
    lead_barium_count = type_counts.get('铅钡', 0)
    
    print("\n【类型数量验证】")
    print(f"高钾玻璃数量: {high_k_count}个")
    print(f"铅钡玻璃数量: {lead_barium_count}个")
    print(f"类型数量是否匹配预期: {'是' if (high_k_count == 18 and lead_barium_count == 40) else '否'}")
    
    # 2. 高钾玻璃详细列表
    print("\n【高钾玻璃文物编号列表】")
    high_k_ids = df1[df1['类型'] == '高钾']['文物编号'].tolist()
    print(f"共{len(high_k_ids)}个: {', '.join(high_k_ids)}")
    
    # 3. 铅钡玻璃详细列表
    print("\n【铅钡玻璃文物编号列表】")
    lead_barium_ids = df1[df1['类型'] == '铅钡']['文物编号'].tolist()
    print(f"共{len(lead_barium_ids)}个: {', '.join(lead_barium_ids)}")
    
    # 4. 表面风化状态分布
    print("\n【表面风化状态分布】")
    weathering_counts = df1['表面风化'].value_counts()
    for status, count in weathering_counts.items():
        print(f"  {status}: {count}个")
    
    # 5. 按类型的风化状态分布
    print("\n【各类型的表面风化分布】")
    type_weathering = pd.crosstab(df1['类型'], df1['表面风化'])
    print(type_weathering)
    
    # 6. 表单2采样点统计
    print("\n【表单2 采样点统计】")
    print(f"总采样点数量: {len(df2)}个")
    
    # 按文物类型统计采样点数量
    df_combined = pd.merge(df1[['文物编号', '类型']], df2, on='文物编号', how='left')
    sample_by_type = df_combined.groupby('类型').size()
    print("\n【按文物类型的采样点数量】")
    for typ, count in sample_by_type.items():
        print(f"  {typ}玻璃: {count}个采样点")
    
    # 7. 采样点类型分布
    def count_sample_types(sample_points):
        base = 0  # 基础采样点
        unweathered = 0  # 未风化点
        severe = 0  # 严重风化点
        part = 0  # 部位采样点
        
        for point in sample_points:
            if '未风化点' in point:
                unweathered += 1
            elif '严重风化点' in point:
                severe += 1
            elif '部位' in point:
                part += 1
            else:
                base += 1
        return {
            '基础采样点': base,
            '未风化点采样': unweathered,
            '严重风化点采样': severe,
            '部位采样点': part
        }
    
    sample_types = count_sample_types(df2['文物采样点'])
    print("\n【采样点类型分布】")
    for typ, count in sample_types.items():
        print(f"  {typ}: {count}个")
    
    # 可视化类型分布
    plt.figure(figsize=(10, 6))
    type_counts.plot(kind='bar', color=['#2ecc71', '#3498db'])
    plt.title('文物类型分布')
    plt.xlabel('类型')
    plt.ylabel('数量')
    plt.xticks(rotation=0)
    # 在柱状图上标注数量
    for i, v in enumerate(type_counts):
        plt.text(i, v + 0.5, str(v), ha='center')
    plt.tight_layout()
    plt.show()
    
    # 可视化各类型的风化状态
    plt.figure(figsize=(10, 6))
    type_weathering.plot(kind='bar', stacked=True, color=['#e74c3c', '#f1c40f'])
    plt.title('各类型文物的表面风化状态')
    plt.xlabel('类型')
    plt.ylabel('数量')
    plt.xticks(rotation=0)
    plt.legend(title='表面风化')
    plt.tight_layout()
    plt.show()
    
    return {
        '总数量': len(df1),
        '类型数量': type_counts,
        '风化状态数量': weathering_counts,
        '类型风化分布': type_weathering,
        '采样点数量': len(df2),
        '采样点类型分布': sample_types
    }

# 执行数据验证
if __name__ == "__main__":
    # 加载指定的初始化数据
    df1, df2 = load_data(use_sample_data=True)
    # 验证数据
    verification_results = verify_initial_data(df1, df2)
    
