import streamlit as st
import pandas as pd
import os
import time
from docx import Document
from io import BytesIO
import numpy as np

# ==========================================
# 工具函数：更智能的文件读取
# ==========================================
def smart_load_file(uploaded_file, header_keywords=None, header_idx_fallback=0, match_mode='any'):
    """
    智能读取上传的文件
    match_mode: 'any' (任一关键词匹配即为表头) | 'all' (所有关键词都在同一行才算表头)
    """
    if uploaded_file is None:
        return None
    
    try:
        filename = uploaded_file.name
        ext = os.path.splitext(filename)[1].lower()
        
        # 1. 预读取寻找表头
        uploaded_file.seek(0)
        if ext == '.csv':
            try:
                df_raw = pd.read_csv(uploaded_file, header=None, nrows=20, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, nrows=20, encoding='gbk')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, nrows=20)
            
        # 2. 定位表头索引
        header_idx = header_idx_fallback
        if header_keywords:
            for i, row in df_raw.iterrows():
                row_str = " ".join(row.astype(str).values)
                
                # 匹配逻辑
                is_match = False
                if match_mode == 'all':
                    # 必须包含所有关键词
                    if all(k in row_str for k in header_keywords):
                        is_match = True
                else:
                    # 包含任一关键词
                    if any(k in row_str for k in header_keywords):
                        is_match = True
                
                if is_match:
                    header_idx = i
                    break
        
        # 3. 正式读取
        uploaded_file.seek(0)
        if ext == '.csv':
            try:
                df = pd.read_csv(uploaded_file, header=header_idx, encoding='utf-8')
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_idx, encoding='gbk')
        else:
            df = pd.read_excel(uploaded_file, header=header_idx)
            
        # 清理列名 (去除前后空格、换行符)
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        return df

    except Exception as e:
        st.error(f"读取文件失败: {filename} - {str(e)}")
        return None

def safe_div(n, d):
    try: return n / d if d and d != 0 else 0
    except: return 0

# ==========================================
# 1. 全院数据处理
# ==========================================
class HospitalStatsProcessor:
    def __init__(self):
        self.output_rows = [
            "门诊次均药品费用（元）", "门诊次均药品费用(不含中药饮片)(元)",
            "门诊药占比（%）", "门诊药占比(不含中药饮片)（%）",
            "住院次均药品费用（元）", "住院次均药品费用(不含中药饮片)(元)",
            "住院药占比（%）", "住院药占比(不含中药饮片)（%）",
            "抗菌药物使用率(%)", "抗菌药物使用强度",
            "门诊基本药物占处方用药百分率(%)", "住院基本药物金额所占比例(不含中药饮片)(%)",
            "住院患者静脉输液使用率（%）", "住院患者人均静脉输液天数",
            "住院患者平均每床日使用静脉输液体积(ml)", "住院患者人均静脉输液药品品种数",
            "重点监控品种收入占比（%）"
        ]

    def extract_data(self, inf_file, qc_file):
        raw_data = {}
        # 输液表 (全院表通常比较标准)
        df_inf = smart_load_file(inf_file, header_keywords=['指标名称'], header_idx_fallback=3)
        if df_inf is not None:
            for _, row in df_inf.iterrows():
                name = str(row.get('指标名称', '')).strip()
                val = row.get('指标值', 0)
                if name == "住院患者静脉输液使用率①(100%)": raw_data["inf_rate"] = val
                elif name == "住院患者人均静脉输液天数": raw_data["inf_days"] = val
                elif name == "住院患者平均每床日使用静脉输液体积(ml)": raw_data["inf_vol"] = val
                elif name == "住院患者人均静脉输液药品品种数": raw_data["inf_types"] = val

        # 质控表
        df_qc = smart_load_file(qc_file, header_keywords=['指标名称'], header_idx_fallback=3)
        if df_qc is not None:
            is_outpatient = True
            out_patients = 0; in_patients = 0
            out_cost_no_herb = 0; in_cost_no_herb = 0

            for i in range(len(df_qc)):
                row = df_qc.iloc[i]
                name = str(row.get('指标名称', '')).strip()
                if not name or name == 'nan': continue
                val = row.get('指标值', 0); mol_val = row.get('分子值', 0)

                if "住院" in name or "抗菌药物" in name or "病人平均药品金额" in name: 
                    if "抗菌药物" in name or "病人平均药品金额" in name: is_outpatient = False

                if is_outpatient:
                    if name == "平均药品金额(元)":
                        raw_data["op_avg_cost"] = val
                        if i+1 < len(df_qc): out_patients = df_qc.iloc[i+1].get('分子值', 0)
                    elif name == "药占比(%)": raw_data["op_drug_ratio"] = val
                    elif name == "药占比(不含中药饮片)(%)":
                        raw_data["op_drug_ratio_no_herb"] = val
                        out_cost_no_herb = mol_val
                    elif name == "国家基本药物占处方用药百分率(%)": raw_data["op_basic_drug"] = val
                else:
                    if name == "病人平均药品金额(元)":
                        raw_data["ip_avg_cost"] = val
                        if i+1 < len(df_qc): in_patients = df_qc.iloc[i+1].get('分子值', 0)
                    elif name == "药占比(%)": raw_data["ip_drug_ratio"] = val
                    elif name == "药占比(不含中药饮片)(%)":
                        raw_data["ip_drug_ratio_no_herb"] = val
                        in_cost_no_herb = mol_val
                    elif name == "基本药物金额所占比例(不含中药饮片)(%)": raw_data["ip_basic_drug_amount"] = val
                    elif name == "抗菌药物使用率(%)": raw_data["antibiotic_rate"] = val
                    elif "抗菌药物使用强度" in name: raw_data["antibiotic_intensity"] = val
                    elif "重点监控品种" in name: raw_data["key_monitor"] = val

            raw_data["op_avg_cost_no_herb"] = safe_div(out_cost_no_herb, out_patients)
            raw_data["ip_avg_cost_no_herb"] = safe_div(in_cost_no_herb, in_patients)

        final_data = {}
        mapping = {
            "门诊次均药品费用（元）": "op_avg_cost", "门诊次均药品费用(不含中药饮片)(元)": "op_avg_cost_no_herb",
            "门诊药占比（%）": "op_drug_ratio", "门诊药占比(不含中药饮片)（%）": "op_drug_ratio_no_herb",
            "住院次均药品费用（元）": "ip_avg_cost", "住院次均药品费用(不含中药饮片)(元)": "ip_avg_cost_no_herb",
            "住院药占比（%）": "ip_drug_ratio", "住院药占比(不含中药饮片)（%）": "ip_drug_ratio_no_herb",
            "抗菌药物使用率(%)": "antibiotic_rate", "抗菌药物使用强度": "antibiotic_intensity",
            "门诊基本药物占处方用药百分率(%)": "op_basic_drug", "住院基本药物金额所占比例(不含中药饮片)(%)": "ip_basic_drug_amount",
            "住院患者静脉输液使用率（%）": "inf_rate", "住院患者人均静脉输液天数": "inf_days",
            "住院患者平均每床日使用静脉输液体积(ml)": "inf_vol", "住院患者人均静脉输液药品品种数": "inf_types",
            "重点监控品种收入占比（%）": "key_monitor"
        }
        for k_cn, k_en in mapping.items(): final_data[k_cn] = raw_data.get(k_en, '-')
        return final_data
