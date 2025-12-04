import streamlit as st
import pandas as pd
import os
import time
from docx import Document
from io import BytesIO
import numpy as np

# ==========================================
# 工具函数
# ==========================================
def smart_load_file(uploaded_file, header_keywords=None, header_idx_fallback=0):
    """
    智能读取上传的文件 (BytesIO)
    """
    if uploaded_file is None:
        return None
    
    try:
        # 获取文件扩展名
        filename = uploaded_file.name
        ext = os.path.splitext(filename)[1].lower()
        
        # 1. 读取前几行以查找表头
        if ext == '.csv':
            try:
                df_raw = pd.read_csv(uploaded_file, header=None, nrows=20, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, nrows=20, encoding='gbk')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, nrows=20)
            
        # 2. 定位表头
        header_idx = header_idx_fallback
        if header_keywords:
            for i, row in df_raw.iterrows():
                row_str = " ".join(row.astype(str).values)
                if any(k in row_str for k in header_keywords):
                    header_idx = i
                    break
        
        # 3. 重新读取完整数据
        uploaded_file.seek(0)
        if ext == '.csv':
            try:
                df = pd.read_csv(uploaded_file, header=header_idx, encoding='utf-8')
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_idx, encoding='gbk')
        else:
            df = pd.read_excel(uploaded_file, header=header_idx)
            
        # 清理列名
        df.columns = [str(c).strip() for c in df.columns]
        return df

    except Exception as e:
        st.error(f"读取文件失败: {filename} - {str(e)}")
        return None

def safe_div(n, d):
    try: return n / d if d and d != 0 else 0
    except: return 0

# ==========================================
# 核心处理逻辑 (与之前相同，只是去掉了GUI交互)
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
        # 1. 输液表
        df_inf = smart_load_file(inf_file, header_keywords=['指标名称'], header_idx_fallback=3)
        if df_inf is not None:
            for _, row in df_inf.iterrows():
                name = str(row.get('指标名称', '')).strip()
                val = row.get('指标值', 0)
                if name == "住院患者静脉输液使用率①(100%)": raw_data["inf_rate"] = val
                elif name == "住院患者人均静脉输液天数": raw_data["inf_days"] = val
                elif name == "住院患者平均每床日使用静脉输液体积(ml)": raw_data["inf_vol"] = val
                elif name == "住院患者人均静脉输液药品品种数": raw_data["inf_types"] = val

        # 2. 质控表
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

class DepartmentStatsMerger:
    def __init__(self):
        self.output_columns = [
            '包含科室名称', '使用抗菌药物的病人数(例)', '参与统计病人数(例)', '病人药品总金额(元)', 
            '病人药品总金额(不含中药饮片)(元)', '病人治疗总金额(元)', '基本药物总金额(不含中药饮片)(元)',
            '住院患者抗菌药物使用量(DDDs)①', '同期收治患者人天数(人天)①', '住院重点监控品种药品金额(元)', 
            '11月不良反应', '11月严重或新的', '静脉输液总体积(ml) (G)', '次均药品费用(不含中药饮片)', 
            '抗菌药物使用率', '抗菌药物使用强度', '药占比', '基药比', '重点监控药品收入占比(%)', 
            '中药饮片金额（元）', '中药饮片使用率(%)', '药品不良反应合计', '住院患者平均每床日使用静脉输液体积(ml)'
        ]

    def process_dept_data(self, qc_file, adr_file, inf_file):
        # 加载数据
        df_qc = self._load_qc(qc_file)
        if df_qc is None: return pd.DataFrame()
        
        df_adr = self._load_adr(adr_file)
        df_inf = self._load_inf(inf_file)

        # 合并
        df_merged = pd.merge(df_qc, df_adr, on='包含科室名称', how='left')
        if df_inf is not None:
            df_merged = pd.merge(df_merged, df_inf, on='包含科室名称', how='left')
        
        # 补齐列
        for c in self.output_columns:
            if c not in df_merged.columns: df_merged[c] = 0
        numeric_cols = [c for c in self.output_columns if c != '包含科室名称']
        for c in numeric_cols: df_merged[c] = pd.to_numeric(df_merged[c], errors='coerce').fillna(0)

        # 补入药学部
        if df_adr is not None and not df_adr.empty:
            pharmacy = df_adr[df_adr['包含科室名称'].str.contains('药学部', na=False)].copy()
            if not pharmacy.empty:
                for c in self.output_columns:
                    if c not in pharmacy.columns: pharmacy[c] = 0
                df_merged = pd.concat([df_merged, pharmacy], ignore_index=True)

        def calc(row):
            r = row.copy()
            def div(n, d): return n/d if d>0 else 0
            r['抗菌药物使用率'] = div(r['使用抗菌药物的病人数(例)'], r['参与统计病人数(例)'])*100
            r['抗菌药物使用强度'] = div(r['住院患者抗菌药物使用量(DDDs)①'], r['同期收治患者人天数(人天)①'])*100
            r['药占比'] = div(r['病人药品总金额(不含中药饮片)(元)'], r['病人治疗总金额(元)'])*100
            r['基药比'] = div(r['基本药物总金额(不含中药饮片)(元)'], r['病人药品总金额(不含中药饮片)(元)'])*100
            r['重点监控药品收入占比(%)'] = div(r['住院重点监控品种药品金额(元)'], r['病人药品总金额(元)'])*100
            r['中药饮片金额（元）'] = r['病人药品总金额(元)'] - r['病人药品总金额(不含中药饮片)(元)']
            r['中药饮片使用率(%)'] = div(r['中药饮片金额（元）'], r['病人药品总金额(元)'])*100
            r['药品不良反应合计'] = r['11月不良反应'] + r['11月严重或新的']
            r['住院患者平均每床日使用静脉输液体积(ml)'] = div(r['静脉输液总体积(ml) (G)'], r['同期收治患者人天数(人天)①'])
            r['次均药品费用(不含中药饮片)'] = div(r['病人药品总金额(不含中药饮片)(元)'], r['参与统计病人数(例)'])
            return r

        m = df_merged.apply(calc, axis=1)
        
        # 住院汇总
        total_row = m[numeric_cols].sum().to_dict()
        total_row['包含科室名称'] = '住院汇总'
        m_total = pd.DataFrame([calc(pd.Series(total_row))])
        
        return pd.concat([m[self.output_columns], m_total[self.output_columns]], ignore_index=True)

    def _load_qc(self, file):
        # 关键词：使用抗菌药物的病人数
        df = smart_load_file(file, header_keywords=['使用抗菌药物的病人数'], header_idx_fallback=5)
        if df is None: return None
        try:
            # 尝试通过列名映射，或者fallback到列索引
            # 优先检查列索引是否有效
            if df.shape[1] > 2:
                col_map = {2:'包含科室名称', 4:'使用抗菌药物的病人数(例)', 5:'参与统计病人数(例)', 7:'病人药品总金额(元)', 
                           11:'病人治疗总金额(元)', 13:'病人药品总金额(不含中药饮片)(元)', 16:'基本药物总金额(不含中药饮片)(元)', 
                           19:'住院患者抗菌药物使用量(DDDs)①', 20:'同期收治患者人天数(人天)①', 22:'住院重点监控品种药品金额(元)'}
                new_cols = {df.columns[k]: v for k, v in col_map.items() if k < df.shape[1]}
                df = df.rename(columns=new_cols)
                if '包含科室名称' in df.columns:
                    df['包含科室名称'] = df['包含科室名称'].astype(str).str.strip()
                    return df[df['包含科室名称'] != 'nan']
        except: pass
        return None

    def _load_adr(self, file):
        df = smart_load_file(file, header_keywords=['不良反应'], header_idx_fallback=3)
        if df is None: return pd.DataFrame(columns=['包含科室名称', '11月不良反应', '11月严重或新的'])
        try:
            if df.shape[1] > 23:
                cols = df.columns.tolist()
                df = df.rename(columns={cols[0]: '包含科室名称', cols[22]: '11月不良反应', cols[23]: '11月严重或新的'})
                df = df[['包含科室名称', '11月不良反应', '11月严重或新的']]
                df['包含科室名称'] = df['包含科室名称'].astype(str).str.strip()
                return df
        except: pass
        return pd.DataFrame(columns=['包含科室名称', '11月不良反应', '11月严重或新的'])

    def _load_inf(self, file):
        df = smart_load_file(file, header_keywords=['科室', '体积'], header_idx_fallback=3)
        if df is None: return None
        dept = next((c for c in df.columns if "科室" in c), None)
        vol = next((c for c in df.columns if "总体积" in c), None)
        if dept and vol:
            df = df[[dept, vol]].rename(columns={dept:'包含科室名称', vol:'静脉输液总体积(ml) (G)'})
            df['包含科室名称'] = df['包含科室名称'].astype(str).str.strip()
            return df
        return None

class WordReportGenerator:
    def fill_cell_by_dept(self, table, data_dict, metric_keys):
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                txt = cell.text.strip().replace('\n','').replace(' ','')
                matched = None
                if txt in data_dict: matched = txt
                else:
                    for d in data_dict.keys():
                        if (d in txt or txt in d) and len(txt) > 2:
                            matched = d; break
                if matched:
                    for k_idx, key in enumerate(metric_keys):
                        if i+1+k_idx < len(row.cells):
                            val = data_dict[matched].get(key, '-')
                            if isinstance(val, (int, float)):
                                row.cells[i+1+k_idx].text = f"{val:.2f}"
                            else: row.cells[i+1+k_idx].text = str(val)

    def generate(self, hospital_df, dept_df, template_file):
        doc = Document(template_file)
        
        # 1. 全院表
        h_data = hospital_df.set_index('指标名称').to_dict('index')
        if len(doc.tables) > 0:
            for row in doc.tables[0].rows:
                metric = row.cells[0].text.strip().replace(' ','')
                key = next((k for k in h_data.keys() if metric in k or k in metric), None)
                if key:
                    for idx, c in enumerate(['2025年11月', '2025年1-11月', '2024年']):
                        if idx+1 < len(row.cells):
                            val = h_data[key].get(c, '-')
                            row.cells[idx+1].text = f"{val:.2f}" if isinstance(val, (int, float)) else str(val)

        # 2. 科室表
        d_data = dept_df.set_index('包含科室名称').to_dict('index')
        for table in doc.tables[1:]:
            try: h_txt = "".join([c.text for c in table.rows[0].cells]).replace(" ","").replace("\n","")
            except: continue

            if "次均药品费用" in h_txt: self.fill_cell_by_dept(table, d_data, ['次均药品费用(不含中药饮片)'])
            elif "使用率" in h_txt and "使用强度" in h_txt and "中药" not in h_txt:
                self.fill_cell_by_dept(table, d_data, ['抗菌药物使用率', '抗菌药物使用强度'])
            elif "药占比" in h_txt: self.fill_cell_by_dept(table, d_data, ['药占比'])
            elif "基药" in h_txt: self.fill_cell_by_dept(table, d_data, ['基药比'])
            elif "重点监控" in h_txt: self.fill_cell_by_dept(table, d_data, ['重点监控药品收入占比(%)'])
            elif "中药" in h_txt and ("金额" in h_txt or "使用率" in h_txt):
                self.fill_cell_by_dept(table, d_data, ['中药饮片使用率(%)', '中药饮片金额（元）', '病人药品总金额(元)'])
            elif "不良反应" in h_txt: self.fill_cell_by_dept(table, d_data, ['11月不良反应', '11月严重或新的', '药品不良反应合计'])
            elif ("输液" in h_txt and "体积" in h_txt) or ("输液" in h_txt and "ml" in h_txt.lower()):
                self.fill_cell_by_dept(table, d_data, ['住院患者平均每床日使用静脉输液体积(ml)'])

        # 保存到内存流
        f = BytesIO()
        doc.save(f)
        f.seek(0)
        return f

# ==========================================
# Streamlit 界面
# ==========================================
def main():
    st.set_page_config(page_title="医院药事月报生成器", layout="wide")
    st.title("🏥 医院药事质控月报生成系统")
    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.header("📂 1. 全院数据上传")
        st.info("用于生成Word报告中的第一张表（全院整体情况）")
        
        h_files = {}
        for p in ["2025年11月", "2025年1-11月", "2024年"]:
            st.subheader(f"📅 {p}")
            inf = st.file_uploader(f"[{p}] 静脉输液表", type=['xlsx', 'xls', 'csv'], key=f"inf_{p}")
            qc = st.file_uploader(f"[{p}] 数据质控表", type=['xlsx', 'xls', 'csv'], key=f"qc_{p}")
            if inf and qc:
                h_files[p] = {"inf": inf, "qc": qc}

    with col2:
        st.header("📂 2. 科室数据上传")
        st.info("用于生成Word报告中的各科室明细表")
        
        qc_dept = st.file_uploader("1. 住院质控数据(大科室)", type=['xlsx', 'xls', 'csv'])
        adr_dept = st.file_uploader("2. 不良反应数据", type=['xlsx', 'xls', 'csv'])
        inf_dept = st.file_uploader("3. 静脉输液202511", type=['xlsx', 'xls', 'csv'])
        
        st.header("📝 3. 模板上传")
        template = st.file_uploader("Word 模板文件 (.docx)", type=['docx'])

    st.markdown("---")
    
    # 开始处理按钮
    if st.button("🚀 开始生成报告", type="primary"):
        if not (len(h_files) == 3 and qc_dept and adr_dept and inf_dept and template):
            st.error("请先上传所有必需的文件！")
            return

        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # 1. 处理全院数据
            status_text.text("正在计算全院指标...")
            h_proc = HospitalStatsProcessor()
            h_all_data = {}
            for p, files in h_files.items():
                h_all_data[p] = h_proc.extract_data(files['inf'], files['qc'])
            
            df_hospital = pd.DataFrame(h_proc.output_rows, columns=["指标名称"])
            for p in ["2025年11月", "2025年1-11月", "2024年"]:
                df_hospital[p] = [h_all_data[p].get(r, '-') for r in h_proc.output_rows]
            
            progress_bar.progress(40)

            # 2. 处理科室数据
            status_text.text("正在计算科室指标...")
            d_proc = DepartmentStatsMerger()
            df_dept = d_proc.process_dept_data(qc_dept, adr_dept, inf_dept)
            
            progress_bar.progress(70)

            # 3. 生成Word
            status_text.text("正在填充 Word 报告...")
            gen = WordReportGenerator()
            word_io = gen.generate(df_hospital, df_dept, template)
            
            progress_bar.progress(100)
            status_text.success("🎉 处理完成！请在下方下载结果。")

            # 4. 下载区域
            st.subheader("📥 结果下载")
            c1, c2, c3 = st.columns(3)
            
            # Excel - 全院
            buffer_h = BytesIO()
            df_hospital.to_excel(buffer_h, index=False)
            c1.download_button("下载 全院指标.xlsx", buffer_h.getvalue(), "全院指标.xlsx")
            
            # Excel - 科室
            buffer_d = BytesIO()
            df_dept.to_excel(buffer_d, index=False)
            c2.download_button("下载 科室明细.xlsx", buffer_d.getvalue(), "科室明细.xlsx")
            
            # Word
            c3.download_button("下载 最终质控报告.docx", word_io.getvalue(), "11月药事质控报告.docx")

        except Exception as e:
            st.error(f"发生错误: {e}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()