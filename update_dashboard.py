import pandas as pd
import re
import os
import sys

# 文件路径配置
# 获取脚本所在目录的绝对路径，确保在任何地方运行都能找到文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, '进度汇总.xlsx')
HTML_FILE = os.path.join(BASE_DIR, 'dashboard.html')

def update_dashboard():
    print("-" * 30)
    print(f"工作目录: {BASE_DIR}")
    
    # 1. 检查文件是否存在
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ 错误：找不到 Excel 文件: {EXCEL_FILE}")
        return
    if not os.path.exists(HTML_FILE):
        print(f"❌ 错误：找不到 HTML 文件: {HTML_FILE}")
        return

    print(f"📖 正在读取 {os.path.basename(EXCEL_FILE)} ...")
    
    try:
        # 2. 使用 Pandas 读取 Excel
        # keep_default_na=False 确保空单元格是空字符串而不是 NaN
        df = pd.read_excel(EXCEL_FILE, keep_default_na=False)
        
        # 简单的数据清洗
        # 确保所有内容都是字符串格式，避免 JSON/CSV 转换时的类型问题
        df = df.astype(str)
        
        # 处理日期格式中可能出现的 " 00:00:00"
        if '发版日期' in df.columns:
            df['发版日期'] = df['发版日期'].str.replace(' 00:00:00', '', regex=False)
        
        # 替换 CSV 中的特殊字符，防止破坏格式
        df = df.replace({'\n': ' ', '\r': ''}, regex=True)

        # 转为 CSV 格式字符串
        csv_data = df.to_csv(index=False)
        
        print(f"✅ 读取成功，共 {len(df)} 条数据")
        
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 3. 读取现有的 HTML 内容
    try:
        with open(HTML_FILE, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except Exception as e:
        print(f"❌ 读取 HTML 失败: {e}")
        return

    # 4. 使用正则替换 HTML 中的数据部分
    # 目标是替换 const rawData = `...`; 中间的内容
    # 使用 re.DOTALL 让 . 可以匹配换行符
    pattern = r'(const\s+rawData\s*=\s*`)([^`]*)(`;)'
    
    # 检查是否找到标记
    if not re.search(pattern, html_content, re.DOTALL):
        print("❌ 错误：在 HTML 中找不到数据标记 (const rawData = `...`)")
        print("请确认 dashboard.html 中包含 const rawData = `...`; 代码块")
        return

    # 执行替换
    # group(1) 是 "const rawData = `"
    # csv_data 是新数据
    # group(3) 是 "`;"
    new_html_content = re.sub(pattern, lambda m: m.group(1) + csv_data + m.group(3), html_content, flags=re.DOTALL)

    # 5. 保存回 HTML 文件
    try:
        with open(HTML_FILE, 'w', encoding='utf-8') as f:
            f.write(new_html_content)
        print(f"🎉 成功！已将 Excel 数据更新到 {os.path.basename(HTML_FILE)}")
    except Exception as e:
        print(f"❌ 写入 HTML 失败: {e}")

if __name__ == '__main__':
    # 检查依赖
    try:
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"缺少依赖库: {e.name}")
        print("请运行: pip install pandas openpyxl")
        sys.exit(1)
        
    update_dashboard()

