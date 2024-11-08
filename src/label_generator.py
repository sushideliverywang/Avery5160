from fpdf import FPDF
import pandas as pd
from config import Avery5160, CUSTOMER_DEFAULTS
import os

def load_data(excel_file):
    """
    从 Excel 文件加载数据，并进行验证
    """
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"Excel file {excel_file} not found.")

    try:
        # 尝试读取 Excel 文件
        data = pd.read_excel(excel_file)
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {e}")

    # 验证数据是否包含所需的列
    required_columns = ['Lot ID', 'safe code', 'Brand', 'Subcategory', 'Item Description', 'Unit Retail']
    for column in required_columns:
        if column not in data.columns:
            raise ValueError(f"Missing required column: {column}")

    return data

def generate_label_contents(data):
    """
    生成标签内容
    """
    label_contents = []
    required_columns = ['Lot ID', 'safe code', 'Brand', 'Subcategory', 'Item Description', 'Unit Retail']
    
    for _, row in data.iterrows():
        # 验证每一行数据是否包含所需的字段
        for column in required_columns:
            if pd.isna(row[column]):
                raise ValueError(f"Missing value in column: {column} for row: {row}")

        # 生成标签内容
        line1 = f"{row['Lot ID']} ({row['safe code']})"  # 例如: C241101-01 (1166.5483)
        line2 = f"{row['Brand']} {row['Subcategory']}"    # 例如: SS Washers
        line3 = f"{row['Item Description']} ${row['Unit Retail']}"  # 例如: WF4BSBOW $999.99
        label_content = f"{line1}\n{line2}\n{line3}"
        label_contents.append(label_content)
    
    return label_contents

def create_pdf(label_contents, template, customer, output_file):
    """
    创建 PDF 文件
    """
    pdf = FPDF(format="letter")
    pdf.set_auto_page_break(auto=True, margin=0)  # 添加这行，禁用自动分页

    # 使用模板参数
    MARGIN_X = template["MARGIN_X"]
    MARGIN_Y = template["MARGIN_Y"]
    LABEL_WIDTH = template["LABEL_WIDTH"]
    LABEL_HEIGHT = template["LABEL_HEIGHT"]
    horizontal_spacing = template["HORIZONTAL_SPACING"]
    COLUMNS = template["COLUMNS"]
    ROWS = template["ROWS"]
    LABELS_PER_PAGE = COLUMNS * ROWS
    LINE_HEIGHT = template["LINE_HEIGHT"]
    LINE_SPACING = template["LINE_SPACING"]
    X_TEXT_OFFSET = template["X_TEXT_OFFSET"]
    Y_TEXT_OFFSET = template["Y_TEXT_OFFSET"]

    # 使用客户自定义参数
    FONT_SIZE = customer["FONT_SIZE"]
    FONT_STYLE = customer.get("FONT_STYLE", "")  # 获取字体样式，默认为空
    BACKGROUND_COLOR = customer["BACKGROUND_COLOR"]
    TEXT_COLOR = customer["TEXT_COLOR"]

    # 设置文本颜色字体
    pdf.set_text_color(*TEXT_COLOR)
    pdf.set_font(customer["FONT"], size=FONT_SIZE, style=FONT_STYLE)
    # 先填充背景色
    pdf.set_fill_color(*BACKGROUND_COLOR)

    for i, content in enumerate(label_contents):
        if i % LABELS_PER_PAGE == 0:
            pdf.add_page()
        col = (i % LABELS_PER_PAGE) % COLUMNS
        row = (i % LABELS_PER_PAGE) // COLUMNS
        x = MARGIN_X + col * (LABEL_WIDTH + horizontal_spacing)
        y = MARGIN_Y + row * LABEL_HEIGHT
        pdf.set_xy(x, y)
        # 绘制标签的边框，设定文字起始位置
        pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, style='F')  # 'F' 表示只填充
        #pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, style='D')  # 'D' 表示只绘制边框
        pdf.set_xy(x + X_TEXT_OFFSET, y + Y_TEXT_OFFSET)  # 设置文本的起始位置
        pdf.multi_cell(LABEL_WIDTH - X_TEXT_OFFSET, LINE_HEIGHT + LINE_SPACING, content, border=0, fill=False)

    pdf.output(output_file)
    print(f"PDF 文件 '{output_file}' 已成功创建！")

# 使用新函数加载数据
excel_file = ".\\exsample\\49-C241101.xlsx"
data = load_data(excel_file)

# 使用新函数生成标签内容
label_contents = generate_label_contents(data)

# 使用新函数创建 PDF 文件
create_pdf(label_contents, Avery5160, CUSTOMER_DEFAULTS, "Avery5160_label.pdf")

