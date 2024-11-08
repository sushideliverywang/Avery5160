from fpdf import FPDF
import pandas as pd
from config import Avery5160, CUSTOMER_DEFAULTS

def load_data(excel_file):
    """
    从 Excel 文件加载数据
    """
    return pd.read_excel(excel_file)

def generate_label_contents(data):
    """
    生成标签内容
    """
    label_contents = []
    for _, row in data.iterrows():
        line1 = f"{row['Lot ID']} ({row['Unnamed: 8']})"  # 例如: C241101-01 (1166.5483)
        line2 = f"{row['Brand']} {row['Subcategory']}"    # 例如: SS Washers
        line3 = f"{row['Item Description']} ${row['Unit Retail']}"  # 例如: WF4BSBOW $999.99
        label_content = f"{line1}\n{line2}\n{line3}"
        label_contents.append(label_content)
    return label_contents

def create_pdf(label_contents, template=Avery5160, customer=CUSTOMER_DEFAULTS, output_file="labels.pdf"):
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
        # 绘制标签的边框
        pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, style='F')  # 'F' 表示只填充
        pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, style='D')  # 'D' 表示只绘制边框
        pdf.set_xy(x + X_TEXT_OFFSET, y + Y_TEXT_OFFSET)  # 设置文本的起始位置
        pdf.multi_cell(LABEL_WIDTH - X_TEXT_OFFSET, LINE_HEIGHT + LINE_SPACING, content, border=0, fill=False)

    pdf.output(output_file)

# 使用新函数加载数据
excel_file = ".\\exsample\\49-C241101.xlsx"
data = load_data(excel_file)

# 使用新函数生成标签内容
label_contents = generate_label_contents(data)

# 使用新函数创建 PDF 文件
create_pdf(label_contents)

print(f"PDF 文件 'labels.pdf' 已成功创建！")