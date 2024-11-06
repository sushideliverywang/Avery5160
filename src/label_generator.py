from fpdf import FPDF
import pandas as pd

# 从 Excel 文件加载数据
excel_file = "49-C241101.xlsx"
data = pd.read_excel(excel_file)

# 生成标签内容
label_contents = []
for _, row in data.iterrows():
    line1 = f"{row['Lot ID']} ({row['Unnamed: 8']})"  # 例如: C241101-01 (1166.5483)
    line2 = f"{row['Brand']} {row['Subcategory']}"    # 例如: SS Washers
    line3 = f"{row['Item Description']} ${row['Unit Retail']}"  # 例如: WF4BSBOW $999.99
    label_content = f"{line1}\n{line2}\n{line3}"
    label_contents.append(label_content)

# 初始化 PDF 文档
pdf = FPDF(format="letter")
pdf.set_auto_page_break(auto=True, margin=0)  # 添加这行，禁用自动分页

# PDF 文档和标签配置参数
MARGIN_X = 2.7625          # 左边距
MARGIN_Y = 9.2            # 上边距
LABEL_WIDTH = 66.675       # 标签宽度
LABEL_HEIGHT = 26        # 标签高度
horizontal_spacing = 5.175  # 列间距
COLUMNS = 3
ROWS = 10
LABELS_PER_PAGE = COLUMNS * ROWS

# 文本布局参数
FONT_SIZE = 14             # 字体大小
LINE_HEIGHT = 5            # 行高
LINE_SPACING = 2           # 行间距参数
X_TEXT_OFFSET = 2          # 文本左边距
Y_INITIAL_OFFSET = 3       # 首行文本上边距

# 颜色参数（RGB值，0-255）
BACKGROUND_COLOR = (235, 245, 255)  # 淡蓝色
TEXT_COLOR = (255, 0, 0)            # 红色

# 设置字体
pdf.add_page()
pdf.set_font("Helvetica", style='B', size=FONT_SIZE)  # 添加 'B' 使字体加粗

# 生成标签
for i, label_content in enumerate(label_contents):
    # 计算当前页内的标签位置
    label_index = i % LABELS_PER_PAGE
    row = label_index // COLUMNS
    col = label_index % COLUMNS

    # 如果需要新页面
    if label_index == 0 and i > 0:
        pdf.add_page()

    # 计算 x 和 y 位置
    x = MARGIN_X + col * (LABEL_WIDTH + horizontal_spacing)
    y = MARGIN_Y + row * LABEL_HEIGHT

    # 先填充背景色
    pdf.set_fill_color(*BACKGROUND_COLOR)
    pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, style='F')  # 'F' 表示只填充
    
    # 绘制标签的边框
    #pdf.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT)
    
    # 设置文本颜色
    pdf.set_text_color(*TEXT_COLOR)
    
    # 分别处理三行文本
    lines = label_content.split('\n')
    
    # 打印文本
    for line_num, line in enumerate(lines):
        y_position = y + Y_INITIAL_OFFSET + (LINE_HEIGHT + LINE_SPACING) * line_num
        pdf.set_xy(x + X_TEXT_OFFSET, y_position)
        pdf.cell(LABEL_WIDTH - 2, LINE_HEIGHT, line, ln=0)

# 保存 PDF
pdf_file_name = "avery_5160_labels_perfect_version.pdf"
pdf.output(pdf_file_name)

print(f"PDF 文件 '{pdf_file_name}' 已成功创建！")