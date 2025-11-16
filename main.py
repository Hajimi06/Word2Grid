# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。from docx import Document
from PIL import Image, ImageDraw, ImageFont

# ===== 基本参数配置 =====
DOCX_PATH = "input.docx"        # 输入 Word 文件路径
OUTPUT_IMG = "output_grid.png"  # 输出图片文件名

COLS_PER_LINE = 23              # 每行格子数
CELL_SIZE = 40                  # 每个格子边长（像素）
MARGIN = 40                     # 图片四周留白（像素）

# 字体路径：请改成你自己电脑上的中文字体路径
# Windows 上可以用：C:/Windows/Fonts/simhei.ttf 或 simsun.ttc 等
FONT_PATH = "C:/Windows/Fonts/simhei.ttf"
FONT_SIZE = 24


def read_docx_text(path):
    """
    从 docx 中读取段落文本，返回段落列表（去掉空段落）
    """
    doc = Document(path)
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def layout_to_grid(paragraphs, cols=23):
    """
    将段落文本排版到格子中。
    规则：
    - 每行最多 cols 个格子
    - 每一段的首行都空两个格
    - 标点（逗号、冒号、引号等）之间尽量挤到同一格：
        - 如果当前是标点，且前一个格子是“纯标点格”，并且长度 < 2，就合并到前一个格子
        - 所以 '，：'、'：“'、'，”' 等可以出现在同一个格子
    """
    lines = []
    current_line = []

    # 认为是标点的字符集合（可以按需要再加）
    PUNCS = set('，。、！？：；、“”‘’（）()《》〈〉——…,:;"\'')

    for text in paragraphs:
        chars = list(text)

        # 每一段开始前，如果当前行有内容，先收尾换行
        if current_line:
            lines.append(current_line)
            current_line = []

        # 每一段首行空两个格
        current_line.extend(["", ""])

        for ch in chars:
            is_punc = ch in PUNCS

            # 如果是标点，尝试和前一个格子合并
            if is_punc and current_line:
                last = current_line[-1]
                # 前一个格子是“纯标点格”且里面标点数量 < 2
                if last != "" and all(c in PUNCS for c in last) and len(last) < 2:
                    current_line[-1] = last + ch   # 例如 '，' + '：' -> '，：'
                    continue

            # 否则正常进入新格子
            if len(current_line) >= cols:
                lines.append(current_line)
                current_line = []
            current_line.append(ch)

    # 收尾最后一行
    if current_line:
        lines.append(current_line)

    # 补齐到固定列数
    padded_lines = []
    for line in lines:
        if len(line) < cols:
            line = line + [""] * (cols - len(line))
        else:
            line = line[:cols]
        padded_lines.append(line)

    return padded_lines


def draw_grid_image(lines, cols=23, cell_size=40, margin=40,
                    font_path=FONT_PATH, font_size=24, out_path="output.png"):
    """
    根据排版结果生成方格纸图片，并写入文字。
    兼容新版 Pillow（使用 textbbox 代替 textsize）。
    """
    rows = len(lines)

    img_width = cols * cell_size + margin * 2
    img_height = rows * cell_size + margin * 2

    img = Image.new("RGB", (img_width, img_height), "white")
    draw = ImageDraw.Draw(img)

    # 加载字体
    font = ImageFont.truetype(font_path, font_size)

    # 画方格线
    left = margin
    top = margin
    right = margin + cols * cell_size
    bottom = margin + rows * cell_size

    # 竖线
    for c in range(cols + 1):
        x = left + c * cell_size
        draw.line([(x, top), (x, bottom)], fill="black", width=1)

    # 横线
    for r in range(rows + 1):
        y = top + r * cell_size
        draw.line([(left, y), (right, y)], fill="black", width=1)

    # 写字符（居中）
    for r, line in enumerate(lines):
        for c, ch in enumerate(line):
            if not ch:
                continue  # 空格子不写字
            cell_x = left + c * cell_size
            cell_y = top + r * cell_size

            # 新版 Pillow：用 textbbox 计算文字尺寸
            bbox = draw.textbbox((0, 0), ch, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]

            text_x = cell_x + (cell_size - w) / 2
            text_y = cell_y + (cell_size - h) / 2

            draw.text((text_x, text_y), ch, font=font, fill="black")

    img.save(out_path)
    print(f"已保存图片到: {out_path}")

def main():
    # 1. 读取 Word 文本
    paragraphs = read_docx_text(DOCX_PATH)

    # 2. 布局到格子（每段首行空两格）
    lines = layout_to_grid(
        paragraphs,
        cols=COLS_PER_LINE,
    )

    # 3. 画图输出
    draw_grid_image(
        lines,
        cols=COLS_PER_LINE,
        cell_size=CELL_SIZE,
        margin=MARGIN,
        font_path=FONT_PATH,
        font_size=FONT_SIZE,
        out_path=OUTPUT_IMG
    )


if __name__ == "__main__":
    main()



# 按装订区域中的绿色按钮以运行脚本。
if __name__ == '__main__':
    print_hi('PyCharm')

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
