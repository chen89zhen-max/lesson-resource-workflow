"""Copy the generic editable starter; filling content still requires a PPTX editor."""
import argparse
from pathlib import Path
import shutil


def main():
    parser = argparse.ArgumentParser(description='建立六类教学版式的可编辑起稿副本，不自动编写教学内容。')
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    template = Path(__file__).resolve().parent.parent / 'templates/vocational_teaching.pptx'
    output = args.output.expanduser().resolve()
    if output.suffix.lower() != '.pptx':
        parser.error('输出扩展名必须为 .pptx')
    if output.exists():
        parser.error('输出已存在，请使用新文件名。')
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, output)
    print(f'已建立起稿副本：{output}\n仍需替换占位内容、按课程增删页面，并完成渲染与看图检查。')


if __name__ == '__main__':
    main()
