import subprocess
import argparse
import os
# from common.utils import get_metadata
# from common import __main__

# 调用pandoc
def call_pandoc(args):
    # 获取输入文件的绝对路径
    input_path = os.path.abspath(args.input_file)
    template_path = os.path.abspath(args.template)
    if not args.output:
        # 没有指定输出路径，则使用输入文件的路径和文件名
        output_path = os.path.splitext(input_path)[0] + ".docx"
    else:
        output_path = os.path.abspath(args.output)
    # 构建 pandoc 命令
    command = ["pandoc", "-s", input_path, "--reference-doc", template_path, "-o", output_path]
    # 执行命令
    subprocess.run(command, cwd=os.path.dirname(input_path))

def call_handle(args):
    input_path = os.path.abspath(args.input_file)
    template_path = os.path.abspath(args.template)
    command = ["python3", "-m", "common", input_path, "exp"]
    subprocess.run(command)
    


# 打开文件路径
def open_dir(file_path):
    '''在文件管理器中显示生成的文件'''
    pass

# 命令行主体
parser = argparse.ArgumentParser(prog='md4report')
# 输入文件
parser.add_argument(
    "input_file",
    help="input markdown file",
    # required=True
)
# 引用文件
parser.add_argument(
    "-t",
    "--template",
    help="template name",
    default="./common/ref/exp.docx",
    required=False
)
# 输出文件
parser.add_argument(
    "-o",
    "--output",
    help="output docx file",
    required=False
)
args = parser.parse_args()

# 调用 pandoc

call_pandoc(args)
call_handle(args)