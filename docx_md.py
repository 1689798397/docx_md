import glob
import os


# docx转md
def docx_md(docx_path):
    for docx_abs_path in glob.glob(docx_path + r'\**\*.docx', recursive=True):
        # 切换工作目录
        os.chdir(os.path.dirname(docx_abs_path))
        # 图片名
        pic_dir_name = os.path.basename(docx_abs_path).replace(".docx", "")
        # md文件名
        md_abs_path = docx_abs_path.replace(".docx", ".md")
        cmd_str = "pandoc --extract-media " + pic_dir_name + " -s " + docx_abs_path + " -o " + md_abs_path
        print(cmd_str)
        os.system(cmd_str)


# md转docx
def md_docx(md_path):
    for md_abs_path in glob.glob(md_path + r'\**\*.md', recursive=True):
        # 切换工作目录
        os.chdir(os.path.dirname(md_abs_path))
        # docx文件名
        docx_abs_path = md_abs_path.replace(".md", ".docx")
        cmd_str = "pandoc -s " + md_abs_path + " -o " + docx_abs_path
        print(cmd_str)
        os.system(cmd_str)


# html转md
def html_md(html_path):
    for md_abs_path in glob.glob(html_path + r'\**\*.md', recursive=True):
        # 切换工作目录
        os.chdir(os.path.dirname(md_abs_path))
        # docx文件名
        md_abs_path = md_abs_path.replace(".html", ".md")
        cmd_str = "pandoc -f html-native_divs-native_spans -i " + html_path + " -t markdown -o " + md_abs_path
        print(cmd_str)
        os.system(cmd_str)


if __name__ == "__main__":
    path = input("请输入包含docx文件的路径：")
    while not os.path.exists(path):
        path = input("路径有误请重新输入：")
    target = input("请输入转换方式\n\t\t1:docx转markdown\n\t\t2:markdown转docx\n\t\t3:html转markdown\n请选择：")
    target_list = ["1", "2", "3"]
    while target not in target_list:
        path = input("转换方式有误请重新输入：")
    if target == "1":
        docx_md(path)
    elif target == "2":
        md_docx(path)
    elif target == "3":
        html_md(path)
