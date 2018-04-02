#!/usr/bin/env python3

import os
import click

from win32com import client


@click.command()
@click.option("-f", "--filename_list", help="please input problem name(list)", prompt="problem list", type=click.Tuple)
@click.option("-i", "--fileout", help="please input problem outfile name", default="output/merge.docx")
def merge_docx(filename_list, fileout="output/merge.docx"):
    # 初始化
    click.echo(filename_list)
    print(type(filename_list))
    word = client.Dispatch("Word.Application")
    word.Visible = False
    # 建立一个新的word临时
    newdoc = word.Documents.Add()
    for file in filename_list[::-1]:
        # 把每道题都存进去合并
        if file.endswith("docx"):
            file = os.path.join(os.getcwd(), "problems/{}".format(file))
        else:
            file = os.path.join(os.getcwd(), "problems/{}.docx".format(file))
        newdoc.Application.Selection.Range.InsertFile(file)

    # 保存文件
    newdoc.SaveAs(os.path.join(os.getcwd(), fileout))
    # 关闭文件
    newdoc.Close()
    word.Quit()
    print(u"出题完毕，请在output文件中查看")

if __name__ == '__main__':
    # merge_docx(["math-1", "math-2", "math-3"])
    merge_docx()
