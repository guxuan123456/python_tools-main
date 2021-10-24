"""
1. 如何设置编辑器字体的大小?
File(文件)-> Settings(设置) -> Editor(编辑器) -> Font(字体), 修改字体的大小
2. 注释: 代码的解释说明
3. 批量处理实现PPT转PDF
"""

# 1). 导入需要的模块(打开应用程序的模块)
from pprint import PrettyPrinter
from typing import Mapping
import win32com.client
import os

def ppt2pdf(filename, output_filename):
    """
    PPT文件导出为pdf格式
    :param filename: PPT文件的文件路径
    :param output_filename: 导出的pdf文件的路径
    :return:
    """
    # 2). 打开PPT程序
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    # ppt_app.Visible = True  # 程序操作应用程序的过程是否可视化
    ppt = ppt_app.Presentations.Open(filename)

    # 4). 打开的PPT另存为pdf文件。17数字是ppt转图片，32数字是ppt转pdf。
    ppt.SaveAs(output_filename, 32)
    print("导出成pdf格式成功!!!")
    # 退出PPT程序
    ppt_app.Quit()

def get_files(file_dir):
    file_name_ppt_list = []
    for filepath,dirnames,filenames in os.walk(file_dir):
        for filename in filenames:
            print(os.path.join(filepath,filename))
            file_name = os.path.join(filepath,filename)
            file_name_ppt_list.append(file_name)
    return file_name_ppt_list


'''
    输入文件夹路径，
    该python程序可将该文件夹下的所有PPT文件转变成对应的PDF文件。注意是所有的PPT文件
'''

if __name__ == '__main__':
    dirpath = 'D:\\100000000-计网助教\\slides-2021-4比3'
    file_name_ppt_path = get_files(dirpath)
    for filename in file_name_ppt_path: # 遍历文件夹下的所有文件
        # 判断文件的类型，对所有的ppt文件进行处理(ppt文件以ppt或者pptx结尾的)
        if filename.endswith('ppt') or filename.endswith('pptx'):   # 需要转换的PPT
            print("正在转换PPT文件：" + filename)           # PPT素材1.pptx -> PPT素材1.pdf
            # 将filename以.进行分割，和“PDF”拼接获取存储的PDF路径，我们和PPT存在相同的地方
            filename_list = filename.split('.')
            base = '_'.join(filename_list[0:-1])
            output_filename = base + '.pdf'         # PPT素材1.pdf
            # 将ppt转成pdf文件
            ppt2pdf(filename, output_filename)
