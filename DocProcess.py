import sys
import shutil
import os


from DocProcessFunc import RawDataInit,generateResult

from colorama import init
init(autoreset=True)


if __name__ == '__main__':


    #TODO 颜色列表扩展重用，不要出现颜色列表超出索引

    workBookPaths = os.listdir('RawData\\')

    #TODO 判断文件的内容构成一个函数,判断一些错误情况,并且reverseworkBookPath.
    #Datainit
    workBookPaths,nameList, weight_name=RawDataInit(workBookPaths)


    # 预处理每个月份的表格
    workbook = []



    #TODO 处理过程函数化 def generateResult
    generateResult(workBookPaths,0,nameList,weight_name)


    # 生成有权结果,判断是否要生成有权结果
    ifweight=''
    while ifweight!='y' and ifweight!='n':
        # print(ifweight!='y')
        ifweight = input('Do you need to generate result with weight?[y/n]')

    if ifweight == 'y':
        generateResult(workBookPaths,ifweight,nameList,weight_name)

    print('Process Finished.lzx.')
    print('Your result are in \\result folder.')
    print('Deleting cache.')
    shutil.rmtree('cache')
    input("“Press Any Key to close terminal.”")
    sys.exit()