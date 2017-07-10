#!/usr/bin/env python
#coding=utf-8
#-------------------------------------------------------------------------------
# Name:        模块1
# Purpose:
#
# Author:      Administrator
#
# Created:     07-07-2017
# Copyright:   (c) Administrator 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from distutils.core import setup
import py2exe
import sys,os

sys.path.append(os.getcwd())  # 把当前路径（即程序所在路径）暂时加入系统的path变量中

#this allows to run it with a simple double click.
sys.argv.append('py2exe')

py2exe_options = {
        "includes": ["sip"],  # 如果打包文件中有PyQt代码，则这句为必须添加的
        "dll_excludes": ["MSVCP90.dll"],  # 这句必须有，不然打包后的程序运行时会报找不到MSVCP90.dll，如果打包过程中找不到这个文件，请安装相应的库
        "compressed": 1,
        "optimize": 2,
        "ascii": 0,
        "bundle_files": 0,  # 关于这个参数请看第三部分中的问题(2)
        }

setup(
      name = 'format_xls',
      version = '1.0',
      windows = [{'script':'format_xls.py'}],   # 括号中更改为你要打包的代码文件名
      zipfile = None,
      options = {'py2exe': py2exe_options}
      )
