#!/usr/bin/env python
#coding=utf-8
#-------------------------------------------------------------------------------
# Name:        模块1
# Purpose:
#
# Author:      Administrator
#
# Created:     05-07-2017
# Copyright:   (c) Administrator 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from PyQt4 import QtCore,QtGui
from PyQt4.QtGui import QFileDialog
from read_xls2 import *
import ui_format_xls,sys
out_dir='.'
import sys, os,re
sys.path.append(os.getcwd())  # 把当前路径（即程序所在路径）暂时加入系统的path变量中

class format_xls(QtGui.QDial,ui_format_xls.Ui_Dialog):
    def __init__(self, parent=None):
        super(format_xls, self).__init__(parent)
        self.setupUi(self)

    @QtCore.pyqtSignature("")
    def on_pushButton_clicked(self):
        global out_dir
        fileName=QFileDialog.getOpenFileName(self,
                                    "文件选择",
                                    "./",
                                    "Xls Files (*.xls)")
        ILILST=re.findall('[0-9]{18}',self.textEdit.toPlainText()) \
        if len(re.findall('[0-9]{18}',self.textEdit.toPlainText()))>0  \
        else list()
##        print(ILILST)
         #======过滤已有报关单号======#
        listadd=list()
        for item_ in loadlxs(fileName):
            if item_[2] in ILILST:
                continue
            listadd.append(item_)
##        print(type(listadd),':',listadd)
        if fileName!='':
            out_file='out_'+fileName.split(r'/')[-1]
            writexls(listadd,out_dir+'/'+out_file)
##            writexls(loadlxs(fileName),out_dir+'/'+out_file)
            print(out_dir+'/'+out_file)

        os.startfile(out_dir+'/'+out_file)

    @QtCore.pyqtSignature("")
    def on_pushButton_2_clicked(self):
        global out_dir
        out_dir=QFileDialog.getExistingDirectory(self,"选择目录","./")
        if out_dir !='':
            print(out_dir)


def main():
    app=QtGui.QApplication(sys.argv)
    form = format_xls()
    form.show()
    app.exec_()


if __name__ == '__main__':
    main()
