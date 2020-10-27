# -*- coding: utf-8 -*-

'''
@Time    : 2020/9/5 11:02
@Author  : fenglei
@FileName: dayin1.py
@Software: PyCharm

'''
import win32com.client
import pythoncom
from tqdm import tqdm
from tkinter import ttk
import tkinter as tk  # 使用Tkinter前需要先导入
from pyautocad import Autocad, APoint
import os
import math
from tkinter.filedialog import askdirectory
import pandas as pd

window = tk.Tk()  # 第1步，实例化object，建立窗口window
# 第2步，给窗口的可视化起名字
window.title('自动打印V2.0     -------桥二所-XXX')
# 第3步，设定窗口的大小(长 * 宽)
window.geometry('600x400')  # 这里的乘是小x
# 第4步，在图形界面上设定标签
l = tk.Label(window, text='你好！请先打开AutoCad2014程序', bg='pink', font=('黑体', 12), width=75,
             height=2).place(x=0, y=12)
Plt_huituyi = 'DWG To PDF.pc3'
prog_id = ['AutoCAD.Application.19.1', 'AutoCAD.Application.22']


def APoint(x, y):
    """坐标点转化为浮点数"""
    # 需要两个点的坐标
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))


def vtpnt(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtfloat(lst):
    """列表转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def vtint(lst):
    """列表转化为整数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, lst)


def vtvariant(lst):
    """列表转化为变体"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, lst)


def button_click():
    def get_zuo_lst(kuang_ceng="图框", line_fou=False):
        global shuchu
        acad = win32com.client.Dispatch("AutoCAD.Application.19")
        doc = acad.ActiveDocument
        try:
            doc.SelectionSets.Item("SS1").Delete()
            doc.SelectionSets.Item("SS2").Delete()
        except:
            print("Delete selection failed")
        slt = doc.SelectionSets.Add("SS1")
        slt1 = doc.SelectionSets.Add("SS2")
        df = pd.DataFrame(columns=['zuo_x', 'zuo_y', 'you_x', 'you_y', "chang", "space"])
        df1 = pd.DataFrame(columns=['x', 'y', 'l_duan'])
        i = 0
        j = 0
        for space in range(2):
            inder = 1
            if auto_tukuang:
                filterType = [0,  67]
                filterData = ["LWPOLYLINE",  space]
            else:
                filterType = [0, 8, 67]  # 定义过滤类型
                filterData = ["LWPOLYLINE", kuang_ceng, space]  # 设置过滤参数
            filterType = vtint(filterType)  # 数据类型转化
            filterData = vtvariant(filterData)  # 数据类型转化
            slt.Select(5, 0, 0, filterType, filterData)  # 实现过滤
            for sl in slt:
                if sl.Length > 60:
                    point_4 = sl.Coordinates
                    if len(list(set(point_4[::2]))) == 2 or len(list(set(point_4[1::2]))) == 2:
                        point_zuo_x = min(point_4[::2])
                        point_zuo_y = min(point_4[1::2])
                        point_you_x = max(point_4[::2])
                        point_you_y = max(point_4[1::2])
                        l_duanbian = round(point_you_y - point_zuo_y)  # 找到图框短边

                        su_scale = [0.2, 0.1, 0.5, 1 / 3, 0.25, 1 / 6, 1 / 8, 1, 2, 3, 4, 5, 6, 8, 10]
                        a1 = math.floor((l_duanbian + 2) / 297)
                        a2 = math.ceil((l_duanbian - 2) / 297)
                        if l_duanbian / 297 in su_scale or a1 == a2:
                            if l_duanbian > 0:
                                chang = (point_you_x - point_zuo_x) * 297 / l_duanbian
                                if chang / 297 > 1.4:
                                    new = pd.DataFrame({'zuo_x': point_zuo_x,
                                                        'zuo_y': point_zuo_y,
                                                        'you_x': point_you_x,
                                                        'you_y': point_you_y,
                                                        'chang': chang,
                                                        'space': space},
                                                       index=[i])
                                    df = df.append(new)
                                    # shuchu = '找到' + sp[space] +'第%d个多段线图框——%.0f个端点，长边长度%.2f\n' % (
                                    #     inder, len(sl.Coordinates) / 2, chang)
                                    # text_out.insert(tk.END, shuchu)
                                    # text_out.update()
                                    # text_out.see(tk.END)
                                    inder = inder + 1
                                    i = i + 1
            if line_fou:
                if auto_tukuang:
                    filterType1 = [0, 67]
                    filterData1 = ["LINE", space]
                else:
                    filterType1 = [0, 8, 67]
                    filterData1 = ["LINE", kuang_ceng, space]
                filterType1 = vtint(filterType1)  # 数据类型转化
                filterData1 = vtvariant(filterData1)  # 数据类型转化
                slt1.Select(5, 0, 0, filterType1, filterData1)  # 实现过滤
                for line in slt1:
                    if line.Length > 60:
                        su_scale = [0.2, 0.1, 0.5, 1 / 3, 0.25, 1 / 6, 1 / 8, 1, 2, 3, 4, 5, 6, 8, 10]
                        if round(line.Length, 3) / 297.00 in su_scale:
                            out = line.Length
                            x_start = line.StartPoint[0]
                            x_end = line.EndPoint[0]
                            y_start = line.StartPoint[1]
                            y_end = line.EndPoint[1]
                            mid_x = (x_start + x_end) / 2
                            mid_y = (y_start + y_end) / 2
                            new1 = pd.DataFrame({'x': mid_x,
                                                 'y': mid_y,
                                                 'l_duan': out},
                                                index=[j])
                            df1 = df1.append(new1)
                            j = j + 1
                df1 = df1.drop_duplicates().sort_values(by=['y', 'x'], ascending=[False, True]).reset_index(drop=True)
                mid_zuo = df1[df1.index % 2 == 0].reset_index(drop=True)
                mid_you = df1[df1.index % 2 == 1].reset_index(drop=True)
                for k in range(len(mid_zuo)):
                    new3 = pd.DataFrame({'zuo_x': mid_zuo['x'][k],
                                         'zuo_y': mid_zuo['y'][k] - mid_zuo['l_duan'][k] / 2,
                                         'you_x': mid_you['x'][k],
                                         'you_y': mid_you['y'][k] + mid_you['l_duan'][k] / 2,
                                         'chang': (mid_you['x'][k] - mid_zuo['x'][k]) * 297 / mid_you['l_duan'][k],
                                         'space': space},
                                        index=[i + k])
                    df = df.append(new3)
                    # shuchu = '找到第%d个线段图框——长边长度%.2f\n' % (inder, new3['chang'])
                    # text_out.insert(tk.END, shuchu)
                    # text_out.update()
                    # text_out.see(tk.END)
                    inder = inder + 1
        df = df.drop_duplicates(['zuo_x']).sort_values(by=['you_y', 'zuo_x'], ascending=[False, False]).reset_index(drop=True)
        shuchu = '当前DWG在模型和布局中共找到%d张图\n' % (len(df))
        text_out.insert(tk.END, shuchu)
        text_out.update()
        text_out.see(tk.END)
        return df['zuo_x'], df['zuo_y'], df['you_x'], df['you_y'], df['chang'], df['space']

    def print_cad(kind, x1, y1, x2, y2, l_length, space, path, Fname, plt_huituyi, Scale=1,
                  paper_for_pdf='UserDefinedMetric (320.00 x 440.00毫米)',
                  paper_for_plt='UserDefinedMetric (440.00 x 320.00毫米)', wunao=False):

        def paper_real_name(l_orpaname, defaultnumber):
            defaultnumber = defaultnumber + 10
            li = [float(l_orpaname[i].split('x ')[1].split('.00')[0]) for i in range(len(l_orpaname))]
            for i in range(len(li) - 1):
                for j in range(len(li) - 1 - i):
                    if li[j] > li[j + 1]:
                        li[j], li[j + 1] = li[j + 1], li[j]
                        l_orpaname[j], l_orpaname[j + 1] = l_orpaname[j + 1], l_orpaname[j]
            li_min = abs(li[0] - defaultnumber)
            index = 0
            for i in range(1, len(li) - 1):
                li_min2 = abs(li[i] - defaultnumber)
                if li_min2 < li_min:
                    li_min = li_min2
                    index = i
            if li[index] < defaultnumber:
                index = index + 1
            return l_orpaname[index]

        a = acaddoc.ActiveSpace
        name = acaddoc.layouts.item(space).Name
        layout = acaddoc.layouts.item(name)  # 先来个layout对象
        plot = acaddoc.Plot  # 再来个plot对象
        acaddoc.SetVariable('BACKGROUNDPLOT', 0)  # 前台打印
        layout.StyleSheet = style  # 选择打印样式
        layout.PlotWithLineweights = True  # 打印线宽
        if 'PDF' in kind:
            name = 'DWG To PDF.pc3'
            layout.ConfigName = name  # 选择打印机
            if wunao:
                paper_l1_size = paper_real_name(zdy_pdf, l_length)
                layout.CanonicalMediaName = paper_l1_size
            else:
                layout.CanonicalMediaName = paper_for_pdf  # 'UserDefinedMetric (320.00 x 440.00毫米)'
            layout.PlotRotation = 1  # 纵向打印
        elif 'PLT' in kind:
            layout.ConfigName = plt_huituyi
            if wunao:
                paper_l1_size = paper_real_name(zdy_plt, l_length)
                layout.CanonicalMediaName = paper_l1_size
            else:
                layout.CanonicalMediaName = paper_for_pdf  # 'UserDefinedMetric (320.00 x 440.00毫米)'
            layout.PlotRotation = 1  # 纵向打印
        layout.PaperUnits = 1  # 图纸单位，1为毫米
        layout.StandardScale = 0  # 图纸打印比例
        layout.PlotWithPlotStyles = True  # 依照样式打印
        layout.PlotHidden = False  # 隐藏图纸空间对象
        target = acaddoc.Viewports.Item(0).Target
        x_pian = target[0]
        y_pian = target[1]
        po1 = APoint(x1 * Scale - x_pian, y1 * Scale - y_pian)
        po2 = APoint(x2 * Scale - x_pian, y2 * Scale - y_pian)
        layout.SetWindowToPlot(po1, po2)
        layout.PlotType = 3.5  # 按照窗口打印，别问我为什么是3.5我试出来的。
        layout.CenterPlot = True
        Fname = name + Fname
        plot.PlotToFile(path + '/' + Fname)

    global shuchu
    global name_kuang
    global xian_fou
    dwg_filaname = os.listdir(dwg_real_path)
    for f in dwg_filaname:
        if f[-4:] == '.dwg':
            acad = win32com.client.Dispatch("AutoCAD.Application.19.1")
            try:
                acad.ActiveDocument.Application.Documents.Open(dwg_real_path + '/' + f)
                acaddoc = acad.ActiveDocument
                # acaddoc.Utility.Prompt("Hello AutoCAD\n")
                if tukuang_kind.get() == 1:
                    xian_fou = True
                else:
                    xian_fou = False
                zx_x, zx_y, ys_x, ys_y, l_length, df_space= get_zuo_lst(kuang_ceng=name_kuang, line_fou=xian_fou)
                cadfile_name = acaddoc.Name.split('.dw')[0] + '-'
                fname = [cadfile_name + str(i) for i in range(1, len(zx_x) + 1)]
                kinds = []
                if pdf_kind.get() == 1:
                    kinds.append('PDF')
                if pdf_kind.get() == 0:
                    if 'PDF' in kinds:
                        kinds.remove('PDF')
                if plt_kind.get() == 1:
                    kinds.append('PLT')
                if plt_kind.get() == 0:
                    if 'PLT' in kinds:
                        kinds.remove('PLT')
                for kind in kinds:
                    path = "./自动打印输出"
                    path = os.path.abspath(path)
                    if not os.path.exists(path):  # 判断是否存在文件夹如果不存在则创建为文件夹
                        os.makedirs(path)
                    acaddoc.Regen
                    for i in tqdm(range(len(fname))):
                        print_cad(kind, zx_x[i], zx_y[i], ys_x[i], ys_y[i], l_length[i], df_space[i], path, fname[i],
                                  Plt_huituyi, Scale,
                                  paper_for_pdf,
                                  paper_for_plt, wunao)
                    shuchu = '--*--' + cadfile_name[:25] + '--*--打印共' + str(len(fname)) + '张\n'
                    text_out.insert(tk.END, shuchu)
                    text_out.update()
                    text_out.see(tk.END)
                try:
                    acad.ActiveDocument.Close()
                except:
                    shuchu = '注意：无法关闭保存该DWG文件\n'
                    text_out.insert(tk.END, shuchu)
                    text_out.update()
                    text_out.see(tk.END)
            except:
                print('nice work')

    shuchu = '*****************************\n*                           *\n' \
             '*       ！打印成功！        *\n*                           *\n*****************************'
    text_out.insert(tk.END, shuchu)
    text_out.see(tk.END)


def get_paper_list(kand):
    global Plt_huituyi
    global shuchu
    acad = win32com.client.Dispatch("AutoCAD.Application.19.1")
    acaddoc = acad.ActiveDocument
    acaddoc.Utility.Prompt("get paper list\n")
    layout1 = acaddoc.layouts.item('Model')
    acaddoc.SetVariable('BACKGROUNDPLOT', 0)
    layout1.StyleSheet = style
    layout1.PlotWithLineweights = True
    if 'PDF' in kand:
        name = 'DWG To PDF.pc3'
        layout1.ConfigName = name
        paper_names = layout1.GetCanonicalMediaNames()
        shuchu = '本机PDF绘图仪为' + name + '!\n'
        text_out.insert(tk.END, shuchu)
    elif 'PLT' in kand:
        try:
            Plt_huituyi = 'DesignJet 430 C4714A FENG.pc3'
            layout1.ConfigName = Plt_huituyi
        except:
            plt_names = list(layout1.GetPlotDeviceNames())
            for plt_name in plt_names:
                if len(plt_name.split(' ')) > 2:
                    mid_plt = plt_name.split(' ')[1]
                    if mid_plt == '430' or mid_plt == '750C':
                        layout1.ConfigName = plt_name
                        Plt_huituyi = plt_name
                        break
        shuchu = '本机PLT绘图仪为' + Plt_huituyi + '!\n'
        text_out.insert(tk.END, shuchu)
        paper_names = layout1.GetCanonicalMediaNames()
    return paper_names


def get_style_sheet():
    acad2 = win32com.client.Dispatch("AutoCAD.Application.19.1")
    acaddoc2 = acad2.ActiveDocument
    acaddoc2.Utility.Prompt("get style sheets\n")
    layout2 = acaddoc2.layouts.item('Model')
    style_sheets = layout2.GetPlotStyleTableNames()
    k = ['acad.ctb']
    for f1 in style_sheets:
        if f1[-4:] == '.ctb' and f1 not in k:
            k.append(f1)
    return k


# def button1_click():
#     kand = ['PDF']
#     f = get_paper_list(kand)
#     comboxlist["values"] = f
#     print(f)


def selectPath():
    global dwg_real_path  # name_kuang
    path_ = askdirectory()
    dwg_path.set(path_)
    dwg_real_path = dwg_path.get()


def button_wunao_click():
    text_out.delete(0.0, tk.END)
    global zdy_pdf
    global zdy_plt
    global wunao
    global shuchu
    shuchu = '自动选纸激活成功，请等待程序运行！\n'
    text_out.insert(tk.END, shuchu)
    wunao = True
    pdf1 = get_paper_list(['PDF'])
    for i in range(len(pdf1)):
        if len(pdf1[i].split(' (')) > 1:
            if pdf1[i].split(' (')[0] == 'UserDefinedMetric' and float(pdf1[i].split(' (')[1].split(' ')[0]) >= 310:
                zdy_pdf.append(pdf1[i])
    plt2 = get_paper_list(['PLT'])
    for i in range(len(plt2)):
        if len(plt2[i].split(' (')) > 1:
            if plt2[i].split(' (')[0] == 'UserDefinedMetric' and float(plt2[i].split(' (')[1].split(' ')[0]) >= 310:
                zdy_plt.append(plt2[i])


def button2_click():
    # kand = ['PLT']
    f = get_style_sheet()   # f = get_paper_list(kand)
    comboxlist_plt["values"] = f
    comboxlist_plt.current(0)


def selectukuang(*args):
    global shuchu
    global auto_tukuang
    global name_kuang  # name_kuang
    name_kuang = l2.get()
    if name_kuang == '':
        shuchu = '*********注意********：\n选定的图框为空，请重新输入图框所在图层名！！！\n'
        text_out.insert(tk.END, shuchu)
        text_out.update()
    else:
        auto_tukuang = False
        shuchu = '选定的图框为-------' + name_kuang + '\n'
        text_out.insert(tk.END, shuchu)
        text_out.update()


sp = ['模型界面', '布局界面']
zdy_pdf = []
zdy_plt = []
auto_tukuang = True
dwg_path = tk.StringVar()
dwg_real_path = ''
xian_fou = True
name_kuang = '图框'
style = 'acad.ctb'
wunao = False
Scale = 1
pdf_kind = tk.IntVar()
Entyr_dwg = tk.Entry(window, textvariable=dwg_path, width=65).place(x=45, y=108)
button_dwg = tk.Button(window, text="路径选择",command=selectPath).place(x=490, y=103)
ckbutton1 = tk.Checkbutton(window, text='PDF', variable=pdf_kind, onvalue=1, offvalue=0).place(x=40, y=150)

tukuang = tk.StringVar
l2 = tk.Entry(window, textvariable=tukuang, font=('黑体', 12), width=15)
l2.place(x=360, y=65)
l3 = tk.Label(window, text='请输入图框所在图层名，如0或者图框', font=('黑体', 8), width=33)
l3.place(x=140, y=65)
button_tukuang = tk.Button(window, text="确认图层", command=selectukuang).place(x=490, y=60)
tukuang_kind = tk.IntVar()
ckbutton3 = tk.Checkbutton(window, text='线段图框', variable=tukuang_kind, onvalue=1, offvalue=0).place(x=40, y=60)

plt_kind = tk.IntVar()
ckbutton2 = tk.Checkbutton(window, text='PLT', variable=plt_kind, onvalue=1, offvalue=0)
ckbutton2.select()
ckbutton2.place(x=100, y=150)
# button1 = tk.Button(window, text="PDF图纸列表", font=('黑体', 8), width=10, height=1, command=button1_click)
# button1.place(x=100, y=135)
button2 = tk.Button(window, text="样式选择", command=button2_click)
button2.place(x=490, y=145)
button = tk.Button(window, text="打印所选", background="green", font=('黑体', 20), width=8, height=1, command=button_click)
button.place(x=460, y=315)
button3 = tk.Button(window, text="自动选纸", background="yellow", font=('黑体', 20), width=8, height=1,
                    command=button_wunao_click)
button3.place(x=460, y=205)
paper_for_pdf = 'UserDefinedMetric (310.00 x 430.00毫米)'
paper_for_plt = 'UserDefinedMetric (310.00 x 430.00毫米)'
scr = tk.Scrollbar(window)
scr.place(x=437, y=205)
shuchu = tk.StringVar
text_out = tk.Text(window, background="Ivory", width=60, height=12)
text_out.place(x=15, y=205)
scr.config(command=text_out.yview())

# def go_pdf(*args):
#     print(comboxlist.get(), type(comboxlist.get()))
#     global paper_for_pdf
#     paper_for_pdf = comboxlist.get()


def go_plt(*args):
    global style
    style = comboxlist_plt.get()


# comvalue = tk.StringVar()
# comboxlist = ttk.Combobox(window, textvariable=comvalue, width=50)
# comboxlist.bind("<<ComboboxSelected>>", go_pdf)
# comboxlist.place(x=180, y=135)

comvalue_plt = tk.StringVar()
comboxlist_plt = ttk.Combobox(window, textvariable=comvalue_plt, width=40)
comboxlist_plt.bind("<<ComboboxSelected>>", go_plt)
comboxlist_plt.place(x=180, y=150)
window.mainloop()
