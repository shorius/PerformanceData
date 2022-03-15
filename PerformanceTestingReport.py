# !/usr/bin/python
# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
import pandas as pd
import numpy as np
import plotly
import plotly.graph_objects as go
import plotly.figure_factory as ff
import plotly.express as px
from plotly.subplots import make_subplots
import time
import datetime


def plotrun(text):
    '''

    :param text: 檔案名稱
    :return:
    '''
    # fp = 'PD_20210413_10_43_47.xlsx'
    fp = text

    ##### Cut the main table from excel, option -> skiprows is variable
    data = pd.read_excel(fp, skiprows=11, sheet_name=None)
    BS = data.get('all').fillna(0)

    ##### Cut Driveri table from excel
    data1 = pd.read_excel(fp, sheet_name=0, header=None)
    ALL = data1.fillna(0)
    driver = ALL.iloc[0, [1, 2]].tolist()

    ##### graph data
    total_values = BS.iloc[0:-4]
    total_values['Time_second'] = pd.to_datetime(total_values['time'], unit='ms')
    total_values['absTimeINT'] = pd.to_numeric(total_values['absTime'], downcast='integer')
    total_values['Time_format'] = pd.to_datetime(total_values['absTimeINT'] / 1000, unit='s').dt.strftime(
        '%Y-%m-%d %H:%M:%S')

    ##### Replace BigJank data trans 0 to NaN
    seriesBigJank = total_values['BigJank']
    rep_BigJank = seriesBigJank.replace({0: None})

    if max(total_values[['BigJank', 'Jank']].max()) == 0:
        jankrange = 1
    else:
        jankrange = max(total_values[['BigJank', 'Jank']].max()) + 1

    ##### Subplots Setting
    fig = make_subplots(
        rows=3, cols=2,
        shared_xaxes=True,
        vertical_spacing=0.08,
        row_heights=[20, 20, 60],
        specs=[[{"type": "table", "colspan": 2}, None],
               [{"type": "table", "colspan": 2}, None],
               [{"colspan": 2, "secondary_y": True}, None]],
        subplot_titles=("DeviceInfo", "Summary", "FPS"))

    # fig = go.Figure()

    ##################### Add DeviceInfo table data#####################
    # DeviceInfo_values = pd.read_excel(fp, sheet_name=0, header=3, nrows=1, usecols="A:P", squeeze=True)
    # mobile = {'LYA-L29': 'HUAWEI Mate 20 Pro', 'SM-A7050': 'SAMSUNG Galaxy A70', 'CPH1907': 'OPPO Reno 2'}
    #
    # # DeviceValue = DeviceInfo_values.columns.values[:-3]
    # if DeviceInfo_values['CPU Type'].str.contains('Apple').bool() == False:
    #     DeviceValue = DeviceInfo_values.keys()[:-3].tolist()
    #     DN = DeviceInfo_values.iloc[0, 1]
    #     DName = mobile[DN]
    # else:
    #     DeviceValue = DeviceInfo_values.keys()[:-8].tolist()
    #     DName = DeviceInfo_values.iloc[0, 1]

    DeviceInfo_values, DeviceValue, DName, interFrame = deviceIF(fp)

    fig.add_trace(
        go.Table(
            name='DeviceInfo',
            header=dict(
                values=['<b>' + str(i) + '</b>' for i in DeviceValue],
                font=dict(size=15),
                # fill=dict(color=['#527BBC']),
                align=['center', 'center'],
                height=35,
            ),
            cells=dict(
                values=[DeviceInfo_values[i] for i in DeviceValue],
                align=['center', 'center'],
                font=dict(size=13),
                height=45
            )
        ),
        row=1, col=1
    )

    avg_values = BS.iloc[-2::]

    choice_data = ['Num', 'FPS', 'Jank', 'AppCPU[%]', 'Memory[MB]', 'TotalCPU[%]']
    avg_data = avg_values[['Num', 'FPS', 'Jank', 'AppCPU[%]', 'Memory[MB]', 'TotalCPU[%]']]

    fig.add_trace(
        go.Table(
            name='Summary',
            header=dict(
                values=['<b>' + str(i) + '</b>' for i in choice_data],
                # line_color='darkslategray',
                font=dict(size=15),
                height=28,
                align=['center', 'center']
            ),
            cells=dict(
                values=[avg_data[i] for i in choice_data],
                # line_color='darkslategray',
                font=dict(size=15, ),
                height=30,
                align=['center', 'center']
            )
        ),
        row=2, col=1
    )

    ##### Add Bar chart ###

    d = dict(total_values['label'].value_counts()[total_values['label'].unique()])
    basicColor = ['#F47262', '#6a67ce', '#49c0b6', '#fd9f3e', '#5F77BE', '#f29556', '#489cd4', '#8c88cd']
    colors = basicColor * 5

    if interFrame == True:
        xrange = max(total_values[['FPS', 'Stutter[%]', 'TotalCPU[%]']].max())
    else:
        xrange = max(total_values[['FPS', 'TotalCPU[%]']].max())
    # print(xrange)

    for index, (k, v) in enumerate(d.items()):
        fig.add_trace(go.Bar(
            name=k,
            x=[total_values['Time_second'][v]],
            y=[int(xrange) + 5, int(xrange) + 10],
            orientation='h',
            textposition="inside",
            insidetextanchor="middle",
            texttemplate=k,
            textfont=dict(size=13, color='white'),
            showlegend=False,
            width=(int(xrange) + 10) * 0.09,
            marker=dict(color=colors[index]),
            hovertemplate="End at %{x}",
        ),
            secondary_y=False,
            row=3, col=1
        )

    ##### Add Sactter data

    fig.add_trace(
        go.Scatter(x=total_values['Time_second'], y=total_values['TotalCPU[%]'],
                   marker=dict(color='#CCCC99'),
                   mode='lines',
                   name='TotalCPU',
                   fill='tozeroy',
                   opacity=0.4,
                   hovertemplate=" <b>%{y}%<br>",
                   ),
        secondary_y=False,
        row=3, col=1
    )

    fig.add_trace(
        go.Scatter(x=total_values['Time_second'], y=total_values['Jank'],
                   mode='lines+markers',
                   marker=dict(size=5, color='#d2691e'),
                   name='Jank',
                   opacity=0.7,
                   hovertemplate="<b>%{y}<br>",
                   ),
        secondary_y=True,
        row=3, col=1
    )

    fig.add_trace(
        go.Scatter(x=total_values['Time_second'], y=rep_BigJank,
                   mode='lines+markers',
                   marker=dict(size=7, color='#EE6363'),
                   name='BigJank',
                   fill='toself',
                   marker_symbol='circle-open',
                   marker_line_width=2,
                   hovertemplate="<b>%{y}<br>",
                   ),
        secondary_y=True,
        row=3, col=1
    )

    fig.add_trace(
        go.Scatter(x=total_values['Time_second'], y=total_values['FPS'],
                   marker=dict(color='#4682B4'),
                   name='FPS',
                   hovertemplate="<b>%{y}<br>",
                   #                xaxis='x1', yaxis='y1'
                   ),
        secondary_y=False,
        row=3, col=1
    )

    if interFrame == True:
        fig.add_trace(
            go.Scatter(
                x=total_values['Time_second'],
                y=total_values['InterFrame'],
                marker=dict(color='#56a0d3'),
                name='InterFrame',
                hovertemplate="<b>%{y}<br>",
            ),
            secondary_y=False,
            row=3, col=1
        )
    else:
        pass

    ####### Create scatter trace of text labels

    fig.update_layout(
        width=1900,
        height=1100,
        autosize=False,
        hovermode='x unified',
        xaxis_title="Time",
        yaxis_title="<b>FPS</b>",
        legend=dict(
            orientation="v",
            y=0.4,
            x=1.0
        ),
        yaxis={'domain': [0, .45], 'anchor': 'y1'},
        barmode='stack'
    )

    fig.update_xaxes(
        rangeselector=dict(
            buttons=list([
                dict(count=10,
                     label="10s",
                     step="second",
                     stepmode="backward"),
                dict(count=30,
                     label="30s",
                     step="second",
                     stepmode="backward"),
                dict(count=1,
                     label="1m",
                     step="minute",
                     stepmode="backward"),
                dict(step="all")
            ]),
            # x=0,
            y=0.53,
        ),
        calendar='taiwan',
        type='date',
        tick0=total_values['Time_second'].values[0],
        tickmode='auto',
        dtick=86400000.0,
        rangeslider=dict(
            visible=True,
            thickness=0.06,
            borderwidth=0,
        ),
        tickformatstops=[
            dict(dtickrange=[None, 1000], value="%M:%S"),
            dict(dtickrange=[1000, 10000], value="%M:%S"),
            dict(dtickrange=[10000, 60000], value="%M:%S"),
            dict(dtickrange=[total_values['Time_second'].values[0], total_values['Time_second'].values[-1]],
                 value="%M:%S"),
        ]
    )

    fig.update_annotations(font=dict(size=20))
    fig.add_annotation(text="TotalCPU  Usage [%]",
                       xref="paper",
                       yref="paper",
                       font=dict(size=13, color='#808000'),
                       x=0.45, y=0.1,
                       showarrow=False)

    fig.update_yaxes(title_text="<b>Jank</b>", range=[0, jankrange], secondary_y=True)
    fig.update_layout(
        title=go.layout.Title(
            text="<b>Performance Testing Report _ %s </b><br> %s / %s " % (DName, driver[0], driver[1]),
            font=go.layout.title.Font(size=23)))

    config = dict({
        'modeBarButtonsToRemove': ['toggleSpikelines', 'hoverCompareCartesian'],
        'displaylogo': False,
        'scrollZoom': True,
        'displayModeBar': False,
        'editable': False
    })

    # fig.write_html('TestingReport.html')
    plotly.offline.plot(fig, config=config, auto_open=False, filename='TestReport.html')


def deviceIF(fp):
    '''
    既有設備資訊
    :param fp:
    :return:
    '''
    DeviceInfo_values = pd.read_excel(fp, sheet_name=0, header=3, nrows=1, usecols="A:P", squeeze=True)
    mobile = {'LYA-L29': 'HUAWEI Mate 20 Pro', 'SM-A7050': 'SAMSUNG Galaxy A70', 'CPH1907': 'OPPO Reno 2'}
    # DeviceValue = DeviceInfo_values.columns.values[:-3]
    interFrame = False
    if DeviceInfo_values['CPU Type'].str.contains('Apple').bool() == False:
        DeviceValue = DeviceInfo_values.keys()[:-3].tolist()
        DN = DeviceInfo_values.iloc[0, 1]
        DName = mobile[DN]
        interFrame = True
        if DeviceInfo_values['CPU Type'].str.contains('sm').bool() == True:
            interFrame = False
    else:
        DeviceValue = DeviceInfo_values.keys()[:-8].tolist()
        DName = DeviceInfo_values.iloc[0, 1]

    return DeviceInfo_values, DeviceValue, DName, interFrame


def changeName(fp):
    '''
    # Get The Device Name for renewing the report's name
    :param fp: 最終檔案
    :return:
    '''

    global finalReportName
    # renewName = str(deviceIF(fp)[2]) + "_PerformanceTest.html"
    renewName = str(deviceIF(fp)[2]) + '_TestReport_.html'
    originalName = 'TestReport.html'
    reportExists = os.path.exists(renewName)
    now = time.strftime("%Y%m%d_%H_%M_%S")

    # 如果資料夾不存在就建立
    if not os.path.isdir('.\Report'):
        os.mkdir('.\Report')

    # 檢查 Report資料夾內是否有存在相同檔名
    for root, dirs, filenames in os.walk('Report'):
        print('Report資料夾下的檔案', filenames)
        if len(filenames) > 0:
            for names in filenames:
                print(names)
                if names == renewName:
                    finalReportName = renewName.replace('.html', str(now) + '.html')  # Rename
                    break
                else:
                    finalReportName = renewName
        else:
            finalReportName = renewName
        finalReportPath = os.path.join(os.getcwd(), root, finalReportName)

    shutil.copyfile(os.path.join(os.getcwd(), originalName), os.path.join(os.getcwd(), 'Report', finalReportName))
    print(originalName, "copied as", finalReportName)  # 輸出提示
    print('最終路徑', finalReportPath)


    return finalReportPath, finalReportName

