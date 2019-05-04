# -*- coding: utf-8 -*-
"""
Created on Sat May  4 15:12:57 2019

@author: willi
"""

def update_pres_text(week_num, dataframes, prs):
    
    import datetime
    now = datetime.datetime.now()
    
    df1 = dataframes[0]
    df2 = dataframes[1]
    df3 = dataframes[2]
    df4 = dataframes[3]
    
    # create the text variables
    X1 = week_num  # assign the report week
    X2 = df1[df1['Week'] == week_num]['Hours booked'].values[0] # hours booked for specific week
    X3 = df1[df1['Week'] == week_num]['Hours Expected'].values[0]  # hours expected for specific week
    X4 = df1[df1['Week'] == week_num]['Lag/Lead'].values[0]  # YTD target lag/lead hours
    X5 = (df1[df1['Week'] == week_num]['Hours/Resource'].values[0] / 40) * 100  # resource utilization
    X6 = str(now.year)
    X7 = "NA"
    X8 = "NA"
    X9 = "NA"
    df2['util'] = df2.sum(axis = 1)  # create max utilization column
    X10 = df2[df2['util'] == max(df2['util'])].index.values[0]  # name of project with max 3 week utilization
    X11 = df2[df2['util'] == min(df2['util'])].index.values[0]  # name of project with min 3 week utilization
    X12 = df2.loc[X10][2]/df2.loc[X10][3] * 100  # print the change in utiliztion for the max project
    X13 = df2.loc[X11][2]/df2.loc[X11][3] * 100  # print the change in utiliztion for the max project
    X14 = df3[df3['Growth'] == max(df3['Growth'])].index.values[0]  # find department with max growth
    X15 = df3[df3['Growth'] == min(df3['Growth'])].index.values[0]  # find department with min growth
    
    # add a row that is the ratio of dev to maintenance
    X16 = df4['Development '].mean() / df4['Maintenance'].mean()  # average ratio across week in question

    values = {'X1':X1, 'X2':X2, 'X3':X3, 'X4':X4, 'X5':X5, 
              'X6':X6, 'X7':X7, 'X8':X8, 'X9':X9, 'X10':X10, 
              'X11':X11, 'X12':X12, 'X13':X13, 'X14':X14, 
              'X15':X15, 'X16':X16}
    
    for i in range(16, 0, -1):
        key = 'X' + str(i)
        replaceText(key, values[key], prs)
    
def replaceText(key, value, prs):
    
    if type(value) == int:
        formatted_value = str(value)
    elif type(value) != str:
        formatted_value = format(float(value), ',.1f')
    else:
        formatted_value = str(value)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = run.text.replace(key, formatted_value)