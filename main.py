# -*- coding: utf-8 -*-
"""
Created on Sat May  4 15:13:35 2019

@author: willi
"""

def main(template_name):
    
    # import modules to run report
    from pptx import Presentation
    import pandas as pd
    import numpy as np
    
    # team created modules
    import fill_text
    import email_report
    import create_graphs
    
    # prompt user for week to generate report on
    week_num = int(input('Enter week to analyze: '))
    
    # intitialize presentation object
    prs = Presentation(template_name)    
    
    # open and parse dataset
    xls = pd.ExcelFile('dataset_for_charts.xlsx')
    df1 = xls.parse(0)
    df2 = xls.parse(1, index_col=0)
    df3 = xls.parse(2, index_col=0)
    df4 = xls.parse(3)
    
    # store dataset list for passing to other functions
    dataframes = [df1, df2, df3, df4]
    
    # replace text in template with analyzed text
    fill_text.update_pres_text(week_num, dataframes, prs)
    
    # generate graphs based on the week's data for the report
    create_graphs.create_charts(prs, dataframes)
    
    # save the file and append updated and the week number to it
    save_name = template_name[:template_name.find(".")] + \
    '_for_wk_' + str(week_num) + '.pptx'
    
    # save and close out of the report
    prs.save(save_name)
    print('Report successfully generated for week', str(week_num))
    
    # check if the report should be emailed
    ans = input('Send the report? (y/n): ')
    if ans.lower() == 'y':
        # email the report to stakeholders
        email_report.send_email(week_num, save_name)
    else:
        print('Report will not be emailed. Closing application...')

# start the program
main('weekly_report.pptx')