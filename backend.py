import pandas as pd
import calendar
import format_report

def to_df(file_name):
    
    header_names = [
        'Operational Date', 'Time', 'Event #', 'Category Code', 'Schedule Arrival', 'Actual Arrival', 'Delay Code', 'Primary Incident', 'Train', 
        'Consist', 'Car', 'Route', 'Location', 'Dep.', 'Arr.', 'Status', 'blank', 'Remarks', 'Remarks Added By', 'Event Creation Time']

    return pd.read_excel(file_name, skiprows = 3, header = None,  names = header_names)

def break_down_date(date_string):
    year = date_string[:4]
    month = int(date_string[5:7])
    full_month_name = calendar.month_name[month]
    day = date_string[-2:]

    return year, full_month_name, day


def filter_and_sort( working_df, operational_day, BT_delay_codes, QQ_delay_codes):

    # sort report by time
    working_df = working_df.sort_values(by='Time')

    # check if any delays carryover 
    carryover = len(working_df.query('Time < @operational_day'))

    # if so, move them to the bottom
    if carryover > 0 :
        # store carryovers
        carryover_df = working_df.query('Time < @operational_day')
        # remove carryovers
        working_df = working_df[working_df['Time'] >= operational_day]
        # add carryovers to bottom
        working_df = pd.concat([working_df, carryover_df], ignore_index=True)
        
    
    # seperate BT and QQ delays
    return working_df[working_df['Delay Code'].isin(BT_delay_codes)], \
            working_df[working_df['Delay Code'].isin(QQ_delay_codes)]


def process_report(downloaded_file):

    operational_day = '03:00'
    

    BT_delay_codes = ['BTCC', 'BTCT', 'BTDV', 'BTFP', 'BTKP', 'BTKQ', 'BTMN', 'BTOP', 'BTRV', 'BTSF',
                    'BTUP', 'BTWA', 'BTWC', 'BTWM', 'BTWR'] 

    QQ_delay_codes = ['QQAC', 'QQAF', 'QQDF', 'QQEF', 'QQEX', 'QQFI', 'QQHB', 'QQLA', 'QQME',
                    'QQMQ', 'QQOL', 'QQPF', 'QQRF', 'QQSF', 'QQTE', 'QQWC', 'QQWN', 'QQWR',
                    'QQWS']
    
    original_df = to_df(downloaded_file)


    # remove the one blank column
    original_df = original_df.drop('blank', axis=1)

    # filter for BT and QQ codes
    Alstom_delays_df = original_df[original_df['Delay Code'].isin(BT_delay_codes + QQ_delay_codes)]

    # Check if report spans multiple days
    dates_list = list(set(Alstom_delays_df['Operational Date'].tolist()))
    dates_list.sort()

    file_names = []
    dates_without_delays = []

    # go through each date in report
    for date in dates_list:
        
        working_df = Alstom_delays_df.query('`Operational Date` == @date')

        year, full_month_name, day = break_down_date(date)
        
        # seperate, and sort into BT and QQ
        BT_df, QQ_df = filter_and_sort(working_df, operational_day, BT_delay_codes, QQ_delay_codes)
        
        # save and format
        if len(BT_df) > 0: 
            filename = f'Atlas - L102 - Delay to Train Details {full_month_name} {day} BT.xlsx'
            BT_df.to_excel(filename, index=False)
            format_report.format(filename)
            file_names.append(filename)
        else:
            dates_without_delays.append (f'No BT delays for {full_month_name} {day}')

        if len(QQ_df) > 0:     
            filename = f'Atlas - L102 - Delay to Train Details {full_month_name} {day} QQ.xlsx'
            QQ_df.to_excel(filename, index=False)
            format_report.format(filename)
            file_names.append(filename)
        else:
            dates_without_delays.append (f'No QQ delays for {full_month_name} {day}')

    
    return file_names, dates_without_delays