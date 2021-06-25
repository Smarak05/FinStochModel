import os
import pandas as pd 
from pickle import dump, load

directory = "./ProjectData"
companies_map = {}
names = []


def parseSheet(sheet):
    # Remove useless rows at top (row 12)
    name = sheet.iloc[3, 0]
    name = name.split(">")[0].strip()
    column_names = sheet.iloc[11, :]
    column_names = column_names.tolist()
    names.append(name)
    # sheet.columnns = column_names
    
    for idx, _ in enumerate(sheet.columns):
        sheet.rename(columns={sheet.columns[idx]: column_names[idx]}, inplace=True)


    data_dict = {
        'Profitabiliy': [13, 16],
        'Margin Analysis': [19, 29],
        'Asset Turnover': [32, 35],
        'Short Term Liquidity': [38, 44],
        'Long Term Solvency': [47, 61],
        'Growth Over Prior Year': [64, 85],
        'Compound Annual Growth Rate Over Two Years': [88, 109],
        'Compound Annual Growth Rate Over Three Years': [112, 133],
        'Compound Annual Growth Rate Over Five Years': [136, 157]
    }

    for key in data_dict:
        start, end = data_dict[key]
        data = sheet.iloc[start:end+1, :]
        data = data[data['For the Fiscal Period Ending\n'].notnull()]
        data_dict[key] = data

    companies_map[name] = data_dict

def main():
    for filename in os.listdir(directory):
        # print(directory + "/" + filename)
        x = pd.ExcelFile(directory + "/" + filename)
        for sheet in x.sheet_names:
            x1 = pd.read_excel(x, sheet)
            parseSheet(x1)
    
    file_name = 'parsed_data'
    with open (file_name, "wb") as f:
        dump(companies_map, f)

main()



    