import requests
import logging
import pandas as pd
# pip install -r requirements.txt


if __name__ == "__main__":
    # Setting logging
    logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO)
    logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=logging.CRITICAL)

    logging.info('Accessing country data...')

    # Accessing country data
    all_countries_data = requests.get('https://restcountries.com/v3.1/all').json()

    logging.info('Completed data reception.')

    interest_data = {"Name":[], "Capital":[], "Area":[], "Currencies":[]}

    logging.info('Saving data of interest...')

    # Keeping information of interest to each country
    for item in all_countries_data:
        interest_data["Name"].append(item["name"]["common"])

        logging.info(f'Saving data from: {item["name"]["common"]}')

        # Some countries do not have area, capital or currency values, so it is necessary to process the data
        try:
            interest_data["Capital"].append(item["capital"])
        except:
            interest_data["Capital"].append(["-"])
        try:
            interest_data["Area"].append(item["area"])
        except:
            interest_data["Area"].append("-")
        try:
            interest_data["Currencies"].append(list(item["currencies"].keys()))
        except:
            interest_data["Currencies"].append(["-"])

    logging.info('Data saved successfully.')
    logging.info('Creating XLSX file...')

    # Creating XLSX file using pandas.DataFrame
    dataframe = pd.DataFrame(interest_data).sort_values(by=['Name'])
    dataframe = dataframe.iloc[dataframe['Name'].str.normalize('NFKD').str.encode('ascii', errors='ignore').argsort()]  # 
    dataframe.update(dataframe.applymap(lambda x: ', '.join(x) if isinstance(x, list) else x))  # Replacing lists with strings
    excel = pd.ExcelWriter('Countries_List.xlsx', engine='xlsxwriter')
    dataframe.to_excel(excel, sheet_name='Sheet', index=None, startrow=2, header=False)

    logging.info('File created successfully.')
    logging.info('Editing worksheet...')

    # Editing XLSX file
    workbook = excel.book
    worksheet = excel.sheets['Sheet']

    # Creating text formatting
    title_format = workbook.add_format({
                                        'bold':True,
                                        'font_color':'#4F4F4F',
                                        'font_size':16,
                                        'center_across':True,
                                        'font_name':'Times New Roman',
                                        })
    column_name_format = workbook.add_format({
                                        'bold':True,
                                        'font_color':'#808080',
                                        'font_size':12,
                                        'align':'left',
                                        'font_name':'Times New Roman',
                                        })
    align_format = workbook.add_format({
                                        'align':'left',
                                        'font_name':'Times New Roman',
                                        })
    number_format = workbook.add_format({
                                        'num_format':'#,##0.00',  # Dont work on WebExcel
                                        'font_name':'Times New Roman',
                                        })

    # Applying text formatting
    worksheet.merge_range('A1:D1', 'Countries_List', title_format)
    worksheet.set_column(0, 3, 12, align_format)
    worksheet.set_column(2, 2, 12, number_format)
    for col_num, value in enumerate(dataframe.columns.values):  # Inserting column names with formatting
        worksheet.write(1, col_num, value, column_name_format)

    logging.info('Worksheet successfully edited.')
    logging.info('End.')

    excel.save()  # Finalizing worksheet
