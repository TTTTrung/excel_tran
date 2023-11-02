import pandas as pd
import glob
import os
import shutil
from datetime import datetime


def main(product_id_mapping):
        try:
                folder_path = input('Pls type a directory name :')

                if os.path.exists('data/' + folder_path):
                        file_paths = glob.glob(os.path.join('data\{}'.format(folder_path), '*.xls'))
                        for fileN,file_path in enumerate(file_paths):
                                df = pd.read_excel(file_path, skiprows=7,header=None)
                        
                                column_names = ['Column1', 'Column2', 'Column3', 'Column4', 'Column5', 'Column6', 'Column7', 'Column8', 'Column9', 'Column10', 'Column11', 'Column12', 'Column13', 'Column14', 'Column15', 'Column16', 'Column17', 'Column18', 'Column19', 'Column20', 'Column21']
                                df.columns = column_names

                                df = df.iloc[:-3]
                                df = df.fillna(0)
                                df[column_names[4:20]] = df[column_names[4:20]].astype(float)

                                pd.set_option('display.max_columns', None)
                                pd.set_option('display.float_format', '{:.2f}'.format)


                        # Replace '.' with empty string in Column1
                                df['Column1'] = df['Column1'].str.replace(r'^\.*', '', regex=True)

                                df = df.drop(columns=['Column21'], axis=1)

                                new_columns = []
                                for col_num in range(4, 20):
                                        new_col_name = f'New_{column_names[col_num]}'
                                        df[new_col_name] = df[column_names[col_num]].shift(-1)

                                df = df.iloc[::2]
                                df['Column2'] = df['Column1'].map(lambda x: product_id_mapping.get(x, {}).get('Costcenter'))
                                df['รายได้อื่นๆ'] = df['Column12']+df['Column13'] +df['Column14']+df['Column16']+df['Column17']
                                df['Column5'] = df['Column5'] + df['New_Column5']
                                df['Column6'] = df['Column6'] + df['New_Column6']
                                df['Column7'] = df['Column7'] + df['New_Column7'] + df['Column8'] + df['New_Column8']
                                df['Column9'] = df['Column9'] + df['New_Column9']
                                df['Column10'] = df['Column10'] + df['New_Column10']

                                df['Column8'] = df['Column5'] + df['Column6'] + df['Column7']+df['Column9']+df['Column10']+df['Column11']+df['Column12']+df['Column13']+df['Column14']+df['Column15']+df['Column16']+df['Column17']+df['Column18']


                                columns_to_drop = ['Column4','New_Column5','New_Column6','New_Column7','New_Column8','New_Column9','New_Column10','New_Column11','New_Column12','New_Column13','New_Column14','New_Column15','New_Column16','New_Column17','New_Column18','New_Column19','New_Column20']
                                df = df.drop(columns=columns_to_drop, axis=1)

                                index=df.columns.get_loc('Column18')
                                df.insert(index, 'Column8', df.pop('Column8'))

                                index=df.columns.get_loc('Column18')
                                df.insert(index, 'รายได้อื่นๆ', df.pop('รายได้อื่นๆ'))

                                df.rename(columns={'Column1': 'ประเภท', 'Column2': 'CC.', 'Column3': 'คน', 'Column5': 'เงินเดือน/ค่าจ้าง','Column6':'ค่าครองชีพ','Column7':'ตำแหน่ง/วิชา','Column8':'รวมรายได้','Column9':'ค่าล่วงเวลา','Column10':'เบี้ยขยัน','Column11':'โบนัส','Column12':'เงินรางวัล','Column13':'ค่าพาหนะ','Column14':'ค่ากะ','Column15':'เงินชดเชย','Column16':'ค่าเบี้ยเลี้ยง','Column17':'รับอื่นๆ','Column18':'ค่าทักษะ','Column19':'กองทุน','Column20':'ประกันสังคม'},inplace=True)

                                # Create a new row with the total people
                                total_row = pd.DataFrame({'CC.': ['รวม'], 'คน': [df['คน'].sum()], 'เงินเดือน/ค่าจ้าง': [df['เงินเดือน/ค่าจ้าง'].sum()], 'ค่าครองชีพ': [df['ค่าครองชีพ'].sum()], 'ตำแหน่ง/วิชา': [df['ตำแหน่ง/วิชา'].sum()], 'รวมรายได้': [df['รวมรายได้'].sum()], 'ค่าล่วงเวลา': [df['ค่าล่วงเวลา'].sum()], 'เบี้ยขยัน': [df['เบี้ยขยัน'].sum()], 'โบนัส': [df['โบนัส'].sum()], 'เงินรางวัล': [df['เงินรางวัล'].sum()], 'ค่าพาหนะ': [df['ค่าพาหนะ'].sum()], 'ค่ากะ': [df['ค่ากะ'].sum()], 'เงินชดเชย': [df['เงินชดเชย'].sum()], 'ค่าเบี้ยเลี้ยง': [df['ค่าเบี้ยเลี้ยง'].sum()], 'รับอื่นๆ': [df['รับอื่นๆ'].sum()], 'ค่าทักษะ': [df['ค่าทักษะ'].sum()], 'รายได้อื่นๆ': [df['รายได้อื่นๆ'].sum()], 'กองทุน': [df['กองทุน'].sum()], 'ประกันสังคม': [df['ประกันสังคม'].sum()]})
                        
                                df = pd.concat([df, total_row], ignore_index=True)

                                # Create a new DataFrame with the total_row
                                df.to_excel('holder/{}.xlsx'.format(fileN),index=False)
                
                        secondOut(folder_path,product_id_mapping)
                else:
                        print("the directory doesn't exist")
                         
        except Exception as e:
                print("An error occurred:", str(e))
def secondOut(name,product_id_mapping) :
        try:    
                global truCheck
                current_datetime = datetime.now()
                formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")  # Use underscores instead of colons
                directory = f'outfinal/{name}_{formatted_datetime}'  # Use forward slash or backslash as needed

                if os.path.exists(directory):
                # If it exists, remove the directory and its contents
                        shutil.rmtree(directory)

                os.makedirs(directory)
                print(f"Directory '{directory}' created or overwritten successfully.")
                folder_path = 'holder'
                file_paths = glob.glob(os.path.join(folder_path, '*.xlsx'))

                column_names_Tru = ['เงินเดือน/ค่าจ้าง', 'ค่าครองชีพ', 'ตำแหน่ง/วิชา', 'ค่าล่วงเวลา', 'เบี้ยขยัน', 'โบนัส', 'รายได้อื่นๆ', 'ค่าทักษะ','เงินชดเชย']

                column_names_TRT_TUC = ['เงินเดือน/ค่าจ้าง', 'ค่าครองชีพ', 'ตำแหน่ง/วิชา', 'ค่าล่วงเวลา', 'เบี้ยขยัน', 'โบนัส', 'รายได้อื่นๆ', 'ค่าทักษะ','เงินชดเชย','กองทุน','ประกันสังคม']

                all_dataframes = [] 
                conca2=[]
                conca3=[]

                for file_path in file_paths:
                        df = pd.read_excel(file_path,)
                        df = df.iloc[:-1]

                        dfs_to_concat = [] # List to hold dataframes for each column

                        if truCheck:

                                for column_name in column_names_Tru:
                                        df1 = df[['ประเภท', 'CC.', column_name]].copy()
                                        df1['ชื่อบัญชี']=column_name
                                        df1['บัญชี'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get(column_name))
                                        df1.rename(columns={column_name: 'จำนวนเงิน'}, inplace=True)
                                        df1['prod.'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('Product'))
                                        index=df1.columns.get_loc('CC.')
                                        df1.insert(index, 'บัญชี', df1.pop('บัญชี'))
                                        index = df1.columns.get_loc('จำนวนเงิน')
                                        df1.insert(index, 'prod.', df1.pop('prod.'))
                                        index = df1.columns.get_loc('ประเภท')
                                        df1.insert(index, 'ชื่อบัญชี', df1.pop('ชื่อบัญชี'))
                                        dfs_to_concat.append(df1)
                        else :
                               for column_name in column_names_TRT_TUC:
                                        df1 = df[['ประเภท', 'CC.', column_name]].copy()
                                        df1['ชื่อบัญชี']=column_name
                                        df1['บัญชี'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get(column_name))
                                        df1.rename(columns={column_name: 'จำนวนเงิน'}, inplace=True)
                                        df1['prod.'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('Product'))
                                        index=df1.columns.get_loc('CC.')
                                        df1.insert(index, 'บัญชี', df1.pop('บัญชี'))
                                        index = df1.columns.get_loc('จำนวนเงิน')
                                        df1.insert(index, 'prod.', df1.pop('prod.'))
                                        index = df1.columns.get_loc('ประเภท')
                                        df1.insert(index, 'ชื่อบัญชี', df1.pop('ชื่อบัญชี'))
                                        dfs_to_concat.append(df1)
                               
                                
                        # Concatenate dataframes for this file and append to all_dataframes
                        result_df = pd.concat(dfs_to_concat)
                        all_dataframes.append(result_df)
                        
                        df2 = df[['ประเภท', 'CC.', 'กองทุน']].copy()
                        df2['ชื่อบัญชี'] = 'กองทุน'
                        df2['บัญชี'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('กองทุน'))
                        df2.rename(columns={'กองทุน': 'จำนวนเงิน'}, inplace=True)
                        df2['prod.'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('Product'))
                        index=df2.columns.get_loc('CC.')
                        df2.insert(index, 'บัญชี', df2.pop('บัญชี'))
                        index = df2.columns.get_loc('จำนวนเงิน')
                        df2.insert(index, 'prod.', df2.pop('prod.'))
                        index = df2.columns.get_loc('ประเภท')
                        df2.insert(index, 'ชื่อบัญชี', df2.pop('ชื่อบัญชี'))
                        conca2.append(df2)

                        df3 = df[['ประเภท', 'CC.', 'ประกันสังคม']].copy()
                        df3['ชื่อบัญชี'] = 'ประกันสังคม'
                        df3['บัญชี'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('ประกันสังคม'))
                        df3.rename(columns={'ประกันสังคม': 'จำนวนเงิน'}, inplace=True)
                        df3['prod.'] = df['ประเภท'].map(lambda x: product_id_mapping.get(x, {}).get('Product'))
                        index=df3.columns.get_loc('CC.')
                        df3.insert(index, 'บัญชี', df3.pop('บัญชี'))
                        index = df3.columns.get_loc('จำนวนเงิน')
                        df3.insert(index, 'prod.', df3.pop('prod.'))
                        index = df3.columns.get_loc('ประเภท')
                        df3.insert(index, 'ชื่อบัญชี', df3.pop('ชื่อบัญชี'))
                        conca3.append(df3)
                        
                soc = pd.concat(conca3, ignore_index=True)   
                pf = pd.concat(conca2 , ignore_index=True) 
                sal = pd.concat(all_dataframes, ignore_index=True)


                if truCheck:
                        template =pd.read_excel('template.xlsx',sheet_name='Sheet1')
                        cce = pd.concat([sal, template], ignore_index=True)
                else:
                        template =pd.read_excel('template.xlsx',sheet_name='Sheet2')
                        cce = pd.concat([sal, template], ignore_index=True)
                        cce = cce.drop(columns='prod.', axis=1)


                cce.to_excel('{}/sal.xlsx'.format(directory),index=False)

                

                if truCheck:
                        pf.to_excel('{}/pf.xlsx'.format(directory),index=False)
                        soc.to_excel('{}/soc.xlsx'.format(directory), index=False)
                        print('1')
                
                merge(directory)
        except Exception as e:
                print("An error occurred:", str(e))

def merge(directory):
        try:
                excel_file_paths = 'holder'
                file_path= glob.glob(os.path.join(excel_file_paths,'*.xlsx'))

                # Initialize an empty list to store DataFrames
                dfs = []

                # Loop through each Excel file
                for excel_file_path in file_path:
                # Read data from the Excel file into a DataFrame
                        df = pd.read_excel(excel_file_path)
                        df = df.drop(columns='รายได้อื่นๆ', axis=1)
                
                # Append the DataFrame to the list
                        dfs.append(df)

                # Concatenate all DataFrames vertically into one DataFrame
                merged_df = pd.concat(dfs, ignore_index=True)

                merged_df.to_excel('{}/รวมเเล้ว.xlsx'.format(directory),index=False)
        except Exception as e:
                print("An error occurred:", str(e))   

def delete_files_in_directory():
        try:
                directory_path="holder"
                files = glob.glob(os.path.join(directory_path, '*'))
                for file in files:
                        if os.path.isfile(file):
                                os.remove(file)
                print("All files deleted successfully.")
        except OSError:
                print("Error occurred while deleting files.")         
if __name__ == "__main__":
        print("Choose an option:")
        print("1. Option 1 Tru")
        print("2. Option 2 Trt")
        print("3. Option 3 Tuc")
        truCheck = False
        choice = input("Enter your choice: ")
    
        if choice == '1':
                from mapping.Tru_mapping import product_id_mapping_Tru
                product_id_mapping = product_id_mapping_Tru
                truCheck = True
                main(product_id_mapping)
        elif choice == '2':
                from mapping.Trt_mapping import product_id_mapping_Trt
                product_id_mapping = product_id_mapping_Trt
                main(product_id_mapping)
        elif choice == '3':
                from mapping.Tuc_mapping import product_id_mapping_Tuc
                product_id_mapping = product_id_mapping_Tuc
                main(product_id_mapping)
        else:
                print("Invalid choice. Please select a valid option.")
        delete_files_in_directory()
        print('=================================================')
        print('credit santhiti malee :)')
 