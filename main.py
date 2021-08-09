import pandas as pd
import os
import shutil
from pathlib import Path
import sys
from openpyxl import load_workbook
from openpyxl import Workbook
import xlsxwriter

class Sprint:
    def __init__(self, repository):
        """
        Initializes the class.
        Args:
            repository (str): local path of the folder where the scripts are located.
        """
        current_path = Path.cwd()
        self.current_path = current_path
        self.repository = os.path.normpath(repository)
        self.Output_folder = os.path.join(current_path,"Output")
        self.path_excel_scripts = os.path.join(current_path,"Files_to_be_searched.xlsx")

        shutil.rmtree(self.Output_folder, ignore_errors=True)

    def add_sheet_excel(self,file_path_excel,df,sheet_name_choosen):
        """
        It inserts a df in a sheet on an excels and creates the excel if not exists.
        Args:
            file_path_excel (str): file path of the excel.
            df (dataframe) : DataFrame
            sheet_name_choosen (str) : name of the sheet
        """
        if len(df.index) == 0:
            pass
        else:
            workbook = xlsxwriter.Workbook(file_path_excel)    
            if os.path.isfile(file_path_excel):
                writer = pd.ExcelWriter(file_path_excel,engine="openpyxl", mode='a')
                df.to_excel(writer, sheet_name=sheet_name_choosen)
                writer.save()
                writer.close()
            else:
                df.to_excel(file_path_excel, sheet_name=sheet_name_choosen)
    
    def dependencies_tree(self):
        """
        It creates an excel with all the dependencies and functions of the scripts in the output folder All_Dependencies_Tree.xlsx.
        """
        paths_excels = self.list_directory(self.Output_folder)
        merge = pd.DataFrame()
        number = 0
        for file in paths_excels:
            if ".xlsx" in file:
                field = f"File_level_{number}"
                length = f"Length_level_{number}"
                number = number + 1
                df = pd.read_excel(file, sheet_name = "Files_length")
                merge = merge.append(df)
                merge = merge.rename(columns={"File":field, "Length": length})
                merge = merge.apply(lambda x: sorted(x, key=pd.isnull)).dropna(how = 'all')
        try:
            merge = merge.drop(["Unnamed: 0"], axis = 1)
        except:
            pass
        merge = merge.fillna("")    
        float_col = merge.select_dtypes(include=['float64'])
        for col in float_col.columns.values:
            merge[col] = merge[col].astype('int64')
        path_excel_scripts = os.path.join(self.Output_folder,"All_Dependencies_Functions_and_Tables.xlsx")
        self.add_sheet_excel(path_excel_scripts,merge, "Dependencies")
        
        merge = pd.DataFrame()
        number = 0
        for file in paths_excels:
            if ".xlsx" in file:
                try:
                    Scripts = f"Scripts_level_{number}"
                    Functions = f"Functions_level_{number}"
                    number = number + 1
                    df = pd.read_excel(file, sheet_name = "Store_procedures")
                    merge = merge.append(df)
                    merge = merge.rename(columns={"Scripts":Scripts, "Functions": Functions})
                    merge = merge.apply(lambda x: sorted(x, key=pd.isnull)).dropna(how = 'all')
                except:
                    pass
        merge = merge.fillna("")
        try:
            merge = merge.drop(["Unnamed: 0"], axis = 1)
        except:
            pass
        path_excel_scripts = os.path.join(self.Output_folder,"All_Dependencies_Functions_and_Tables.xlsx")
        self.add_sheet_excel(path_excel_scripts,merge, "Functions")
        
        
        merge = pd.DataFrame()
        number = 0
        for file in paths_excels:
            if ".xlsx" in file:
                try:
                    Scripts = f"Scripts_level_{number}"
                    Table = f"Table_level_{number}"
                    number = number + 1
                    df = pd.read_excel(file, sheet_name = "Tables")
                    merge = merge.append(df)
                    merge = merge.rename(columns={"Scripts":Scripts, "Table": Functions})
                    merge = merge.apply(lambda x: sorted(x, key=pd.isnull)).dropna(how = 'all')
                except:
                    pass
        merge = merge.fillna("")
        try:
            merge = merge.drop(["Unnamed: 0"], axis = 1)
        except:
            pass
        path_excel_scripts = os.path.join(self.Output_folder,"All_Dependencies_Functions_and_Tables.xlsx")
        self.add_sheet_excel(path_excel_scripts,merge, "Tables")
        
        
        

        merge = pd.DataFrame()
        number = 0
        for file in paths_excels:
            if ".xlsx" in file:
                try:
                    File_not_found = f"File_not_found_level_{number}"
                    number = number + 1
                    df = pd.read_excel(file, sheet_name = "Files_not_found_in_repository")
                    merge = merge.append(df)
                    merge = merge.rename(columns={"Files_not_found_in_repository":File_not_found})
                    merge = merge.apply(lambda x: sorted(x, key=pd.isnull)).dropna(how = 'all')
                except:
                    pass
        merge = merge.fillna("")
        merge = merge.drop_duplicates()
        try:
            merge = merge.drop(["Unnamed: 0"], axis = 1)
        except:
            pass
        self.add_sheet_excel(path_excel_scripts,merge, "Files_not_found_in_repository")
        
        
    def list_directory(self,path):
        """
        It lists the files of a of a given folder no matter its depth.
        Args:
            path (str): local path of the folder where the scripts are located.
        """
        full_path_files = []
        for subdir, dirs, files in os.walk(path):
            try:
                for filename in files:
                    filepath = subdir + os.sep + filename
                    full_path_files.append(filepath)
            except:
                pass
        return full_path_files
        
        
    def copy_files_level(self, level_number):
        """
        Copy files to an Output folder that were stated on the excel Files_to_be_searched.xlsx.
        Args:
            level_number (int): it referes to the depth of the depedency were, for instance, 0 is the files stated on the excel
            and 1 is the dependencies found.
        """
        
        if level_number == 0:
            backup_excel = os.path.join(self.current_path,"Files_to_be_searched_backup.xlsx")
            shutil.copy(self.path_excel_scripts,backup_excel)
        else:
            pass
        
        level_number = str(level_number)
        level_number_excel = os.path.join(self.Output_folder,f"Level_{level_number}_excel.xlsx")
        level_number = os.path.join(self.Output_folder,f"Level_{level_number}_files")
        df = pd.read_excel(self.path_excel_scripts)
        list_scritps = df["Scripts"].str.cat(sep=', ').replace(" ", "")
        list_scritps = list_scritps.split(",")
        list_path = self.list_directory(self.repository)
        os.makedirs(level_number)


        for x in list_scritps:
            for i in list_path:
                file = i.split("\\")[-1]
                if str(x) == str(file):
                    shutil.copy(i,level_number)
                    
        lista = []      
        for path in list_path:
            path = path.split("\\")[-1]
            lista.append(path)
            
        lista_not_found = []
        for file in list_scritps:    
            if file not in lista:
                lista_not_found.append(file)
        mylist_not_found = list(dict.fromkeys(lista_not_found))

        
        df = pd.DataFrame(mylist_not_found, columns = ['Files_not_found_in_repository'])
        self.add_sheet_excel(level_number_excel,df,"Files_not_found_in_repository")


    def excel_dependencies_level(self, level_number):
        """
        It creates an excel with the dependencies of the files that were stated on the excel Files_to_be_searched.xlsx.
        Args:
            level_number (int): it referes to the depth of the depedency were, for instance, 0 is the files stated on the excel
            and 1 is the dependencies found.
        """
 
        level_number = str(level_number)
        level_number_files = os.path.join(self.Output_folder,f"Level_{level_number}_files")
        Scripts_lista = self.list_directory(level_number_files)
        level_number_excel = os.path.join(self.Output_folder,f"Level_{level_number}_excel.xlsx")

        extentions = [".py", ".conf", ".txt", ".param", ".sh", ".csv", ".hql", ".sql", ".lst", ".ls", ".env",".docx",".doc",".xlsx"]
        lista= []
        for file in Scripts_lista:
            try:
                with open(file, "r") as data:
                    data = data.readlines()
                    for line in data:
                        for i in extentions:
                            if i in line:
                                line1 = line.split(i)[0].split("/")[-1]
                                line2 = line.split(i)[0].split(" ")[-1]
                                if len(line1) < len(line2):
                                    line = line1
                                else:
                                    line = line2
                                line = line.split("=")[-1]
                                b = (file + ";;;" + line + i)
                                if line in file or "}" in line or "]" in line or ":" in line or "$" in line or "*" in line:
                                    pass
                                else:
                                    lista.append(b)
            except:
                pass
        if len(lista) == 0 or level_number == "8":
            backup_excel = os.path.join(self.current_path,"Files_to_be_searched_backup.xlsx")
            df = pd.read_excel(backup_excel)
            excel_files = os.path.join(self.current_path,"Files_to_be_searched.xlsx")
            df.to_excel(excel_files)
            os.remove(backup_excel)
            self.excel_length_level(level_number)
            self.excel_stored_procedures(level_number)
            self.dependencies_tree()
            print("The process has finished successfully, check the Output folder for the results")
            sys.exit("NO MORE DEPENDENCIES WERE FOUND")
        else:
            pass
        lista2=[]
        for x in lista:
            a = x.split(";;;")
            lista2.append(a)
        df = pd.DataFrame(lista2, columns = ['File', 'Dependency'])
        df["File"] = df["File"].str.split("\\").str[-1]
        df = df.drop_duplicates()
        df = df.reset_index()
        df = df[['File', 'Dependency']]
        
        file_col_list = df['File'].tolist()
        df = df[~df['Dependency'].isin(file_col_list)]
        self.add_sheet_excel(level_number_excel,df,"Dependencies")
        df = df.rename(columns={"Dependency": "Scripts"})
        df = df[["Scripts"]]
        if len(df.index) == 0:
            pass
        else:
            df.to_excel(self.path_excel_scripts)


    def excel_length_level(self, level_number):
        """
        It creates an excel with the length of the files that were stated on the excel Files_to_be_searched.xlsx.
        Args:
            level_number (int): it referes to the depth of the depedency were, for instance, 0 is the files stated on the excel
            and 1 is the dependencies found.
        """

        level_number = str(level_number)
        level_number_excel = os.path.join(self.Output_folder,f"Level_{level_number}_excel.xlsx")
        
        path_scripts_new = os.path.join(self.Output_folder,f"Level_{level_number}_files")

        Scripts_lista_path = self.list_directory(path_scripts_new)
        lista_len = []
        for file in Scripts_lista_path:
            try:
                with open(file, "r") as f:
                    for i, l in enumerate(f):
                        a = i + 1
                        b = file.split("\\")[-1]
                        d = b + ";;;" + str(a) 
                    lista_len.append(d)
            except:
                pass

        listadf = []
        for x in lista_len:
            a = x.split(";;;")
            listadf.append(a)
        df = pd.DataFrame(listadf, columns = ['File','Length'])
        
        self.add_sheet_excel(level_number_excel,df,"Files_length")



    def excel_stored_procedures(self,level_number):
        """
        It creates an excel with the store procedures found within the files that were stated on the excel Files_to_be_searched.xlsx.
        Args:
            level_number (int): it referes to the depth of the depedency were, for instance, 0 is the files stated on the excel
            and 1 is the dependencies found.
        """

        level_number = str(level_number)
        level_number_files = os.path.join(self.Output_folder,f"Level_{level_number}_files")
        level_number_excel = os.path.join(self.Output_folder,f"Level_{level_number}_excel.xlsx")

        full_path_files = self.list_directory(level_number_files)
        procedure = ".f_"
        lista= []
        for x in full_path_files:
            try:
                with open(x, "r") as data:
                    data = data.readlines()
                    for line in data:
                        if procedure in line and "select " in line:
                            line = line.split("select ")[-1].split("(")[0]
                            x = x.split("\\")[-1]
                            b = (x + ";;;" + line)
                            lista.append(b)
            except:
                pass
        lista2=[]
        for x in lista:
            a = x.split(";;;")
            lista2.append(a)
            
        if len(lista2) == 0:
            pass
        else:
            df = pd.DataFrame(lista2)
            df = df.drop_duplicates()
            df.columns = ["Scripts", "Functions"]     
            self.add_sheet_excel(level_number_excel,df,"Store_procedures")
            
    def tables(self,level_number):
        """
        It creates an excel with the talbess found within the files .
        Args:
            level_number (int): it referes to the depth of the depedency were, for instance, 0 is the files stated on the excel
            and 1 is the dependencies found.
        """
        level_number_files = os.path.join(self.Output_folder,f"Level_{level_number}_files")
        level_number = str(level_number)
        level_number_excel = os.path.join(self.Output_folder,f"Level_{level_number}_excel.xlsx")

        full_path_files = self.list_directory(level_number_files)
        schemas = [" ami."," udw."," udw_stage."," ami_stage."," ami_ops.", " ami_pts."]
        lista= []
        for x in full_path_files:
            try:
                with open(x, "r") as data:
                    data = data.readlines()
                    for line in data:
                        for schema in schemas:
                            if schema in line:
                                line = line.split(schema)[-1]
                                line1 = line.split(" ")[0]
                                line2 = line.split(";")[0]
                                
                                if len(line1) < len(line2):
                                    line = line1
                                else:
                                    line = line2
                                line = line.split("(")[0]
                                line = schema + line
                                line = line.replace(" ","")
                                if ".f_" in line:
                                    pass
                                else:
                                    x = x.split("\\")[-1]
                                    b = (x + ";;;" + line)
                                    lista.append(b)
            except:
                continue

        lista2=[]
        for x in lista:
            a = x.split(";;;")
            lista2.append(a)

        if len(lista2) == 0:
            pass
        else:
            df = pd.DataFrame(lista2)
            df = df.drop_duplicates()
            df.columns = ["Scripts", "Tables"]     
            self.add_sheet_excel(level_number_excel,df,"Tables")


def main(repository):
    """
    It interates over the functions of the object Sprint.
    Args:
        repository (str): local path of the folder where the scripts are located.
    """
    sprint = Sprint(repository)
    for number in range(9):
        sprint.copy_files_level(number)
        sprint.excel_dependencies_level(number)
        sprint.excel_length_level(number)
        sprint.excel_stored_procedures(number)

if __name__ == '__main__':
    print("Introduce the absolute local path of the repository")
    repository = " C:\Users\j.sanguineti.arena\OneDrive - Accenture\scripts\Sprint_Start\Bitbucket"

    #repository = str(input())
    main(repository)