import pandas as pd
import re


def findLocation(df, name):
    # Search for name in the dataframe, 
    # and return its location
        (row, col) = df.shape    # 列数 (col number)
        for i in range(row):
            for j in range(col):
                if df.iloc[i][j] == name:
                    return (i, j)
        return None   # Didn't find

class SourceFile:
    def __init__(self, filepath):
        # Default only read the first sheet of source file as dataframe
        self.df = pd.read_excel(filepath, sheet_name=0, header=None)

        # Row number and col number of dataframe
        (self.row, self.col) = self.df.shape
    
    def buildDict(self, headerName):
        # Build dictionary based on source file dataframe and key column head name. 
        # Use key column as dict's key, other columns as dict's value
        location = findLocation(self.df, headerName)
        if location:
            (headerRow, keyColumn) = location
        else:
            raise Exception("FAILED to find key column header in source file")

        _dict = {}
        for i in range(headerRow + 1, self.row):
            one_person_dict = {}
            for j in range(0, self.col):
                if j != keyColumn:
                    info_type = self.df.iloc[headerRow][j]
                    one_person_dict[info_type] = self.df.iloc[i][j]
            name = self.df.iloc[i][keyColumn]
            _dict[name] = one_person_dict
        return _dict


class TargetFile:
    def __init__(self, filepath):
        # Read all sheets in target file as dictionary (sheet name, content df)
        self.dfs = pd.read_excel(filepath, sheet_name=None, header=None)

        # Add .auto infront of .xls
        self.saveFilepath = re.sub('.xls', ".auto.xls", filepath) # Save file path

    def fillIn(self, _dict, headerName):
        # Use information in dict to fill in target file
        for df in self.dfs.values():
            self.fillInSheet(df, _dict, headerName)

    def fillInSheet(self, df, _dict, headerName):
        # Fill in one sheet
        (row, col) = df.shape
        location = findLocation(df, headerName)
        if location:
            (headerRow, keyColumn) = location
        else:
            # Skip this sheet if key column not found
            return

        for i in range(headerRow + 1, row):
            for j in range(0, col):
                if j != keyColumn:
                    name = df.iloc[i][keyColumn]
                    info_type = df.iloc[headerRow][j]
                    # Make sure name and info_type exist in _dict
                    if name in _dict.keys() and info_type in _dict[name].keys():
                        df.iloc[i][j] = _dict[name][info_type]
    
    def save(self):
        # Save to local file
        writer = pd.ExcelWriter(self.saveFilepath)  # pylint: disable=abstract-class-instantiated
        for sheet_name, sheet in self.dfs.items():
            sheet.to_excel(writer, sheet_name=sheet_name, index=False, header=None)
        writer.save()

        

class Model:
    def __init__(self, viewInstance):
        self.sf = None  #source file object
        self.tf = None  #target file object
        self.view = viewInstance  # View instance

    def process(self, source_file_name, target_file_name, headerName):
        # Initialize source file and target file instance
        try:
            self.sf = SourceFile(source_file_name)
            self.tf = TargetFile(target_file_name)
        except Exception as e:
            self.view.debugPrint(2, str(e))
            self.view.debugPrint(3, 'FAILED during reading files')
            return


        # Write to target data frame
        try:
            _dict = self.sf.buildDict(headerName)
            self.tf.fillIn(_dict, headerName)
        except Exception as e:
            self.view.debugPrint(2, str(e))
            self.view.debugPrint(3, 'FAILED during data manipulation')
            return


        # Save data frame to target file
        try:
            self.tf.save()
        except Exception as e:
            self.view.debugPrint(2, str(e))
            self.view.debugPrint(3, 'FAILED during writing target file')
            return

        self.view.debugPrint(1, 'Process FINISH!!')