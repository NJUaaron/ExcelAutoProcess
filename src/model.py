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
    def __init__(self):
        self.sheetNames = None  # sheet name array
        self.dfs = None         # dict
    
    def read(self, filepath):
        self.dfs = pd.read_excel(filepath, sheet_name=None, header=None)
        self.sheetNames = self.dfs.keys()
        return self.sheetNames

    
    def buildDict(self, headerName, option):
        # Build dictionary based on source file dataframe and key column header name. 
        # Use key column as dict's key, other columns as dict's value
        # option: which sheet should be processed
        _dict = {}
        if (option == 'All sheets'):
            # Read all sheets
            for df in self.dfs.values():
                self.buildDictWithSingleSheet(headerName, df, _dict)
        else:
            # Read only one sheet
            df = self.dfs[option]
            self.buildDictWithSingleSheet(headerName, df, _dict)
        return _dict

    
    def buildDictWithSingleSheet(self, headerName, df, _dict):
        # Build dictionary based on dataframe and key column header name. 
        # Use key column as dict's key, other columns as dict's value

        location = findLocation(df, headerName)
        if not location:
            return
        (headerRow, keyColumn) = location
        (row, col) = df.shape

        for i in range(headerRow + 1, row):
            one_person_dict = {}
            for j in range(0, col):
                if j != keyColumn:
                    info_type = df.iloc[headerRow][j]
                    one_person_dict[info_type] = df.iloc[i][j]
            name = df.iloc[i][keyColumn]
            _dict[name] = one_person_dict


class TargetFile:
    def __init__(self):
        self.sheetNames = None  # sheet name array
        self.dfs = None         # dict
        self.saveFilepath = None
    
    def read(self, filepath):
        # Read all sheets in target file as dictionary (sheet name, content df)
        self.dfs = pd.read_excel(filepath, sheet_name=None, header=None)

        # Add .auto in front of .xls
        self.saveFilepath = re.sub('.xls', ".auto.xls", filepath) # Save file path
        # Only output xlsx file, instead of xls file
        if self.saveFilepath[-4:] == '.xls':
            self.saveFilepath = self.saveFilepath + 'x'

        self.sheetNames = self.dfs.keys()
        return self.sheetNames

    def fillIn(self, _dict, headerName, option):
        # Use information in dict to fill in target file
        if (option == 'All sheets'):
            # Write all sheets
            for df in self.dfs.values():
                self.fillInSingleSheet(df, _dict, headerName)
        else:
            # Write only one sheet
            df = self.dfs[option]
            self.fillInSingleSheet(df, _dict, headerName)

    def fillInSingleSheet(self, df, _dict, headerName):
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
        self.sf = SourceFile()  #source file object
        self.tf = TargetFile()  #target file object
        self.view = viewInstance  # View instance
    
    def readSheetName(self, fileName, type):    # return all the sheets name in file as a set
        if type == 1:
            return self.sf.read(fileName)
        else:
            return self.tf.read(fileName)


    def process(self, headerName, selectedSheet1, selectedSheet2):  # Main functionality
        # Check source and target data
        if not self.sf.dfs:
            self.view.debugPrint(3, 'FAILED due to the lack of source file data')
            return
        if not self.tf.dfs:
            self.view.debugPrint(3, 'FAILED due to the lack of target file data')
            return


        # Convert source data to dictionary

        _dict = self.sf.buildDict(headerName, selectedSheet1)
        if not _dict:
            # _dict is empty
            self.view.debugPrint(2, 'Key column header name is not found in source file')
            self.view.debugPrint(3, 'FAILED when convert source data to dictionary')
            return

            
        # Write data to target dataframe 
        try:
            self.tf.fillIn(_dict, headerName, selectedSheet2)
        except Exception as e:
            self.view.debugPrint(2, str(e))
            self.view.debugPrint(3, 'FAILED when write data to target dataframe')
            return

        # Save data frame to target file
        try:
            self.tf.save()
        except Exception as e:
            self.view.debugPrint(2, str(e))
            self.view.debugPrint(3, 'FAILED during writing target file')
            return

        self.view.debugPrint(1, 'Process FINISH!!')