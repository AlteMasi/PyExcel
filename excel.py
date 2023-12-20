import openpyxl, os
from openpyxl import Workbook
from typing import Union

class Cell:
    def __init__(self, x:Union[str,int] = 1, y:int = 1):
        self.x = x
        self.y = y

class ExcelFile:
    def __init__(self, name:str="excelfile", path:str=None) -> None:
        """ 
            Create an Excel File.

            Example:
            ```
            file = ExcelFile(name=filename)
            file.write(Cell("A",1),value,False)
            print(file.read(Cell(1,1)))
            file.save() 
            ```
        """
        self.name = name
        self.__extension = ".xlsx"
        self.path = path if path else f"{self.name}{self.__extension}"
        self.open()
    
    def open(self):
        """ Open Workbook """
        # try to load workbook
        self.wb = "workbook"
        try: self.wb = openpyxl.load_workbook(self.path)
        # create a new workbook the current one doesn't exist
        except FileNotFoundError: self.wb = Workbook()
        self.ws = self.wb.active
    
    def save(self):
        """ Save Workbook """
        try: self.wb.save(self.path)
        except PermissionError: raise PermissionError("can't save the file because it's already open")

    def change_sheet(self, name):
        """ change worksheet """
        self.ws = self.wb[name]

    def create_sheet(self, name):
        """ create a worksheet """
        self.ws = self.wb.create_sheet(name)
        self.save()

    def change_sheet_name(self, name):
        """ change a worksheet name """
        self.ws.title = name
        self.save()

    def get_letter(self, number):
        """ 
            return a letter from a number
            ```
            file = ExcelFile(name=filename)
            for i in range(100):
                print(file.get_letter(i)) 
            ```
        """
        delta = 64
        if number > 90 - delta: 
            prefix = chr(int(number/26)+delta)
            if ord(prefix) > 90: return None
            return prefix + chr(number%26+delta+1)
        return chr(number+delta)
    
    def write(self, cell:Union[Cell,tuple], value:str, save:bool=True):
        """ write value on a cell object """
        # convert cell to Cell
        cell : Cell = self.get_cell(cell)
        # if x is number -> convert to letter
        if type(cell.x) is int: cell.x = self.get_letter(cell.x)
        # write value on the cell
        self.ws[f"{cell.x}{cell.y}"] = value
        # if save is True then save Workbook
        if save: self.save()

    def read(self, cell:Union[Cell,tuple])->str:
        """ read value of a cell object """
        # convert cell to Cell
        cell : Cell = self.get_cell(cell)
        # if x is number -> convert to letter
        if type(cell.x) is int: cell.x = self.get_letter(cell.x)
        # read value of a cell
        return self.ws[f"{cell.x}{cell.y}"].value
    
    def get_cell(self, cell:Union[Cell,tuple])->Cell:
        if type(cell) in [tuple,list]:
            return Cell(cell[0],cell[1])
        elif type(cell) is Cell:
            return cell
        else:
            raise TypeError(f"{type(cell)} should be a Cell, a tuple or a list")
