import re
import openpyxl

# ===================================================
# == Implement Excel functions you need below here ==
# ===================================================

from math import atan2, sqrt as SQRT, degrees as DEGREES, cos as COS, sin as SIN, pow as POWER

ABS = abs

def ATAN2(x, y):
    return atan2(y, x)

def IF(stmt, if_true, if_false):
    return if_true if stmt else if_false

class _xlfn:
    @staticmethod
    def XOR(left, right):
        return (left or right) and not (left and right)

def AND(left, right):
    return left and right

def INDEX(arr, y, x):
    new = arr[x-1]
    return new[y-1]

def MATCH(target, items, _):
    i = 0
    for sub in items:
        for item in sub:
            if item == target:
                return i + 1
            i += 1
    
    return "#VALUE!"
    
def SWITCH(target, *options):
    for k, v in zip(options[::2], options[1::2]):
        if k == target:
            return v
    return "#VALUE!"

def globs():
    return {
        k: v for k, v in globals().items()
        if k in ("SQRT", "ATAN2", "IF", "_xlfn", 
                 "AND", "INDEX", "MATCH", "SWITCH",
                 "DEGREES", "COS", "SIN", "POWER", "ABS")
        # Add your functions here to expose them to Excel
    }

# ===============================================================
# == DO NOT EDIT BELOW THIS UNLESS YOU KNOW WHAT YOU ARE DOING ==
# ===============================================================
    
class Cell:
    def __init__(self, cell, parent):
        self.pos = cell.coordinate
        self.value = cell.value
        self.parent = parent

    def evaluate(self):
        if isinstance(self.value, str) and self.value.startswith("="):
            new_val = self.value[1:]
            
            for sub_cell in re.finditer("[A-Z]+[0-9]+:[A-Z]+[0-9]+", new_val):
                name = sub_cell.group(0)
                new_val = new_val.replace(name, repr([c for c in self.parent.replace_range(name)]))

            for sub_cell in re.finditer("[A-Z]+[0-9]+", new_val):
                name = sub_cell.group(0)
                if name not in globals():
                    new =  repr(self.parent[name].evaluate())
                    new_val = new_val.replace(name, new)
            
            new_val = re.sub("([^=])=([^=])", r"\1==\2", new_val)
            res = eval(new_val, globs())
            # print(self.value)
            # print(f"[{self.pos}] ={new_val} |=> {res}")
            return res
        return self.value

class Workbook:
    def __init__(self, cells):
        self.cn = {}
        for col in cells:
            for cell in col:
                self.new_cell(cell)

    def new_cell(self, cell):
        self.cn[cell.coordinate] = Cell(cell, self)
    
    def __getitem__(self, item):
        return self.cn[item]
    
    def __setitem__(self, key, value):
        self.cn[key].value = value

    def replace_range(self, range_):
        match = re.match(r"(?P<chars_start>[A-Z]+)(?P<ints_start>[0-9]+):(?P<chars_end>[A-Z]+)(?P<ints_end>[0-9]+)", range_)
        start_c = match.group("chars_start")
        end_c = match.group("chars_end")
        start_i = int(match.group("ints_start"))
        end_i = int(match.group("ints_end"))
        
        cells = []
        
        for c in range(ord(start_c), ord(end_c)+1):
            row = []
            for i in range(start_i, end_i+1):
                cell = chr(c) + str(i)
                row.append(self[cell].evaluate())
            cells.append(row)
        return cells
    
    def eval_stmt(self, stmt):
        return Cell(stmt, self).evaluate()

# Entry point
def load_file(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.worksheets[0]
    cells = sheet[sheet.dimensions]
    return Workbook(cells)
