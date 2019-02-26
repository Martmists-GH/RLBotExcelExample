import re
import itertools
from tabulate import tabulate
import openpyxl

# ===================================================
# == Implement Excel functions you need below here ==
# ===================================================

from math import atan2, sqrt as SQRT, degrees as DEGREES, cos as COS, sin as SIN, pow as POWER, radians as RADIANS

ABS = abs

def SUM(args):
    args = [arg if isinstance(arg, int) else SUM(arg) for arg in args]
    return sum(args)

def ATAN2(x, y):
    return atan2(y, x)

def IF(stmt, if_true, if_false):
    return if_true if stmt else if_false

class _xlfn:
    @staticmethod
    def XOR(left, right):
        return (left or right) and not (left and right)

def AND(*args):
    return all(args)

def OR(*args):
    return any(args)
    
def INDEX(arr, y, x):
    new = arr[x-1]
    return new[y-1]

def SIGN(x):
    if x == 0:
        return 0
    if x > 0:
        return 1
    if x < 0:
        return -1

def MATCH(target, items, _):
    i = 0
    for sub in items:
        for item in sub:
            if item == target:
                return i + 1
            i += 1
    
    raise Exception(f"Unable to find key {target} in {items}")
    
def SWITCH(target, *options):
    for k, v in zip(options[::2], options[1::2]):
        if k == target:
            return v
    raise Exception(f"Unable to find key {target} in {options[::2]}")

def AVERAGE(*items):
    return sum(items) / len(items)
    
def _phase_inner():
    num = 0
    while True:
        num += 1
        if num > 10:
            num = 0
        yield num
        
iter_ = _phase_inner()
    
def phase(*_):
    return next(iter_)

def globs():
    return {
        k: v for k, v in globals().items()
        if k in ("SQRT", "ATAN2", "IF", "_xlfn", "phase",
                 "AND", "INDEX", "MATCH", "SWITCH", "OR",
                 "DEGREES", "COS", "SIN", "POWER", "ABS",
                 "RADIANS", "AVERAGE", "SIGN", "SUM")
        # Add your functions here to expose them to Excel
    }

# ===============================================================
# == DO NOT EDIT BELOW THIS UNLESS YOU KNOW WHAT YOU ARE DOING ==
# ===============================================================


def col_list(x, y):
    def col_index(char):
        ord_list = [*map(lambda i: ord(i)-65, char)][::-1]
        return sum(26**i + j*26**i for i, j in enumerate(ord_list))-1
    x,y = [*map(col_index, (x, y))]
    cols = itertools.chain(*[itertools.product(map(chr, range(65, 91)), repeat=i) for i in range(1, 8)])
    for i, j in enumerate(cols):
        if x <= i <= y:
            yield "".join(j)
        elif i > y:
            break
            
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
                    new =  repr(self.parent.eval(name))
                    new_val = new_val.replace(name, new)
            
            new_val = re.sub(r"([^\>\<=])=([^\>\<=])", r"\1==\2", new_val)
            try:
                res = eval(new_val, globs())
            except Exception as exc:
                raise Exception(f"Exception in {self.pos}: {exc}")
            
            # print(f"[{self.pos}] ={new_val} |=> {res}")
            return res
        return self.value

class Workbook:
    def __init__(self, cells):
        self.cn = {}
        self.cache = {}
        for col in cells:
            for cell in col:
                self.new_cell(cell)

    def new_cell(self, cell):
        self.cn[cell.coordinate] = Cell(cell, self)
    
    def __getitem__(self, item):
        return self.cn[item]
    
    def __setitem__(self, key, value):
        self.cn[key].value = value
    
    def eval(self, item):
        if item in self.cache:
            return self.cache[item]
        res = self[item].evaluate()
        self.cache[item] = res
        return res
    
    def clear(self):
        self.cache = {}

    def replace_range(self, range_):
        match = re.match(r"(?P<chars_start>[A-Z]+)(?P<ints_start>[0-9]+):(?P<chars_end>[A-Z]+)(?P<ints_end>[0-9]+)", range_)
        start_c = match.group("chars_start")
        end_c = match.group("chars_end")
        start_i = int(match.group("ints_start"))
        end_i = int(match.group("ints_end"))
        
        cells = []
        
        for c in col_list(start_c, end_c):
            row = []
            for i in range(start_i, end_i+1):
                cell = c + str(i)
                row.append(self.eval(cell))
            cells.append(row)
        return cells
    
    def __repr__(self):
        text = ""
        table = []
        for row in range(1, 82):
            r = []
            for col in col_list("A", "AB"):
                r.append(self.eval(col+str(row)) or "")
            table.append(r)
        return tabulate(table)
    
    def eval_stmt(self, stmt):
        # FIXME
        return Cell(stmt, self).evaluate()

# Entry point
def load_file(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.worksheets[0]
    cells = sheet[sheet.dimensions]
    return Workbook(cells)
