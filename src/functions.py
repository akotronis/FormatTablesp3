import xlsxwriter

def rgb2hex(tup):
    return '#%02x%02x%02x' % tup

def num2col(tup1, tup2=None, to_range=False): 
    '''Converts two tuples of zero indexed positions to a cell range (string)
    e.g.: num2col((0,0), (1,3)) -> 'A1:D2'
    if cell_2=None (i.e. we have only one tuple), the input is converted to range or single cell
    according to the value of to_range
    e.g.: num2col((0,0), to_range=True)) -> 'A1:A1'
    e.g.: num2col((0,0), to_range=False)) -> 'A1'
    e.g.: num2col((0,0), (1,3)) -> 'A1:D2'
    '''
    func = xlsxwriter.utility.xl_col_to_name
    _from = f'{func(tup1[1])}{tup1[0]+1}'
    if tup2 is not None:
        _to = f'{func(tup2[1])}{tup2[0]+1}'
    else:
        if to_range:
            _to = _from
        else:
            return _from
    return f'{_from}:{_to}'

def col2num(cell_range, to_single=True):
    '''Inverse procedure of num2col
        e.g.: col2num('A1:C3') -> ((0,0), (2,2))
        e.g.: col2num('A1', to_single=True) -> (0,0)
        e.g.: col2num('A1', to_single=False) -> ((0,0),(0,0))
    '''
    if ':' in cell_range:
        cell_1, cell_2 = list(map(xlsxwriter.utility.xl_cell_to_rowcol, cell_range.split(':')))
        return cell_1, cell_2
    else:
        to_num = xlsxwriter.utility.xl_cell_to_rowcol(cell_range)
        return to_num if to_single else (to_num, to_num)

def reorder_range(*args):
    '''Reorder range to be top left bottom right like while defining the same range
        e.g.: A3:D1 -> A1:D3
        e.g.: (2,0), (0,3) -> (0,0), (2,3)
    '''
    if isinstance(args[0], str):
        inp = args[0].upper()
        left, right = inp.split(':')
        if left == right:
            return inp
        left_chars, left_digits = ''.join([c for c in left if not c.isdigit()]), int(''.join([d for d in left if d.isdigit()]))
        right_chars, right_digits = ''.join([c for c in right if not c.isdigit()]), int(''.join([d for d in right if d.isdigit()]))
        l, r = (min(left_chars, right_chars), min(left_digits, right_digits)), (max(left_chars, right_chars), max(left_digits, right_digits))
        return f'{l[0]}{l[1]}:{r[0]}{r[1]}'
    else:
        tup1, tup2 = args
        if tup1 == tup2:
            return args
        return (min(tup1[0],tup2[0]), min(tup1[1],tup2[1])), (max(tup1[0],tup2[0]), max(tup1[1],tup2[1]))

def url_from_sheet_range(sheet_name, _range):
    ''' _range comes in like ((0,1),(2,3))
    '''
    _range = dup(_range)
    return f'internal:{sheet_name}!{num2col(*_range, to_range=True)}'

def has_row_label(row_label, rows, lookupto=None):
    '''lookupto indicates the row until which to search. NOT zero indexed. Starts from 1
        rows is a list of lists where each list starts with the row label ('RW', etc)
        row label is like 'RW'
    '''
    search_in = rows if lookupto is None else rows[:lookupto]
    for row in search_in:
        if row[0] == row_label:
            return True
    return False

def rows_from_label(label, rows, lookupto=None, no_desc=True):
    '''Returns rows from rows that have first element (label)=label as a list of lists,
        WITHOUT the label if no_desc=True, else WITH the label
        rows_from_label('BT'): [['RS', 'x'],['BT'],['BT','y']] -> [[],['BT','y']]
        label can be a string or a list of strings
        lookupto indicates the row until which to search. NOT zero indexed. Starts from 1
    '''
    start = 1 * no_desc
    search_in = rows if lookupto is None else rows[:lookupto]
    if isinstance(label, str):
        return [row[start:] for row in search_in if row[0] == label]
    return [row[start:] for row in search_in if row[0] in label]

def dup(*args):
    '''dup(0,0) -> ((0,0),(0,0))
        dup((0,0)) -> ((0,0),(0,0))
        dup(((0,0),(0,0))) -> ((0,0),(0,0))
        dup(((0,0),(1,1))) -> ((0,0),(1,1))
        dup('A1') -> ((0,0),(0,0))
    '''
    first_arg = args[0]
    if isinstance(first_arg, int):
        return (args, args)
    elif isinstance(first_arg, tuple):
        if isinstance(first_arg[0], tuple):
            return first_arg
        return (first_arg, first_arg)
    elif isinstance(first_arg, str):
        return col2num(first_arg, to_single=False)   

def strip_diff_labels(txt, header=True):
    '''If header=True:
        Remove part of the form (abc) including the parenthesis if
            1. stripped input ends with this part AND
            2. in parenthesis there are at most `remove_up_to_chars` chars AND
            3. if in the parenthesis are any chars, first is not digit.
            4. If we have just '()', it is removed.
        If header=False:
        Remove alphanumeric and whitespace
    '''
    if header:
        remove_up_to_chars = 3
        txt = txt.strip()
        if txt.endswith(')') and '(' in txt:
            last_parenthesis_ind = len(txt) - (txt[::-1].find('(') + 1)
            part_in_parenthesis = txt[last_parenthesis_ind + 1:-1]
            if len(part_in_parenthesis) <= remove_up_to_chars:
                if not part_in_parenthesis or not(part_in_parenthesis[0].isdigit()):
                    return txt[:last_parenthesis_ind]
        return txt
    else:
        return ''.join([c for c in txt if not c.isalnum() and c.strip()])

def format_from_symbol(formats, val, row_type, sheet_name):
    '''row_type comes as 'STATS', 'TOT' or 'MAIN'
       stats=True -> we are at stats sheet
    '''
    if '-' not in val and '+' not in val:
        return
    num3_1 = 301
    num3_1 = num3_1 if not sheet_name == 'stats' else num3_1 + 10
    num3_2, num3_3, num3_4, num3_5, num3_6 = num3_1 + 1, num3_1 + 2, num3_1 + 3, num3_1 + 4, num3_1 + 5
    num4_1, num4_2, num4_3, num4_4, num4_5, num4_6 = num3_1 + 100, num3_2 + 100, num3_3 + 100, num3_4 + 100, num3_5 + 100, num3_6 + 100
    num5_1, num5_2, num5_3, num5_4, num5_5, num5_6 = num4_1 + 100, num4_2 + 100, num4_3 + 100, num4_4 + 100, num4_5 + 100, num4_6 + 100
    d = {'TOT':{
                '-':{3:formats[num3_1], 2:formats[num3_2], 1:formats[num3_3]},
                '+':{1:formats[num3_4], 2:formats[num3_5], 3:formats[num3_6]},
            },
        'STATS':{
            '-':{3:formats[num4_1], 2:formats[num4_2], 1:formats[num4_3]},
            '+':{1:formats[num4_4], 2:formats[num4_5], 3:formats[num4_6]},
        },
        
        'MAIN':{
                '-':{3:formats[num5_1], 2:formats[num5_2], 1:formats[num5_3]},
                '+':{1:formats[num5_4], 2:formats[num5_5], 3:formats[num5_6]},
            },
    }
    if '-' in val:
        return d[row_type]['-'][val.count('-')]
    elif '+' in val:
        return d[row_type]['+'][val.count('+')]