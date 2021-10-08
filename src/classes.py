
import csv
import string
import os
import xlsxwriter
from functions import *


class MakeFile():

    def __init__(self, files=[], name='Formatted_Tables', **kwargs):
        # A list of the full paths of the imported files.
        self.file_paths = files
        # The name of the output file (will be like 'Formatted_Tables', 'Formatted_Tables_1', etc). No full path, just the name.
        self.output_name = name
        # The full path of the output file
        self.output_path = self.output_file_name()
        # The instances of class Worksheet the will be made for the tables. NO 'TableOfContents'
        self.wsheets = []
        # The names of the Worksheet instances 'counts', 'percentages', 'stats', NO 'TableOfContents'
        self.wsheetnames = []
        # Pass constants from gui
        MakeFile.DESC_COL = kwargs['DESC_COL']
        MakeFile.ALTERNATE_CLR = kwargs['ALTERNATE_CLR']
        MakeFile.FIRST_COL_WIDTH = kwargs['FIRST_COL_WIDTH']
        MakeFile.OTHER_COL_WIDTH = kwargs['OTHER_COL_WIDTH']
        MakeFile.LINES_BETWEEN_TABLES = kwargs['LINES_BETWEEN_TABLES']
        MakeFile.MIN3_COL = kwargs['MIN3_COL']
        MakeFile.MIN2_COL = kwargs['MIN2_COL']
        MakeFile.MIN1_COL = kwargs['MIN1_COL']
        MakeFile.PLS1_COL = kwargs['PLS1_COL']
        MakeFile.PLS2_COL = kwargs['PLS2_COL']
        MakeFile.PLS3_COL = kwargs['PLS3_COL']
        # Function passed by gui to print output messages
        MakeFile.output = kwargs.get('output')
        
    def output_file_name(self, extension='xlsx'):
        '''- If there exists a file with name 'self.output_name', search for files like self.output_name_1, self.output_name_2, etc
             find maximum _ind if exists and,
           - if exists, new filename=self.output_name_(max_ind+1), 
           - if it doesn't, new filename=self.output_name_1, 
           - If not, new filename=self.output_name
        '''
        try:
            main_file_name = self.output_name
            folder = os.path.split(self.file_paths[0])[0]
            # Get existing folder files of specified extension, without the extension
            existing_files_of_extension_no_ext = [os.path.splitext(fl)[0] for fl in os.listdir(folder) if fl.endswith(f'.{extension}')]

            if main_file_name in existing_files_of_extension_no_ext:
                tbl_files_inds = [int(fl.split('_')[-1]) for fl in existing_files_of_extension_no_ext if main_file_name == '_'.join(fl.split('_')[:-1]) and fl.split('_')[-1].isdigit()]
                if tbl_files_inds:
                    main_file_name = f'{main_file_name}_{max(tbl_files_inds) + 1}'
                else:
                    main_file_name = f'{main_file_name}_1'
            main_file_name = os.path.join(folder, f'{main_file_name}.{extension}')
            return main_file_name
        except:
            if self.output is not None:
                self.output('Error while making output filename.')

    def load_input_files(self):
        '''Returns {'counts':countsInputFile, 'percentages':percentagesInputFile, 'stats':statsInputFile}
           with some of the keys above
        '''
        output = {}
        counts, percentages, stats = [], [], []
        for i, fl in enumerate(self.file_paths):
            input_file = InputFile(fl)
            success = input_file.import_file()
            if success:
                for j, row in enumerate(input_file.rows):
                    if row[0] != 'FT':
                        if row[0] != 'MK':
                            if input_file.is_counts and not input_file.is_percentages:
                                counts.append(row)
                            elif not input_file.is_counts and input_file.is_percentages:
                                percentages.append(row)
                            elif input_file.is_counts and input_file.is_percentages:
                                if row[0] != 'PV':
                                    counts.append(row)
                                    if row[0] != 'RW':
                                        percentages.append(row)
                                else:
                                    new_row = row[:]
                                    new_row[1] = input_file.rows[j-1][1]
                                    percentages.append(new_row)
                        if i == 0 and InputFile.has_diff_lines:
                            if input_file.is_counts and input_file.is_percentages:
                                unwanted_rows_condition_1 = j < len(input_file.rows) - 1 and row[0] in ['PV','RS'] and input_file.rows[j+1][0] == 'MK'
                                unwanted_rows_condition_2 = row[0] == 'RW'
                                if not (unwanted_rows_condition_1 or unwanted_rows_condition_2):
                                    new_row = row[:]
                                    if row[0] == 'MK':
                                        rows_back = 1 if input_file.rows[j-1][0] == 'RS' else 2
                                        new_row[0], new_row[1] = input_file.rows[j-rows_back][0], input_file.rows[j-rows_back][1]
                                    stats.append(new_row) 
                            else:
                                unwanted_rows_condition = j < len(input_file.rows) - 1 and input_file.rows[j+1][0] == 'MK'
                                if not (unwanted_rows_condition):
                                    new_row = row[:]
                                    if row[0] == 'MK':
                                        new_row[0], new_row[1] = input_file.rows[j-1][0], input_file.rows[j-1][1]
                                    stats.append(new_row)
            else:
                if self.output is not None:
                    self.output(f'Error in loading files. Check your input.')
                return
        if counts:
            output['counts'] = InputFile(rows=counts)
        if percentages:
            output['percentages'] = InputFile(rows=percentages)
        if stats:
            output['stats'] = InputFile(rows=stats)
        # Update worksheet names
        self.wsheetnames = output.keys()
        # Output message if worksheet is missing or not.
        if self.output is not None:
            if 'counts' in self.wsheetnames and 'percentages' in self.wsheetnames:
                len_counts = len(output['counts'].rows)
                len_percentages = len(output['percentages'].rows)
                if len_counts != len_percentages:
                    self.output(f'Lines of selected files do not match')
        return output

    def transform_input_files(self, input_files):
        '''Transform input files from the form {'counts':countsInputFile, 'percentages':percentagesInputFile, 'stats':statsInputFile}
           to [
               [('stats',title), ('percentages',title), ('counts',title)],
               [('stats',table_1) ('percentages',table_1), ('counts',table_1)],
               ...
           ] with this order: stats, perc, cnts
           The sheets my be one, two or three.
        '''
        output = []
        sorted_sheetnames = sorted(input_files.keys(), key=lambda x:x[0], reverse=True)
        ws_inpfl = [(ws, input_files[ws]) for ws in sorted_sheetnames]
        for ws_name, input_file in ws_inpfl:
            title, tables = input_file.split_to_parts()
            output.append([(ws_name, title)] + [(ws_name, table) for table in tables])
        return list(zip(*output))

    def add_formats(self, wb):
        toc_format_1 = wb.add_format({'bold': True})
        # TOC Hyperlinks
        toc_format_2 = wb.add_format({'align': 'center', 'font_color': 'blue'})
        toc_format_2.set_underline(1)
        toc_format_2.set_align('vcenter')
        # TOC headers
        toc_format_3 = wb.add_format({'align': 'center', 'bold': True})
        # TOC weighted bases
        toc_format_4 = wb.add_format({'align': 'center', 'num_format': '0'})
        # TOC labels column
        toc_format_5 = wb.add_format({})
        # 0 - item in description column
        format_0 = wb.add_format({'align':'center'})
        format_0.set_align('vcenter')
        # 1 - Row headers - question choices
        format_1 = wb.add_format({'font_size':10,'bg_color':'#D2D2D2','text_wrap':True,'border':2,'align':'left'})
        # format_1.set_left(2)
        format_1.set_align('vcenter')
        # 2 - Row headers (Left) - NO question choices (SAMPLE, TOTAL, MEAN SCORE, SDV etc)
        format_2 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'align':'center','border':2})
        format_2.set_align('vcenter')
        # 3 - IN table - NO question choices (SAMPLE row, ACTUAL SAMPLE row) - NO decimals
        format_3 = wb.add_format({'bold':True,'bg_color':'#E6E6E6','font_size':10,'border':2,'align':'center','num_format':'0'})
        format_3.set_align('vcenter')
        # 4 - IN table - NO question choices (SAMPLE row, TOTAL row, MEAN SCORE row, SDV row etc) - TWO decimals
        format_4 = wb.add_format({'bold':True,'bg_color':'#E6E6E6','font_size':10,'border':2,'align':'center','num_format':'0.00'})
        format_4.set_align('vcenter')
        # 5 - IN table - question choices - NO decimals
        format_5 = wb.add_format({'bg_color':'#f1f1f1','font_size':10,'border':1,'align':'center','num_format':'0'})
        format_5.set_align('vcenter')
        # 5a - IN table - question choices - NO decimals - WHITE COLOR for alternate row color
        format_5a = wb.add_format({'bg_color':'#ffffff','font_size':10,'border':1,'align':'center','num_format':'0'})
        format_5a.set_align('vcenter')
        # 6 - Column Headers - Bottom Line - First Column (TOTAL) - CL row
        format_6 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'border':2,'align':'center'})
        format_6.set_align('vcenter')
        # 7 - Column Headers - Bottom Line - Second to Last Column - CL row
        format_7 = wb.add_format({'bg_color':'#D2D2D2','font_size':10,'border':2,'align':'center','text_wrap':True})
        format_7.set_align('vcenter')
        # 8 - Column Headers - Top Lines
        format_8 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'border':2,'align':'center','text_wrap':True})
        format_8.set_align('vcenter')
        # 9 - Tables sheets title (first line)
        format_9 = wb.add_format({'bold':True,'bg_color':'#f1f1f1','font_size':10,'align':'center','text_wrap':True})
        format_9.set_align('vcenter')
        # 10 - Table description
        format_10 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'align':'center','text_wrap':True})
        format_10.set_align('vcenter')
        # 11 - Column Headers - Top Lines thin left border
        format_11 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'border':2,'left':1,'align':'center','text_wrap':True})
        format_11.set_align('vcenter')
        # 12 - Column Headers - Top Lines thin right border
        format_12 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'border':2,'right':1,'align':'center','text_wrap':True})
        format_12.set_align('vcenter')
        # 13 - Column Headers - Top Lines thin left right border
        format_13 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'border':2,'left':1,'right':1,'align':'center','text_wrap':True})
        format_13.set_align('vcenter')
        # 14 - Table description
        format_14 = wb.add_format({'bold':True,'bg_color':'#D2D2D2','font_size':10,'font_color':'blue','align':'center','text_wrap':True})
        format_14.set_underline(1)
        format_14.set_align('vcenter')
        # 15 - Hypercodes
        format_15 = wb.add_format({'bg_color':'#868686',})
        # 30.. - IN table - NO question choices (SAMPLE row, ACTUAL SAMPLE row) - NO decimals, diffs hyperlinked
        format_301 = wb.add_format({'bold':True,'bg_color':self.MIN3_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_301.set_align('vcenter')
        format_301.set_underline(1)
        format_302 = wb.add_format({'bold':True,'bg_color':self.MIN2_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_302.set_align('vcenter')
        format_302.set_underline(1)
        format_303 = wb.add_format({'bold':True,'bg_color':self.MIN1_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_303.set_align('vcenter')
        format_303.set_underline(1)
        format_304 = wb.add_format({'bold':True,'bg_color':self.PLS1_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_304.set_align('vcenter')
        format_304.set_underline(1)
        format_305 = wb.add_format({'bold':True,'bg_color':self.PLS2_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_305.set_align('vcenter')
        format_305.set_underline(1)
        format_306 = wb.add_format({'bold':True,'bg_color':self.PLS3_COL,'font_size':10,'border':2,'align':'center','num_format':'0','font_color':'blue'})
        format_306.set_align('vcenter')
        format_306.set_underline(1)
        # 40.. - IN table - NO question choices (SAMPLE row, TOTAL row, MEAN SCORE row, SDV row etc) - TWO decimals, diffs hyperlinked
        format_401 = wb.add_format({'bold':True,'bg_color':self.MIN3_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_401.set_align('vcenter')
        format_401.set_underline(1)
        format_402 = wb.add_format({'bold':True,'bg_color':self.MIN2_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_402.set_align('vcenter')
        format_402.set_underline(1)
        format_403 = wb.add_format({'bold':True,'bg_color':self.MIN1_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_403.set_align('vcenter')
        format_403.set_underline(1)
        format_404 = wb.add_format({'bold':True,'bg_color':self.PLS1_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_404.set_align('vcenter')
        format_404.set_underline(1)
        format_405 = wb.add_format({'bold':True,'bg_color':self.PLS2_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_405.set_align('vcenter')
        format_405.set_underline(1)
        format_406 = wb.add_format({'bold':True,'bg_color':self.PLS3_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00','font_color':'blue'})
        format_406.set_align('vcenter')
        format_406.set_underline(1)
        # 50.. - IN table - question choices - NO decimals, diffs hyperlinked
        format_501 = wb.add_format({'bg_color':self.MIN3_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_501.set_align('vcenter')
        format_501.set_underline(1)
        format_502 = wb.add_format({'bg_color':self.MIN2_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_502.set_align('vcenter')
        format_502.set_underline(1)
        format_503 = wb.add_format({'bg_color':self.MIN1_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_503.set_align('vcenter')
        format_503.set_underline(1)
        format_504 = wb.add_format({'bg_color':self.PLS1_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_504.set_align('vcenter')
        format_504.set_underline(1)
        format_505 = wb.add_format({'bg_color':self.PLS2_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_505.set_align('vcenter')
        format_505.set_underline(1)
        format_506 = wb.add_format({'bg_color':self.PLS3_COL,'font_size':10,'border':1,'align':'center','num_format':'0','font_color':'blue'})
        format_506.set_align('vcenter')
        format_506.set_underline(1)
        # 30.. - IN table - NO question choices (SAMPLE row, ACTUAL SAMPLE row) - NO decimals, diffs NO hyperlinked
        format_311 = wb.add_format({'bold':True,'bg_color':self.MIN3_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_311.set_align('vcenter')
        format_312 = wb.add_format({'bold':True,'bg_color':self.MIN2_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_312.set_align('vcenter')
        format_313 = wb.add_format({'bold':True,'bg_color':self.MIN1_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_313.set_align('vcenter')
        format_314 = wb.add_format({'bold':True,'bg_color':self.PLS1_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_314.set_align('vcenter')
        format_315 = wb.add_format({'bold':True,'bg_color':self.PLS2_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_315.set_align('vcenter')
        format_316 = wb.add_format({'bold':True,'bg_color':self.PLS3_COL,'font_size':10,'border':2,'align':'center','num_format':'0',})
        format_316.set_align('vcenter')
        # 40.. - IN table - NO question choices (SAMPLE row, TOTAL row, MEAN SCORE row, SDV row etc) - TWO decimals, diffs NO hyperlinked
        format_411 = wb.add_format({'bold':True,'bg_color':self.MIN3_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_411.set_align('vcenter')
        format_412 = wb.add_format({'bold':True,'bg_color':self.MIN2_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_412.set_align('vcenter')
        format_413 = wb.add_format({'bold':True,'bg_color':self.MIN1_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_413.set_align('vcenter')
        format_414 = wb.add_format({'bold':True,'bg_color':self.PLS1_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_414.set_align('vcenter')
        format_415 = wb.add_format({'bold':True,'bg_color':self.PLS2_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_415.set_align('vcenter')
        format_416 = wb.add_format({'bold':True,'bg_color':self.PLS3_COL,'font_size':10,'border':2,'align':'center','num_format':'0.00',})
        format_416.set_align('vcenter')
        # 50.. - IN table - question choices - NO decimals, diffs NO hyperlinked
        format_511 = wb.add_format({'bg_color':self.MIN3_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_511.set_align('vcenter')
        format_512 = wb.add_format({'bg_color':self.MIN2_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_512.set_align('vcenter')
        format_513 = wb.add_format({'bg_color':self.MIN1_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_513.set_align('vcenter')
        format_514 = wb.add_format({'bg_color':self.PLS1_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_514.set_align('vcenter')
        format_515 = wb.add_format({'bg_color':self.PLS2_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_515.set_align('vcenter')
        format_516 = wb.add_format({'bg_color':self.PLS3_COL,'font_size':10,'border':1,'align':'center','num_format':'0',})
        format_516.set_align('vcenter')
        formats = {
            0:format_0, 1:format_1, 2:format_2, 3:format_3, 4:format_4,
            5:format_5, 6:format_6, 7:format_7, 8:format_8, 9:format_9,
            10:format_10, 11:format_11, 12:format_12, 13:format_13, 14:format_14,
            15:format_15,
            55:format_5a,
            301:format_301, 302:format_302, 303:format_303, 304:format_304, 305:format_305, 306:format_306,
            401:format_401, 402:format_402, 403:format_403, 404:format_404, 405:format_405, 406:format_406,
            501:format_501, 502:format_502, 503:format_503, 504:format_504, 505:format_505, 506:format_506,
            311:format_311, 312:format_312, 313:format_313, 314:format_314, 315:format_315, 316:format_316,
            411:format_411, 412:format_412, 413:format_413, 414:format_414, 415:format_415, 416:format_416,
            511:format_511, 512:format_512, 513:format_513, 514:format_514, 515:format_515, 516:format_516,
            -1:toc_format_1, -2:toc_format_2, -3:toc_format_3, -4:toc_format_4, -5:toc_format_5,
        }
        return formats

    def get_ws(self, name):
        return [myws for myws in self.wsheets if myws.name == name][0]

    def make_content(self):
        # Load files, update sheet names, and transform data
        transformed_input_files = self.transform_input_files(self.load_input_files())
        with xlsxwriter.Workbook(self.output_path, {'strings_to_numbers': True}) as wb:
            # Handle workbook
            formats = self.add_formats(wb)
            # Add 'counts', 'percentages', 'stats' worksheets
            for ws_name in self.wsheetnames:
                ws = wb.add_worksheet(ws_name)
                ws.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
                myws = Worksheet(ws)
                self.wsheets.append(myws)
            # Update the names of the workbook sheets in ToCitem class and the format dict, the row_ind and the rows
            ToCitem.sheets = self.wsheetnames
            ToCitem.formats = formats
            ToCitem.row_ind = 1
            ToCitem.toc_rows = []
            # Write sheet titles for any of 'counts', 'percentages', 'stats'
            title_line = transformed_input_files[0]
            for ws_name, title in title_line:
                myws = self.get_ws(ws_name)
                myws.write_range(Range((0, 1 * self.DESC_COL), f'HRH TABLES\n{title[0]}',formats[9]))
                myws.current_row += 2
            # Initialize table counter
            Table.counter = 0
            # Write worksheet tables and make ToCitems
            table_lines = transformed_input_files[1:]
            for i, line in enumerate(table_lines):
                for ws_name, table in line:
                    myws = self.get_ws(ws_name)
                    table = Table(table, ws_name, first_row=myws.current_row, is_last=i == len(table_lines) - 1)
                    for ranges_name, ranges in table.make_ranges().items():
                        myws.write_ranges(ranges)
                    myws.current_row += table.rows_num + self.LINES_BETWEEN_TABLES
                    myws.last_col = max(myws.last_col, table.cols_num)
                Table.table_colored_cells_dict = {}
                # Make ToCitem
                toc = ToCitem(table)
            # Write Prev Next ranges
            for ws_name in self.wsheetnames:
                myws = self.get_ws(ws_name)
                myws.write_ranges(Table.prev_next_ranges(ws_name))
            # Write ToCitems
            mytocws = Worksheet(wb.add_worksheet('TableOfContents'))
            mytocws.write_ranges(ToCitem.make_ranges())
            # # Update last column to ToC sheet
            mytocws.last_col = ToCitem.last_col
            # Add worksheets settings
            for myws in self.wsheets:
                myws.settings()
            mytocws.settings()
            # Add workbook settings
            wb.worksheets_objs.sort(key=lambda x: x.name)
        
        # Output message if worksheet is missing or not.
        if self.output is not None:
            if 'counts' not in self.wsheetnames and 'percentages' not in self.wsheetnames:
                self.output(f'No table sheets... Check you input')
            elif 'counts' not in self.wsheetnames or 'percentages' not in self.wsheetnames:
                missing = 'counts' if 'counts' not in self.wsheetnames else 'percentages'
                self.output(f"'{missing}' sheet is missing. Was this on purpose?")
            else:
                self.output(f'Tables are ready!')
            
#############################################################################################################
#############################################################################################################

class InputFile():
    has_diff_lines = None

    def __init__(self, filename=None, rows=[]):
        '''filename is expected to be a full path'''
        self.filename = filename
        self.isqps = False
        self.dlm = None
        self.quotchar = None
        # List of lists with the imported file rows
        self.rows = rows
        self.is_counts = None
        self.is_percentages = None

    def isqps_dlm_quotchar(self):
        if not self.filename.endswith('.csv'):
            return
        try:
            with open(self.filename) as f:
                first_line = f.readline().strip()
                last_char = first_line[-1]
                quotchar = last_char if last_char in string.punctuation else ''
                if not 'QPS' in first_line:
                    return
                else:
                    search_str = f'{quotchar}BE{quotchar}'
                    dlm_ind = first_line.find(search_str) + len(search_str)
                    dlm = first_line[dlm_ind]
                    self.isqps, self.dlm, self.quotchar = True, dlm, quotchar
        except:
            return

    def import_file(self):
        self.isqps_dlm_quotchar()
        if all([self.isqps, self.dlm is not None, self.quotchar is not None]):
            try:
                with open(self.filename, newline='') as f:
                    self.rows = list(csv.reader(f, delimiter=self.dlm, quotechar=self.quotchar))
                    InputFile.has_diff_lines = self._has_diff_lines()
                    self.is_counts = self._is_counts()
                    self.is_percentages = self._is_percentages()
                return True
            except:
                return False
        return False

    def _is_counts(self):
        return has_row_label('RW', self.rows)

    def _is_percentages(self):
        return has_row_label('PV', self.rows)

    def has_table_footers(self):
        return has_row_label('FT', self.rows)

    def remove_table_footers(self, rows):
        return [row for row in rows if row[0] != 'FT']

    def _has_diff_lines(self):
        return has_row_label('MK', self.rows, lookupto=20)

    def has_weights(self):
        return has_row_label('RU', self.rows, lookupto=20)

    def remove_diff_lines(self, rows):
        output = [row for row in rows if row[0] != 'MK']
        return output

    def split_to_parts(self):
        '''Input is expected to be the lines of the QPSMR output file as a list of lists
           Returns a tuple (title, tables) where:
           title = a list with the tables title (like ['JF';'TABLES TITLE'])
           tables = a list of tables where table is a list of lists/rows
        '''
        title = rows_from_label('JF', self.rows, lookupto=4)[0]
        tables, table = [], []
        for row in self.rows:
            if row[0] in ['TB', 'EN']:
                if table:
                    table = table[:-1] if table[-1][0] == 'TE' else table
                    tables.append(table)
                table = []
            if row[0] == 'TB' or table:
                table.append(row)
        return title, tables

#############################################################################################################
#############################################################################################################

class Worksheet(MakeFile):
    
    def __init__(self, ws, current_row=0, last_col=0):
        self.ws = ws
        self.name = self.ws.name
        # self.print_description_column = False
        # The current row number of the worksheet to write (zero indexed A1=(0,0))
        self.current_row = current_row
        # The last column number of the worksheet the content occupies (zero indexed A1=(0,0))
        self.last_col = last_col

    def settings(self):
        self.ws.set_zoom(85)
        if self.name == 'TableOfContents':
            self.ws.freeze_panes(1, 0)
            self.ws.set_tab_color('red')
            self.ws.freeze_panes(1, 1)
            self.ws.set_column(0, 0, 15)
            self.ws.set_column(1, self.last_col, 13)
            self.ws.autofilter(f"A1:{chr(ord('@')+self.last_col)}1")
        else:
            self.ws.freeze_panes(1, 1 + self.DESC_COL)
            self.ws.set_column(1 * self.DESC_COL, 1 * self.DESC_COL, self.FIRST_COL_WIDTH)
            self.ws.set_column(1 + self.DESC_COL, self.last_col - 1 + self.DESC_COL, self.OTHER_COL_WIDTH)

    def write(self, *args, **kwargs):
        '''Accepts
               1. args with one string element like 'A1' and value to write as last element
               2. args with one string element like 'A1:B2' and value to write as last element
               3. args with one tuple element like (0,1) and value to write as last element
               4. args with two tuple elements like ((0,1),(2,5)) and value to write as last element
           cell_format=format_object must be passed as keyword argument
           url=url is optional
           In case of hyperlink, text to be displayed must not be passe as string="..." but as positional argument (look 1.-4.)
           Hyperlinks on merged cells won't work
        '''
        value_to_write = args[-1]
        if isinstance(args[0], str):
            if ':' not in args[0]:
                if 'url' in kwargs:
                    self.ws.write_url(args[0], string=value_to_write, **kwargs)
                else:
                    self.ws.write(args[0], value_to_write, kwargs['cell_format'])
            else:
                _from, _to = args[0].split(':')
                if _from != _to:
                    _range = reorder_range(args[0])
                    self.ws.merge_range(_range, value_to_write, cell_format=kwargs['cell_format'])
                else:
                    if 'url' in kwargs:
                        self.ws.write_url(_from, string=value_to_write, **kwargs)
                    else:
                        self.ws.write(_from, value_to_write, kwargs['cell_format'])
        else:
            if (len(args) == 3 and args[0] == args[1]) or len(args) == 2:
                if 'url' in kwargs:
                    self.ws.write_url(*args[0], string=value_to_write, **kwargs)
                else:
                    self.ws.write(*args[0], value_to_write, *kwargs.values())
            else:
                _range = num2col(*reorder_range(*args[:-1]))
                self.ws.merge_range(_range, value_to_write, cell_format=kwargs['cell_format'])

    def write_range(self, _range):
        '''Input _range comes as ((x1,y1),(x2,y2)) or None'''
        if _range is not None:
            kwargs = {'cell_format':_range.format}
            if _range.url is not None:
                kwargs['url'] = _range.url
            self.write(*_range.range, _range.value, **kwargs)

    def write_ranges(self, range_list):
        for rng in range_list:
            self.write_range(rng)

#############################################################################################################
#############################################################################################################

class ToCitem():
    # row index on ToC sheet to write
    row_ind = 1
    # The sheet names of the workbook apart from TableOfContents
    sheets = []
    # All the self.rows
    toc_rows = []
    formats = None
    headers = ['TABLE', 'BASE', 'WBASE', 'COUNTS', 'PRCNTS', 'STATS', 'ROW', 'LABEL']
    last_col = None

    def __init__(self, table):
        l_range = Range((ToCitem.row_ind, 0), f'# {table.label}', ToCitem.formats[0])
        b_range = Range((ToCitem.row_ind, 1), table.base, ToCitem.formats[0])
        w_range = Range((ToCitem.row_ind, 2), table.wbase, ToCitem.formats[-4])
        cps_range = (table.first_row, table.first_col)
        if 'counts' in ToCitem.sheets:
            c_range = Range(
                (ToCitem.row_ind, 3),
                'cnts',
                ToCitem.formats[-2],
                url=url_from_sheet_range('counts', cps_range),
            )
        else:
            c_range = None
        if 'percentages' in ToCitem.sheets:
            p_range = Range(
                (ToCitem.row_ind, 4),
                'prcnts',
                ToCitem.formats[-2],
                url=url_from_sheet_range('percentages', cps_range),
            )
        else:
            p_range = None
        if 'stats' in ToCitem.sheets:
            s_range = Range(
                (ToCitem.row_ind, 5),
                'stats',
                ToCitem.formats[-2],
                url=url_from_sheet_range('stats', cps_range),
            )
        else:
            s_range = None
        r_range = Range((ToCitem.row_ind, 6), table.first_row + 1, ToCitem.formats[0])
        t_range = Range((ToCitem.row_ind, 7), table.title, ToCitem.formats[-5])
        # Must be in the same order as header elements
        self.row = [
            ('TABLE', l_range),
            ('BASE', b_range),
            ('WBASE', w_range),
            ('COUNTS', c_range),
            ('PRCNTS', p_range),
            ('STATS', s_range),
            ('ROW', r_range),
            ('LABEL', t_range),
        ]
        ToCitem.toc_rows.append(self.row)
        ToCitem.row_ind += 1
    
    def make_ranges():
        mapper = {'counts':'COUNTS', 'percentages':'PRCNTS', 'stats':'STATS'}
        drop_cols = [v for k,v in mapper.items() if k not in ToCitem.sheets]
        if not any([tup[1].value for row in ToCitem.toc_rows for tup in row if tup[0] == 'WBASE']):
            drop_cols.append('WBASE')
        filtered_headers, filtered_toc_rows = ToCitem.filter_row_items(drop_cols, keep=False)
        header_ranges = [Range((0,i), h, ToCitem.formats[-3]) for i,h in enumerate(filtered_headers)]
        rows_ranges = [tup[1] for row in filtered_toc_rows for tup in row]
        return header_ranges + rows_ranges
    
    def filter_row_items(header_list, keep=True):
        filtered_headers = header_list if keep else [c for c in ToCitem.headers if c not in header_list]
        filtered_toc_rows = [[c for c in row if c[0] in filtered_headers] for row in ToCitem.toc_rows]
        if filtered_headers != ToCitem.headers:
            for row in filtered_toc_rows:
                for i,r in enumerate(row):
                    r[1].shift(cols=i-r[1].last_col_ind())
        ToCitem.last_col = len(filtered_headers)
        return filtered_headers, filtered_toc_rows

#############################################################################################################
#############################################################################################################

class Table(InputFile, Worksheet, ToCitem):
    # Counter to create table aa every time a table instance is created, taking into account the number of sheets
    counter = 0
    # Dict with elements like table_aa:(first_row, first_col, is_first, is_last)
    table_info_dict = {}
    # Of the form {(0,0):stats_sheet_value, ...}
    table_colored_cells_dict = {}
    # col_by_col, col_by_rest or None
    diff_type = None

    def __init__(self, rows, sheet_name, first_row=2, first_col=0, is_last=False):
        # First table row num (table label cell) OF WORKSHEET for the range definition process (zero indexed A1=(0,0))
        self.first_row = first_row
        # First table col num (table label cell) OF WORKSHEET for the range definition process (zero indexed A1=(0,0))
        self.first_col = first_col + self.DESC_COL
        # True if it is the first worksheet table False if not
        self.is_first = self.first_row == 2
        # True if it is the last worksheet table False if not
        self.is_last = is_last
        # Current row num OF WORKSHEET for the range definition process (zero indexed A1=(0,0))
        # It indicates where to write on worksheet, NOT the indices on self.rows to read from. It justs helps mapping reading to writing
        self.current_row = self.first_row
        # Current column num OF WORKSHEET for the range definition process (zero indexed A1=(0,0))
        # It indicates where to write on worksheet, NOT the indices on self.rows to read from. It justs helps mapping reading to writing
        self.current_col = self.first_col
        # The sheet where the table will be printed
        self.sheet_name = sheet_name
        # dict of the form {'title_rows':title_rows, 'column_header_rows':column_header_rows, 'other_rows':other_rows}
        self.rows = self.split_to_parts(rows)
        # The number of the rows a table occupies
        self.rows_num = sum(len(rows) for rows in self.rows.values())
        # The number of the columns a table occupies. WE DON'T COUNT THE DESCRIPTION COLUMN ON THE LEFT
        self.cols_num = max([max([len(row) for row in rows]) for rows in self.rows.values()]) - 1
        # Label, Title, Base title
        title_rows = self.rows['title_rows']
        self.label = self.value_from_coordinates(title_rows, 0, 0)
        self.title = ' '.join(self.value_from_coordinates(title_rows, 1, 0).split())
        self.base_title = None
        if has_row_label('BT', title_rows):
            self.base_title = ' '.join(rows_from_label('BT', title_rows)[0][0].split())
        self.base, self.wbase = self.locate_bases()
        self.aa = Table.counter // len(self.sheets) + 1
        Table.table_info_dict[self.aa] = (self.first_row, self.first_col, self.is_first, self.is_last)
        Table.counter += 1
        
    def locate_bases(self):
        '''If we have the usual banner:
              The table base has always row label='RT'.
              If tables run with NO weights we have a single ROW labeled 'RT' (integer values)
              If tables run WITH weights, we have 'RT' row (float values) and an 'RU' row above it (integer values)
              base = The value of the cell with the INTEGER table base. 
              wbase = The value of the cell with the FLOAT table base. 
           If we have the pseudobanner banner, return None for both
        '''
        base = wbase = None
        column_header_rows, other_rows = self.rows['column_header_rows'], self.rows['other_rows']
        # Only if we have the usual banner, not the pseudobanner.
        # We identify the usual banner by the existance of 'TOTAL' as first column header... The only way
        if rows_from_label('CL', column_header_rows)[0][1] == 'TOTAL':
            if has_row_label('RU', other_rows, lookupto=3):
                base = rows_from_label('RU', other_rows, lookupto=3)[0][1]
                wbase = rows_from_label('RT', other_rows, lookupto=3)[0][1]
            else:
                base = rows_from_label('RT', other_rows, lookupto=3)[0][1]
        return base, wbase

    def shift_current_cell(self, rows=0, cols=0):
        if self.current_row + rows > -1:
            self.current_row += rows
        if self.current_col + cols > -1:
            self.current_col += cols

    def reset_current_cell(self):
        self.current_row = self.first_row
        self.current_col = self.first_col

    def value_from_coordinates(self, rows, r, c):
        '''lookup value in self.rows by coordinate.
           If a coordinate is out of range return None.
           rows start from 0, columns start from -1 (description column)
        '''
        c += 1
        if not((0 <= r <= len(rows) - 1) and (0 <= c <= len(rows[r]) - 1)):
            return
        return rows[r][c]
    
    def split_to_parts(self, rows):
        # TB (label), VT (table title), BT (base title) rows
        title_rows = []
        # CI (top header), CH (middle header), CL (bottom header) rows
        column_header_rows = []
        # Other rows. RU, RT (Bases), RH (Hypercodes), RS (Median, MS etc), RW (counts), PV (percentages)
        other_rows = []
        # Initialize diff_type when sheet == stats
        if self.sheet_name == 'stats':
            Table.diff_type = None
        letter_found = False
        for row in rows:
            label = row[0]
            if label in ['TB', 'VT', 'BT']:
                title_rows.append(row)
            elif label in ['CI', 'CH', 'CL']:
                column_header_rows.append(row)
            else:
                other_rows.append(row)
                # Determine table diff_type
                if self.sheet_name == 'stats':
                    if Table.diff_type is None:
                        rest_row_set = set(''.join(row[2:]))
                        if set('-+').intersection(rest_row_set):
                            Table.diff_type = 'col_by_rest'
                        elif rest_row_set.intersection(set(string.ascii_letters)):
                            letter_found = True
        if letter_found and Table.diff_type is None:
            Table.diff_type = 'col_by_col'
        return {'title_rows':title_rows, 'column_header_rows':column_header_rows, 'other_rows':other_rows}

    def prev_next_ranges(ws_name):
        ranges = []
        for aa, (first_row, first_col, is_first, is_last) in Table.table_info_dict.items():
            # Up - Down hyperlinks
            if not (is_first and is_last):
                # Down
                down_cell = (first_row, first_col + 2)
                down_value = 'Next'
                next_aa = aa % len(Table.table_info_dict) + 1
                down_hyp_range = (Table.table_info_dict[next_aa][0], first_col)
                down_url = url_from_sheet_range(ws_name, down_hyp_range)
                down = Range(down_cell, down_value, Table.formats[-2], url=down_url)
                # Up
                up_cell = (first_row, first_col + 1)
                up_value = 'Prev'
                prev_aa = (aa + len(Table.table_info_dict) - 2) % len(Table.table_info_dict) + 1
                up_hyp_range = (Table.table_info_dict[prev_aa][0], first_col)
                up_url = url_from_sheet_range(ws_name, up_hyp_range)
                up = Range(up_cell, up_value, Table.formats[-2], url=up_url)
                ranges.extend([up, down])
        return ranges

    def title_ranges(self):
        '''Handle TB, VT, BT rows'''
        ranges = []
        # Question label
        url = url_from_sheet_range('TableOfContents', (self.aa, 0))
        ranges.append(Range((self.current_row, self.current_col), f'# {self.label}', self.formats[14], url=url))
        if self.DESC_COL:
            ranges.append(Range((self.current_row, 0), self.rows['title_rows'][0][0], self.formats[0]))
        self.shift_current_cell(rows=1)
        # Question title
        ranges.append(Range((self.current_row, self.current_col), self.title, self.formats[10]))
        if self.DESC_COL:
            ranges.append(Range((self.current_row, 0), self.rows['title_rows'][1][0], self.formats[0]))
        self.shift_current_cell(rows=1)
        # Question base title
        if self.base_title is not None:
            ranges.append(Range((self.current_row, self.current_col), f'BASE: {self.base_title}', self.formats[10]))
            if self.DESC_COL:
                ranges.append(Range((self.current_row, 0), self.rows['title_rows'][2][0], self.formats[0]))
            self.shift_current_cell(rows=1)
        return ranges

    def column_header_ranges(self):
        '''Handle CI, CH, CL rows'''
        ranges = []
        rows = self.rows['column_header_rows']
        for row in rows:
            label = row[0]
            if self.DESC_COL:
                ranges.append(Range((self.current_row, 0), label, self.formats[0]))
            values = row[1:]
            # Create hyperlink to other sheet same table (counts -> percentages, percentages -> counts)
            if all([label == 'CL', 'counts' in self.sheets, 'percentages' in self.sheets, self.sheet_name != 'stats']):
                cell = (self.current_row, self.current_col)
                hyp_range = (self.first_row, self.current_col)
                other_sheet = 'percentages' if self.sheet_name == 'counts' else 'counts'
                value = f'Go to {other_sheet}'
                url = url_from_sheet_range(other_sheet, hyp_range)
                ranges.append(Range(cell, value, self.formats[-2], url=url))
            non_missing = [(i,v) for i,v in enumerate(values) if v]
            # Create ranges for table column headers
            for i, tup in enumerate(non_missing):
                cell_1 = (self.current_row, non_missing[i][0] + self.current_col)
                value = non_missing[i][1]
                if self.diff_type != 'col_by_col':
                    value = strip_diff_labels(value)
                cell_2 = (cell_1[0], (non_missing[i+1][0] - 1 if i < len(non_missing) - 1 else len(row) - 2) + self.current_col)
                _range = (cell_1, cell_2)
                if (label == 'CL' and i < 1 and tup[1] == 'TOTAL') or label != 'CL':
                    format = self.formats[8]
                else:
                    format = self.formats[7]
                ranges.append(Range(_range, value, format))
            self.shift_current_cell(rows=1)
        return ranges

    def row_ranges(self):
        '''Handle RU, RT, (Weighted sample or not)
           RH (Hypercodes)
           RS (Median, Mean Score, Standard Deviation, etc)
           RW (counts: labels or TOTAL)
           PV (percentages: labels or TOTAL)
           rows
        '''
        ranges = []
        rows = self.rows['other_rows']
        line_is_odd = True
        for row in rows:
            label = row[0]
            if self.DESC_COL:
                ranges.append(Range((self.current_row, 0), label, self.formats[0]))
            # Row labels
            row_label_value = row[1]
            cell = (self.current_row, self.current_col)
            value = row_label_value
            format = self.formats[1]
            if label in ['RU', 'RT', 'RH', 'RS'] or (label in ['RW', 'PV'] and value in ['TOTAL', 'ΣΥΝΟΛΟ']):
                format = self.formats[2]
            ranges.append(Range(cell, value, format))
            # Extend hypercodes row with missing values
            if label == 'RH':
                row.extend((self.cols_num - 1) * [''])
            # Row values in table
            rest_row_values = row[2:]
            for i, item in enumerate(rest_row_values):
                cell = (self.current_row, self.current_col + i + 1)
                value = item
                # Edit cell value
                if self.sheet_name == 'percentages':
                    value = value.replace('%','')
                elif self.sheet_name == 'counts':
                    value = value
                elif self.sheet_name == 'stats':
                    value = value.strip()
                row_header_value = row[1]
                url = None
                if label == 'RH':
                    format = self.formats[15]
                elif label in ['RU', 'RT'] or (label in ['RW', 'PV'] and row_header_value in ['TOTAL', 'ΣΥΝΟΛΟ']):
                    format = self.formats[3]
                elif label  == 'RS':
                    format = self.formats[4]
                elif label  in ['RW', 'PV']:
                    format = self.formats[55] if line_is_odd and self.ALTERNATE_CLR else self.formats[5]
                # Handle cases with stat diff
                if label == 'RS':
                    row_type = 'STATS'
                elif row_header_value in ['TOTAL', 'ΣΥΝΟΛΟ']:
                    row_type = 'TOT'
                else:
                    row_type = 'MAIN'
                if self.sheet_name == 'stats' and label not in ['RU', 'RT', 'RH']:
                    if self.diff_type == 'col_by_rest':
                        value = ''.join([c for c in value if c in ['-','+']])
                    else:
                        value = ''.join([c for c in value if c in string.ascii_letters])
                    if set(value).intersection(set(f'{string.ascii_letters}-+')):
                        if '-' in value or '+' in value:
                            value_to_get_format = value
                        else:
                            value_to_get_format = f'{value}++'
                        Table.table_colored_cells_dict[cell] = value_to_get_format
                        format = format_from_symbol(self.formats, value_to_get_format, row_type, self.sheet_name)
                if self.sheet_name != 'stats' and cell in Table.table_colored_cells_dict:
                    value_to_get_format = Table.table_colored_cells_dict[cell]
                    format = format_from_symbol(self.formats, value_to_get_format, row_type, self.sheet_name)
                    if label == 'RS':
                        value = str(round(float(value), 2)).replace('.',',')
                    else:
                        value = str(int(float(value)))
                    url = url_from_sheet_range('stats', cell)
                ranges.append(Range(cell, value, format, url=url))
            # Set value for odd-even table main row (RW or PV))
            if label in ['PV', 'RW']:
                line_is_odd = not line_is_odd
            self.shift_current_cell(rows=1)
        return ranges
    
    def make_ranges(self):
        ranges_dict = {}
        ranges_dict['title_ranges'] = self.title_ranges()
        ranges_dict['column_header_ranges'] = self.column_header_ranges()
        ranges_dict['row_ranges'] = self.row_ranges()
        return ranges_dict

#############################################################################################################
#############################################################################################################

class Range():
    '''A single cell or a merged range of cells'''

    def __init__(self, _range, value, format, url=None):
        # Of the form ((0,1),(0,1)) or ((0,0),(2,1))
        self.range = dup(_range)
        # The value to be printed on the range
        self.value = value
        # The format of the range
        self.format = format
        # The url of the cell, if any
        self.url = url
        # The number of the rows a range occupies
        self.rows_num = abs(self.range[1][0] - self.range[0][0]) + 1
        # The number of the columns a range occupies.
        self.cols_num = abs(self.range[1][1] - self.range[0][1]) + 1
        self.top_left_row_ind = self.range[0][0]

    def first_row_ind(self):
        return min(self.range[0][0], self.range[1][0])

    def last_row_ind(self):
        return max(self.range[0][0], self.range[1][0])

    def first_col_ind(self):
        return min(self.range[0][1], self.range[1][1])

    def last_col_ind(self):
        return max(self.range[0][1], self.range[1][1])

    def shift(self, rows=0, cols=0):
        cell_1, cell_2 = ((self.range[0][0] + rows, self.range[0][1] + cols), (self.range[1][0] + rows, self.range[1][1] + cols))
        x1, y1 = cell_1
        x2, y2 = cell_2
        if all([x1 > -1, y1 > -1, x2 > -1, y2 > -1]):
            self.range = (cell_1, cell_2)

    def edit(self, **kwargs):
        '''kwargs keys can be some of: `_range`, `value`, `format`, `url`'''
        if '_range' in kwargs:
            self.range = kwargs['_range']
        if 'value' in kwargs:
            self.value = kwargs['value']
        if 'format' in kwargs:
            self.format = kwargs['format']
        if 'url' in kwargs:
            self.url = kwargs['url']

    def change_url_sheet(self, new_sheet):
        if self.url is not None:
            current_sheet = self.url[self.url.find(':') + 1:self.url.find('!')]
            self.url = self.url.replace(current_sheet, new_sheet)