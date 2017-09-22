'''
Siemens PBX Parsing Tool
- This program reads data patterns from Siemens PBX output files
    and returns the data contained within.

AUTHOR: Reuben Reyes
ORIG: 2017/9/17
    - radotreyes.github.io
    - github.com/radotreyes/pbx_parse

'''
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Frame, Button, Style, Label, Entry
from tkinter.filedialog import askopenfilename, asksaveasfile, askdirectory
from tkinter.simpledialog import askstring

import os, re
import parser, settings, savedata

from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill

class AuxList( Frame ):
    def __init__( self, title, list, master = None ):
        super().__init__()

        # parent frame geometry
        w = 150 # width
        h = 300 # height
        x = self.master.winfo_screenwidth() # screen offset X
        y = self.master.winfo_screenheight() # screen offset Y

        x = int( ( x - w ) / 2 ) # center horizontal
        y = int( ( y - h ) / 2 ) # center vertical
        self.master.geometry( '{}x{}+{}+{}'.format( w, h, x, y ) )

        self.master.title( title )
        self.pack( fill = BOTH, expand = 1 )

        self.style = Style()
        self.style.theme_use( 'clam' ) # default theme for now

        self.frame_list = Frame( self )
        self.frame_list.pack( fill = X )
        self.list = Listbox( self.frame_list,
            relief = SUNKEN,
            background = '#FFF' )
        self.list.pack( side = TOP, fill = X, padx = 5, pady = ( 10, 5 ), expand = True )

        self.button_grid = Frame( self, borderwidth = 1 )
        self.button_grid.pack( fill = BOTH, expand = True )
        self.btn_start = Button( self, text = 'Select', command = None )
        self.btn_start.pack( side = LEFT, padx = 5, pady = 10 )
        self.btn_quit = Button( self, text = 'Back', command = None )
        self.btn_quit.pack( side = RIGHT, padx = 5, pady = 10 )

class Main( Frame ):
    def __init__( self ):
        super().__init__()

        # parent frame geometry
        w = 600 # width
        h = 480 # height
        x = self.master.winfo_screenwidth() # screen offset X
        y = self.master.winfo_screenheight() # screen offset Y

        x = int( ( x - w ) / 2 ) # center horizontal
        y = int( ( y - h ) / 2 ) # center vertical
        self.master.geometry( '{}x{}+{}+{}'.format( w, h, x, y ) )

        # this is the list containing all data files to be parsed
        self.files_to_parse = []

        # initialize child UI elements
        self.init_UI_static()
        self.init_UI_dynamic()
        self.set_parse()

    def __str__( self ):
        return 'Tkinter GUI'

    def init_UI_static( self ):
        ''' Initialize static UI elements '''
        # initialize window title and grid layout
        self.master.title( 'Siemens PBX Parsing Tool' )
        self.pack( fill = BOTH, expand = 1 )

        # style
        self.style = Style()
        self.style.theme_use( 'clam' ) # default theme for now

        # Data File
        self.frame_data = Frame( self )
        self.frame_data.pack( fill = X )
        self.lbl_data = Label( self.frame_data, text = 'Data Files:', width = 12, relief = FLAT )
        self.lbl_data.pack( side = LEFT, ipadx = 5, padx = 5, ipady = 5, pady = 5 )

        self.frame_data_btn = Frame( self )
        self.frame_data_btn.pack( fill = X )
        btn_rm_data = Button( self.frame_data_btn, text = 'Remove', command = self.rm_data )
        btn_rm_data.pack( side = RIGHT, padx = 5 )
        btn_add_data = Button( self.frame_data_btn, text = 'Add', command = self.set_data )
        btn_add_data.pack( side = RIGHT, padx = 5 )

        # Excel Workbook
        self.frame_xlsx = Frame( self )
        self.frame_xlsx.pack( fill = X, pady = ( 10, 0 ) )
        self.lbl_xlsx = Label( self.frame_xlsx, text = 'Workbook:', width = 12, relief = FLAT )
        self.lbl_xlsx.pack( side = LEFT, ipadx = 5, padx = 5, ipady = 5, pady = 5 )
        self.frame_xlsx_btn = Frame( self )
        self.frame_xlsx_btn.pack( fill = X )
        btn_load_wb = Button( self.frame_xlsx_btn, text = 'Load...', command = self.set_wb )
        btn_load_wb.pack( side = RIGHT, padx = 5 )
        btn_new_wb = Button( self.frame_xlsx_btn, text = 'New...', command = lambda: self.set_wb( True ) )
        btn_new_wb.pack( side = RIGHT, padx = 5 )

        # Presets File
        self.frame_preset = Frame( self )
        self.frame_preset.pack( fill = X, pady = ( 10, 0 ) )
        self.lbl_preset = Label( self.frame_preset, text = 'Presets File:', width = 12, relief = FLAT )
        self.lbl_preset.pack( side = LEFT, ipadx = 5, padx = 5, ipady = 5, pady = 5 )
        self.frame_preset_btn = Frame( self )
        self.frame_preset_btn.pack( fill = X )
        btn_preset = Button( self.frame_preset_btn, text = 'Choose...', command = self.set_preset )
        btn_preset.pack( side = RIGHT, padx = 5)

        # Lower button grid
        button_grid = Frame( self, borderwidth = 1 )
        button_grid.pack( fill = BOTH, expand = True )

        self.btn_start = Button( self, text = 'Parse', command = self.parse )
        self.btn_start.pack( side = LEFT, padx = 5, pady = 10 )

        btn_quit = Button( self, text = 'Quit', command = self.quit )
        btn_quit.pack( side = RIGHT, padx = 5, pady = 10 )

    def init_UI_dynamic( self ):
        # initial dynamic UI states
        self.wb_file = None
        self.wb_files = StringVar()

        self.preset_file = settings.DEFAULT_PPATH
        self.preset_files = StringVar()
        self.preset_files.set( '{}'.format( self.preset_file ) )

        print( 'Running' )
        # update data file display
        self.display_data = Listbox( self.frame_data,
            relief = SUNKEN,
            background = '#FFF' )
        self.display_data.pack( side = TOP, fill = X, padx = 5, pady = ( 10, 5 ), expand = True )

        # update excel file display
        display_xlsx = Label( self.frame_xlsx,
            textvariable = self.wb_files,
            relief = SUNKEN,
            background = '#FFF' )
        display_xlsx.pack( fill = X, padx = 5, expand = True )

        # update preset file display
        display_preset = Label( self.frame_preset,
            textvariable = self.preset_files,
            relief = SUNKEN,
            background = '#FFF' )
        display_preset.pack( fill = X, padx = 5, expand = True )

    def parse( self ):
        ''' Execute the parser '''
        for f in self.files_to_parse:
            print( 'PARSING:\n\tDATA FILE: {}\n\tWORKBOOK: {}'.format( f, self.wb_file ) )
            parser.RawFile( f, self.wb_file, self )

    def set_parse( self ):
        if self.files_to_parse and self.wb_file and self.preset_file:
            self.btn_start.config( state = 'normal' )
        else:
            self.btn_start.config( state = 'disabled' )

    def set_data( self ):
        ftypes = [ ( 'Text files', '*.txt' ) ]
        f = askopenfilename( filetypes = ftypes )

        # push item to listbox
        self.display_data.insert( END, f )

        # push item to list
        self.files_to_parse.append( f )
        print( self.files_to_parse )

        # update parse button
        self.set_parse()

    def rm_data( self ):
        # remove item from list
        self.files_to_parse.remove( self.display_data.get( ACTIVE ) )

        # remove item from listbox
        self.display_data.delete( ACTIVE )

        # update parse button
        self.set_parse()

    def set_wb( self, new = False ):
        ftypes = [
            ( 'Excel workbooks', '*.xlsx' ),
            ( 'Excel workbooks (legacy)', '*.xls' )
        ]
        if new:
            # prompt user for file name
            while True:
                # make sure file name includes at least one non-whitespace character
                filename = askstring(
                    title = 'Create a new Excel workbook',
                    prompt = 'Enter a name for the new workbook. (Exclude the file extension)'
                )
                try:
                    if not re.search( r'\S', filename ):
                        messagebox.showinfo(
                            'Invalid name',
                            'Please enter a non-blank file name.'
                        )
                    else:
                        filename += '.xlsx'
                        break
                except TypeError:
                    # escape if user hits 'cancel'
                    return False

            # choose a directory and check for duplicate files
            while True:
                filedir = askdirectory(
                    title = 'Choose a directory to save "{}"'.format( filename )
                )
                # escape if user hits 'cancel'
                if not filedir: return False
                filepath = os.path.join( filedir, filename )

                # confirm file overwrite if the file exists
                if os.path.exists( filepath ):
                    ow = messagebox.askquestion(
                        'File already exists',
                        '{} already exists in this directory. Do you want to overwrite this file?'.format( filename )
                    )

                    # overwrite the file if the user clicks 'yes'
                    if ow == 'yes': break

                else:
                    break

            # save the new workbook
            new_wb = Workbook()
            new_wb.save( filepath )

        else:
            filepath = askopenfilename( filetypes = ftypes )

        # update data labels
        self.wb_file = filepath
        self.wb_files.set( '{}'.format( self.wb_file ) )
        self.set_parse()

    def set_preset( self ):
        ftypes = [ ( 'Python files', '*.py' ) ]
        self.preset_file = askopenfilename( filetypes = ftypes )
        self.preset_files.set( '{}'.format( self.preset_file ) )
        self.set_parse()

    def list_select( self, title, presets ):
        top = Toplevel( self )
        top.title( title )
        # parent frame geometry
        w = 300 # width
        h = 400 # height
        x = self.master.winfo_screenwidth() # screen offset X
        y = self.master.winfo_screenheight() # screen offset Y

        x = int( ( x - w ) / 2 ) # center horizontal
        y = int( ( y - h ) / 2 ) # center vertical
        top.geometry( '{}x{}+{}+{}'.format( w, h, x, y ) )

        top.frame_list = Frame( top )
        top.frame_list.pack( fill = BOTH, expand = True )
        top.list = Listbox(
            top.frame_list,
            relief = SUNKEN,
            background = '#FFF'
        )
        top.list.pack(
            fill = BOTH,
            padx = 5,
            pady = ( 10, 5 ),
            expand = True
        )
        for preset in presets:
            top.list.insert( END, preset )

        top.button_grid = Frame( top, borderwidth = 1 )
        top.button_grid.pack( fill = BOTH, expand = True )
        top.btn_start = Button( top, text = 'Select', command = None )
        top.btn_start.pack( side = LEFT, padx = 5, pady = 10 )
        top.btn_quit = Button( top, text = 'Back', command = None )
        top.btn_quit.pack( side = RIGHT, padx = 5, pady = 10 )


if __name__ == '__main__':
    savedata.Presets.load_pfile()
    root = Tk()
    app = Main()
    root.mainloop()
