'''
Siemens PBX Parsing Tool
- This program reads data patterns from Siemens PBX output files
    and returns the data contained within.

AUTHOR: Reuben Reyes
ORIG: 2017/9/17
    - radotreyes.github.io
    - github.com/radotreyes/pbx_parse

'''
import tkinter
from tkinter import *
from tkinter.ttk import Frame, Button, Style, Label, Entry
from tkinter.filedialog import askopenfilename

import os
import parser, settings
from savedata import Presets

from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill

class AppWindow( Frame ):
    def __init__( self ):
        super().__init__()
        self.init_UI_static()
        self.init_UI_dynamic()

        w = 600 # width
        h = 150 # height
        x = self.master.winfo_screenwidth() # screen offset X
        y = self.master.winfo_screenheight() # screen offset Y

        x = int( ( x - w ) / 2 ) # center horizontal
        y = int( ( y - h ) / 2 ) # center vertical
        self.master.geometry( '{}x{}+{}+{}'.format( w, h, x, y ) )

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
        self.frame_data.pack( fill = tkinter.X )
        self.lbl_data = Label( self.frame_data, text = 'Data File', width = 12 )
        self.lbl_data.pack( side = tkinter.LEFT, padx = 5, pady = 10 )

        # Excel Workbook
        self.frame_xlsx = Frame( self )
        self.frame_xlsx.pack( fill = tkinter.X )
        self.lbl_xlsx = Label( self.frame_xlsx, text = 'Workbook', width = 12 )
        self.lbl_xlsx.pack( side = tkinter.LEFT, padx = 5, pady = 10 )

        # Presets File
        self.frame_preset = Frame( self )
        self.frame_preset.pack( fill = tkinter.X )
        self.lbl_preset = Label( self.frame_preset, text = 'Presets File:', width = 12 )
        self.lbl_preset.pack( side = tkinter.LEFT, padx = 5, pady = 10 )

        # button grid
        button_grid = Frame( self, borderwidth = 1 )
        button_grid.pack( fill = BOTH, expand = True )

        self.btn_start = Button( self, text = 'Parse', command = lambda: self.parse( self.data_file, self.xlsx_file ) )
        btn_xlsx = Button( self, text = 'Load Workbook', command = self.set_xlsx )
        btn_file = Button( self, text = 'Load Data File', command = self.set_data )
        btn_preset = Button( self, text = 'Load Presets', command = self.set_preset )
        btn_quit = Button( self, text = 'Quit', command = self.quit )

        self.btn_start.pack( side = tkinter.LEFT, padx = 5, pady = 10 )
        btn_xlsx.pack( side = tkinter.LEFT, padx = 5, pady = 10 )
        btn_file.pack( side = tkinter.LEFT, padx = 5, pady = 10 )
        btn_preset.pack( side = tkinter.LEFT, padx = 5, pady = 10 )
        btn_quit.pack( side = tkinter.RIGHT, padx = 5, pady = 10 )

    def init_UI_dynamic( self ):
        # initial dynamic UI states
        self.data_file = None
        self.data_files = StringVar()

        self.xlsx_file = None
        self.xlsx_files = StringVar()

        self.preset_file = settings.DEFAULT_PPATH
        self.preset_files = StringVar()
        self.preset_files.set( '{}'.format( self.preset_file ) )

        print( 'Running' )
        # update data file display
        display_data = Label( self.frame_data,
            textvariable = self.data_files,
            relief = tkinter.SUNKEN,
            background = '#FFF' )
        display_data.pack( fill = tkinter.X, padx = 5, expand = True )

        # update excel file display
        display_xlsx = Label( self.frame_xlsx,
            textvariable = self.xlsx_files,
            relief = tkinter.SUNKEN,
            background = '#FFF' )
        display_xlsx.pack( fill = tkinter.X, padx = 5, expand = True )

        # update preset file display
        display_preset = Label( self.frame_preset,
            textvariable = self.preset_files,
            relief = tkinter.SUNKEN,
            background = '#FFF' )
        display_preset.pack( fill = tkinter.X, padx = 5, expand = True )

    def parse( self, raw, book ):
        ''' Execute the parser '''
        Presets.load_pfile()
        f = parser.RawFile( raw, str( book ) )

    def set_parse( self ):
        if self.data_file and self.xlsx_file and self.preset_file:
            self.btn_start.config( state = 'normal' )
        else:
            self.btn_start.config( state = 'disabled' )

    def set_data( self ):
        ftypes = [ ( 'Text files', '*.txt' ) ]
        self.data_file = askopenfilename( filetypes = ftypes )
        self.data_files.set( '{}'.format( self.data_file ) )
        print( self.data_file )
        print( type( self.data_file ) )
        print( os.path.split( self.data_file ) )
        self.set_parse()

    def set_xlsx( self ):
        ftypes = [
            ( 'Excel workbooks', '*.xlsx' ),
            ( 'Excel workbooks (legacy)', '*.xls' )
        ]
        self.xlsx_file = askopenfilename( filetypes = ftypes )
        self.xlsx_files.set( '{}'.format( self.xlsx_file ) )
        self.set_parse()

    def set_preset( self ):
        ftypes = [ ( 'Python files', '*.py' ) ]
        self.preset_file = askopenfilename( filetypes = ftypes )
        self.preset_files.set( '{}'.format( self.preset_file ) )
        self.set_parse()

if __name__ == '__main__':
    root = Tk()
    app = AppWindow()
    root.mainloop()
