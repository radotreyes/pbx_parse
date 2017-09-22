'''
*** WARNING: ***
    It's strongly recommended that you don't edit this file unless
    you know exactly what you want to change in the program.

'''
import os

''' PRESET DATA '''
# default preset file to be loaded
DEFAULT_PNAME = 'presets' # file name without extensionw
DEFAULT_PPATH = os.path.join( os.getcwd(), '{}.py'.format( DEFAULT_PNAME ) )

# default content to be written to new preset files
DEFAULT_PCONTENT = "'''\nField name presets are stored here. ONLY edit this file if you wish to directly change the preset data contained within. \n\n*** EDIT THIS FILE AT YOUR OWN RISK! ***\n'''\ndata = "

DEFAULT_PDATA = '{}'.format( {} )
