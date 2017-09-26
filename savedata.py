'''
Handles storage of user data, including:
    - Field name presets
    - Data file structures

    - Structure is:
    _data_
        _entry_
            _structure_
            _pk_
            _command_
'''

import settings, parser, main
from main import *
from relay import *
import re, os, json

class Presets():
    '''
    Handle file presets.
        - All file presets are stored in the same working directory.
    '''
    cd = os.getcwd()
    ppath = settings.DEFAULT_PPATH # preset file path
    pname = settings.DEFAULT_PNAME # preset file name without extension
    pdata = None
    request = None

    # DEFAULT PROCEDURE
    # creset preset file if it doesn't exist
    if not os.path.isfile( ppath ):
        with open( ppath, 'w' ) as f:
            # create an empty data object
            f.write( settings.DEFAULT_PCONTENT )
            f.write( settings.DEFAULT_PDATA )

    @classmethod
    def new_pfile( cls, name ):
        '''
        Create or overwrite a new preset file.
        IN: _name_, preset filename excluding extension.
        '''
        # define the file path
        path = os.path.join( cls.cd, '{}.py'.format( name ) )
        input( 'Creating preset in directory: {}'.format( path ) )

        # If the file exists, confirm if the user wants to overwrite
        if os.path.isfile( path ):
            while True:
                if request:
                    w = choose( 'The file "{}.py" already exists. Overwrite? (Y/N)'.format( name ) )
                else:
                    choose( 'The file "{}.py" already exists. Overwrite? (Y/N)'.format( name ) )
                    w = input( '>> ' ).lower()
                if true( w ):
                    confirm( 'Overwriting...' )
                    break
                elif false( w ):
                    confirm( 'Didn\'t overwrite.' )
                    return False # stop trying to create a new file
                confirm( '\nPlease enter a valid input.\n' )

        # change the current preset file path
        cls.ppath = path
        cls.pname = name

        # create the new preset file
        with open( path, 'w' ) as f:
            f.write( settings.DEFAULT_PCONTENT )
            f.write( settings.DEFAULT_PDATA )

        # load the new preset data
        Presets.load_pfile()

        return path

    @classmethod
    def load_pfile( cls ):
        ''' Load preset data from an existing file. '''
        # change the current preset data
        presets = __import__( cls.pname )
        cls.pdata = presets.data

        print( 'Loading presets from {} ...'.format( cls.ppath ) )

    @classmethod
    def change_pfile( cls, path ):
        '''
        Load another preset file based on user input.
        IN: _name_, preset filename excluding extension.
        '''
        # define the file path
        if not path.endswith( '.py' ):
            path = os.path.join( cls.cd, '{}.py'.format( path ) )

        if cls.cd in path:
            name = path.split( cls.cd + '/' )[1].split( '.py' )[0]
        else:
            name = path.split( '.py' )[0]

        if not os.path.isfile( path ):
            # create the file if it doesn't exist
            confirm( 'The file {}.py doesn\'t exist. Creating it now.'.format( name ) )
            Presets.new_pfile( name )
        else:
            # change the file path and load data
            cls.ppath = os.path.join( cls.cd, '{}.py'.format( name ) )
            cls.pname = name
            Presets.load_pfile()

        confirm( 'Changed presets file to {}.py.'.format( name ) )

    @classmethod
    def get_pdata( cls ):
        '''
        Retrieve the preset data from the currently imported preset file.
        '''
        if cls.pdata:
            preset_list = [ preset for preset in cls.pdata ]
            k = ListDialog( 'Load a preset', 'Please choose a preset:', preset_list )
            cls.request.wait_window( k.top )
            try:
                return cls.pdata[k.key.get()]
            except KeyError:
                return False
        else:
            confirm( 'There are no presets here!' )
            return False

    @classmethod
    def append_pdata( cls, meta, structure ):
        '''
        Add a new entry to the currently imported preset file.
        IN: _value_, the value of the preset that is to be saved.
        '''
        while True:
            key = prompt( 'Please enter a name to save this preset under.' )
            if not key or not re.search( r'\S', key ):
                confirm( '\nPlease enter a non-blank name.\n' )
            elif key.lower() == 'exit':
                return False # exit the function
            else:
                try:
                    # see if the preset already exists
                    if cls.pdata[key]:
                        while True:
                            choose( 'There is already a preset with this name. Do you want to overwrite it? (Y/N)' )
                            w = input( '>> ').lower()
                            if true( w ):
                                cls.pdata[key]['meta'] = meta
                                cls.pdata[key]['structure'] = structure
                                Presets.save_pdata()

                                return False # exit the function
                            elif false( w ):
                                confirm( 'Didn\'t overwrite.' )
                                break
                except KeyError:
                    # the preset doesn't already exist, so create it
                    cls.pdata[key] = {}
                    cls.pdata[key]['meta'] = meta
                    cls.pdata[key]['structure'] = structure
                    Presets.save_pdata()
                    return False # exit the function

    @classmethod
    def save_pdata( cls ):
        '''
        Save preset data class variable to the currently imported
        preset file.
        '''
        with open( cls.ppath, 'w' ) as f:
            f.write( settings.DEFAULT_PCONTENT )
            f.write( '{}'.format( cls.pdata ) )

if __name__ == '__main__':
    Presets.load_pfile()
    print( Presets.ppath )
    print( Presets.pdata )
