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

import settings
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

        # If the file exists, confirm if the user wants to overwrite
        if os.path.isfile( path ):
            while True:
                print( 'The file "{}.py" already exists. Overwrite? (Y/N)'.format( name ) )
                w = input( '>> ' ).lower()
                if w == 'y':
                    print( 'Overwriting...' )
                    break
                elif w == 'n':
                    print( 'Didn\'t overwrite.' )
                    return False # stop trying to create a new file
                print( '\nPlease enter a valid input.\n' )

        # change the current preset file path
        cls.ppath = path
        cls.pname = name

        # create the new preset file
        with open( path, 'w' ) as f:
            f.write( settings.DEFAULT_PCONTENT )
            f.write( settings.DEFAULT_PDATA )

        # load the new preset data
        Presets.load_pfile()

    @classmethod
    def load_pfile( cls ):
        ''' Load preset data from an existing file. '''
        # change the current preset data
        presets = __import__( cls.pname )
        cls.pdata = presets.data

        print( 'Loading presets from {} ...'.format( cls.ppath ) )

    @classmethod
    def change_pfile( cls, name ):
        '''
        Load another preset file based on user input.
        IN: _name_, preset filename excluding extension.
        '''
        # define the file path
        path = os.path.join( cls.cd, '{}.py'.format( name ) )

        if not os.path.isfile( path ):
            # create the file if it doesn't exist
            print( 'The file {}.py doesn\'t exist. Creating it now.'.format( name ) )
            Presets.new_pfile( name )
        else:
            # change the file path and load data
            cls.ppath = os.path.join( cls.cd, '{}.py'.format( name ) )
            cls.pname = name
            Presets.load_pfile()

        print( 'Changed presets file to {}.py.'.format( name ) )

    @classmethod
    def get_pdata( cls ):
        '''
        Retrieve the preset data from the currently imported preset file.
        '''
        print( 'Getting saved presets from {}.py ...'.format( cls.pname ) )
        if cls.pdata:
            print( json.dumps( cls.pdata, indent = 2 ) )
        else:
            print( 'There are no presets here!' )

    @classmethod
    def append_pdata( cls, value ):
        '''
        Add a new entry to the currently imported preset file.
        IN: _value_, the value of the preset that is to be saved.
        '''
        while True:
            key = input( 'Please enter a name to save this preset under, or type \'exit\' to go back:\n>> ' )
            if not key or not re.search( r'\S', key ):
                print( '\nPlease enter a non-blank name.\n' )
            elif key.lower() == 'exit':
                return False # exit the function
            else:
                try:
                    # see if the preset already exists
                    if cls.pdata[key]:
                        while True:
                            print( 'There is already a preset with this name. Do you want to overwrite it? (Y/N)' )
                            w = input( '>> ').lower()
                            if w == 'y':
                                cls.pdata[key] = value
                                Presets.save_pdata()
                                print( json.dumps( cls.pdata, indent = 2 ) )
                                return False # exit the function
                            elif w == 'n':
                                print( 'Didn\'t overwrite.' )
                                break
                            print( '\nPlease enter a valid input.\n' )
                except KeyError:
                    # the preset doesn't already exist, so create it
                    cls.pdata[key] = value
                    Presets.save_pdata()
                    print( json.dumps( cls.pdata, indent = 2 ) )
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
