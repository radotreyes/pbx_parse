import settings
import re, os, json
from openpyxl import Workbook

class Storage():
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
                                Storage.save_pdata()
                                print( json.dumps( cls.pdata, indent = 2 ) )
                                return False # exit the function
                            elif w == 'n':
                                print( 'Didn\'t overwrite.' )
                                break
                            print( '\nPlease enter a valid input.\n' )
                except KeyError:
                    # the preset doesn't already exist, so create it
                    cls.pdata[key] = value
                    Storage.save_pdata()
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
        Storage.load_pfile()

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
            Storage.new_pfile( name )
        else:
            # change the file path and load data
            cls.ppath = os.path.join( cls.cd, '{}.py'.format( name ) )
            cls.pname = name
            Storage.load_pfile()

        print( 'Changed presets file to {}.py.'.format( name ) )

class File():
    '''
    All relevant parsing guidelines, including structure, etc.
    are set on initialization.
    '''

    def __init__( self, name, pk ):
        self.name = name
        with open( self.name ) as file:
            self.raw = file.readlines() # assign raw data

        self.structure = {
            'keys': [], # contains the line of data above each horizontal rule
                        # used to verify uniqueness of each hrule pattern
            'cells': [], # unique horizontal rule patterns in the file
            'fix': [], # specifies if fields need to be align-fixed
            'field_names': [], # user-defined field names
            'set_ids': [], # integers corresponding 1:1 to a data pattern
            'write_columns': [], # starting indices of grouped field names, 0ind
        }
        self.data = {} # actually contains the data
        self.pk = r'' + re.escape( pk )
        self.sheet = test_book.active

        ''' Perform on initialization for now, will move to user control later. '''
        self.set_structure()
        self.scrape()
        self.write()
        ''' '''

    def __str__( self ):
        return 'Data file with name: ' + self.name

    def set_structure( self ):
        '''
        - Assign each data field plus group names to a dict "structure".
        - Use the pattern of the horizontal rules
            PLUS the line immediately above each line of horizontal rules
            to uniquely identify each row of data.
        '''

        sample_data = {
            'locs': [],
            'hrules': [],
            'entries': []
        }

        point = False; f_id = 0
        for i, line in enumerate( self.raw ):
            if Scanner.is_pk( self.pk, line ): # upon finding the PK...
                point = not point # we've entered or exited a data point

            if point and Scanner.is_hrule( line ): # if we find a horizontal rule
                ( k, c ) = Scanner.get_pattern( self.raw, i )

                # store the horizontal rule pattern
                if k not in self.structure['keys']:
                    self.structure['keys'].append( k )
                    self.structure['cells'].append( c )
                    self.structure['set_ids'].append( f_id )

                    # does the data need to be align-fixed?
                    if Scanner.is_data( self.raw[i+1] ):
                        print( 'Found data' )
                        print( self.raw[i+1][2:] )
                        print( 'Next line:' )
                        print( self.raw[i+2] )
                        if re.search( self.raw[i+1][2:].strip(), self.raw[i+2] ):
                            self.structure['fix'].append( True )
                        else:
                            self.structure['fix'].append( False )

                    f_id += 1

                    ''' DISPLAYED IN UI '''
                    sample_data['locs'].append( str( i ) )
                    sample_data['hrules'].append( self.raw[i] )
                    sample_data['entries'].append( self.raw[i+1] )

        print( json.dumps( self.structure, indent = 2) )

        ''' UI '''
        for i, key in enumerate( self.structure['keys'] ):
            self.structure['field_names'].append( [] )

            for j, cell in enumerate( self.structure['cells'][i] ):
                print( 'All fields must be named to properly store data from this raw file.' )
                print( 'Please provide names for each field in the file "' + self.name + '".' )
                print( 'Example data is shown below.\n' )
                print( '(If necessary, please browse through the raw file to determine appropriate names.)\n' )

                ( f, l ) = cell

                print( '#####################################################################\n')
                print( 'Displaying data from LINE ' + sample_data['locs'][i] + ':\n' )
                print( '> ' + key, end = '' )
                print( '> ' + sample_data['hrules'][i], end = '' )
                print( '> ' + sample_data['entries'][i], end = '' )
                print( '> ' + ' ' * int( f ) + '^' )
                print( '\n#####################################################################\n')

                print( 'Please provide a name for this unnamed field.' )
                print( '(Names must consist of non-blank characters.' )
                print( ' For any given line of data, all field names must be unique.)\n' )
                print( '>\t' + self.structure['keys'][i][f:l] )
                print( '>\t' + sample_data['hrules'][i][f:l] )
                print( '>\t' + sample_data['entries'][i][f:l] + '\n' )

                while True:
                    name = input( '>> ' )
                    if not name or not re.search( r'\S', name ):
                        print( '\nPlease enter a non-blank name.\n' )
                    elif name in self.structure['field_names'][i]:
                        print( '\nThat name is already in use for this line. Please enter a unique name.\n' )
                    else:
                        break
                self.structure['field_names'][i].append( name )
                print( '\n' )
                os.system( 'cls' if os.name == 'nt' else 'clear' )
        ''' /UI '''

    def scrape( self ):
        '''
        Look for data according to self.structure and store it in _self.data_.
        '''
        n = -1 # member index
        for i, line in enumerate( self.raw ):
            if Scanner.is_pk( self.pk, line ):
                # create a new data member
                if n != -1:
                    # push previous data set into current member
                    try:
                        member['**SET_ID: ' + set_key] = data_set
                    except UnboundLocalError:
                        pass # carry on in case we haven't found a data_set yet

                    ''' zip the previously parsed data point'''
                    self.data[n] = member
                    if n == 0 or n == 1:
                        print( 'Pushing member[] to self.data[]... (MEMBER: {}, SET_ID: {})'.format( n, set_key ) )

                member = {}
                n += 1

            if Scanner.is_hrule( line ):
                # push previous data set into current member
                try:
                    member['**SET_ID: ' + set_key] = data_set
                    if n == 0 or n == 1:
                        print( 'Pushing new set to member[]... (MEMBER: {}, SET_ID: {})'.format( n, set_key ) )
                except UnboundLocalError:
                    pass # carry on in case we haven't found a data_set yet

                # retrieve the structure index determined by the closest key
                for j, key in enumerate( self.structure['keys'] ):
                    if self.raw[i-1] == key:
                        index = j

                # grab field data from self.structure
                names_to_map = self.structure['field_names'][index]
                cells_to_map = self.structure['cells'][index]
                id_to_map = self.structure['set_ids'][index]
                fix_to_map = self.structure['fix'][index]

                # reset number of times data has been pushed under one hrule
                c = 0

                # sort each set into its own dict
                set_key = str( id_to_map )
                data_set = {}

            if Scanner.is_data( line ):
                this_names = names_to_map

                if fix_to_map:
                    '''
                    This path applies an ALIGN-FIX.
                        - Instead of parsing the line the data is on,
                            parse the next line shifted by 2.
                        - Sometimes necessary in faulty Siemens PBX
                            output files.
                    '''
                    this_data = [
                        Scanner.parse( 'DS' + self.raw[i+1], cell ) for cell in cells_to_map
                    ]
                else:
                    this_data = [
                        Scanner.parse( line, cell ) for cell in cells_to_map
                    ]

                # push data into appropriate set
                if n == 0 or n == 1:
                    print( 'Pushing new entry to member[]... (MEMBER: {}, SET_ID: {}, ENTRY: {})'.format( n, set_key, c ) )
                data_set['**ENTRY#: ' + str( c )] = dict( zip( this_names, this_data ) )

                c += 1 # number of times data has been pushed under one hrule

    def write( self ):
        '''
        Write the file to Excel.
        '''
        flat_fields = [] # flattened list of field names for column assignment
        for group in self.structure['field_names']:
            for i, field in enumerate( group ):
                flat_fields.append( field )
                if i == 0:
                    self.structure['write_columns'].append( len( flat_fields ) - 1 )

        for i, field in enumerate( flat_fields ):
            self.sheet.cell( row = 1, column = i+1 ).value = field

        current_row = 2
        for i, m in enumerate( self.data ):
            current_row += Scanner.transcribe( self.data[m], self.sheet, current_row, self.structure['write_columns'] )

        test_book.save( 'test.xlsx' )

class Scanner():
    @staticmethod
    def is_hrule( line ):
        ''' Number of spaced horizontal rules defines the number of fields for a given member '''
        return re.search( r'-+', line ) and not re.match( r'^DS', line )

    @staticmethod
    def is_pk( pk, line ):
        '''
        Searches for a primary key in the given line.
        IN: _pk_, primary key to look for
            _line_, line to search in
        '''
        return re.search( pk, line )

    @staticmethod
    def get_pattern( hrule_line, index ):
        '''
        Grabs the horizontal rule pattern from the appropriate line.
        IN: _hrule_line_, a line a hrule pattern
            _index_, the index of that line
        '''
        p = re.finditer( r'-+', hrule_line[index] )
        this_key = hrule_line[index - 1]
        this_pattern = []

        while True:
            try:
                this_pattern.append( next( p ).span() )
            except StopIteration: break

        return ( this_key, this_pattern )

    @staticmethod
    def is_data( line ):
        '''
        If "DS" is present at the start of a line, then that line contains data
        IN: _line_, the line to scan
        '''
        return re.match( r'^DS', line )

    @staticmethod
    def parse( line, cell ):
        '''
        Grabs data from a line according to its data structure,
        which is defined in _FileObject.structure_
        IN: _line_, the line to scan
            _cell_, the coordinate span of _line_ which contains data
        '''
        ( i, j ) = cell
        return line[i:j].strip() if re.match( r'\S', line[i:j] ) else ' '

    @staticmethod
    def transcribe( member, ss, cur, wc ):
        '''
        Writes data to Excel.
        IN: _member_, the data member which is being transcribed
            _ss_, the spreadsheet to write to
            _cur_, the current row number which should be written to
            _wc_, write column; the column number which should be written to
        '''
        set_len = 1 # rows to be reserved for this particular data set
        for s in member: # find the number of rows required
            set_id = int( re.search( r'\d', s ).group() )
            set_len = len( member[s] ) if len( member[s] ) > set_len else set_len

            for e in member[s]:
                entry_num = int( re.search( r'\d', e ).group() )
                for i, f in enumerate( member[s][e] ):
                    # print( member[s][e][f] )
                    ss.cell( row = cur + entry_num,
                        column = wc[set_id] + i + 1 ).value = member[s][e][f]

        return set_len

''' for use with test file '''
test_book = Workbook()
file_rp_all = File( 'SLI.txt', 'PAD' )

''' testing presets '''
# Storage.load_pfile()
# Storage.change_pfile( 'poop' )
# Storage.get_pdata()
# Storage.append_pdata( 'test' )
