import settings
from savedata import Presets

import re, os, json
from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill

class RawFile():
    def __init__( self, path ):
        self.name = path.split( '.' )[0]

        with open( path ) as file:
            self.content = file.readlines()

        self.records = self.get_records()

        print( 'Found {} records in this file:'.format( len( self.records ) ) )
        if self.records:
            for key, value in self.records.items():
                print( '\tRecord "{}" with length {}'.format( key, len( value ) ) )

                self.process_record( key, value )

    def get_records( self ):
        '''
        Scans raw file content and separates data records by looking for the string "COMMAND: LI". Data records are saved to _self.records_.
        '''
        records = {}; cache = []
        found = False; num = 0
        for i, line in enumerate( self.content ):
            # scan for EOF
            try:
                nextline = self.content[i+1]
            except IndexError:
                nextline = None

            # if we find a PBX command
            if re.search( r'^COMMAND: LI \w+', line ):
                # indicate that we have found a record
                found = True

                # if any data is currently cached
                if cache:
                    # save it to record list
                    key = self.name + '({})'.format( num ) if num else self.name
                    records[key] = cache
                    num += 1 # increment record number
                    cache = [] # clear cache

            # while in a record, save lines to cache
            if found:
                cache.append( line )

            # we've reached EOF, save cache to record list
            if not nextline:
                key = self.name + '({})'.format( num ) if num else self.name
                records[key] = cache

        return records

    def process_record( self, name, rawdata ):
        # parses the command associated with this data record
        cmd = rawdata[0].split( 'COMMAND:' )[1].split( '\n' )[0]

        # creates a new Record object from this data record
        test_record = Record( name, rawdata, cmd )


class Record():
    '''
    Data contained in a RawFile, as separated by a "COMMAND: LI". All relevant parsing guidelines, including structure, etc.
    are set on initialization.
    '''

    def __init__( self, name, rawdata, command ):
        # Open file
        self.raw = rawdata

        input( '\nParsing record ' + name + '.' )

        ''' LOADING PRESETS '''
        if Presets.pdata:
            while True:
                print( 'You have saved presets.' )
                print( 'Would you like to load a preset for this file? (Y/N)' )
                use = input( '>> ' ).lower()
                if use == 'y':
                    pdata = Presets.get_pdata()
                    self.meta = pdata['meta']
                    self.structure = pdata['structure']
                    self.presets = True
                    break
                elif use == 'n':
                    self.presets = False
                    break
                else:
                    print( 'Please enter a valid input.' )
        else:
            self.presets = False

        if not self.presets:
            # Metadata
            self.meta = {
                'name': name,
                'pk': None,
                'pk_inline': False, # if the PK contains an inline data point
                'command': command
            }
            self.set_meta()

            self.structure = {
                'keys': [], # contains the line of data above each horizontal rule
                            # used to verify uniqueness of each hrule pattern
                'cells': [], # unique horizontal rule patterns in the file
                'fix': [], # specifies if fields need to be align-fixed
                'field_names': [], # user-defined field names
                'set_ids': [], # integers corresponding 1:1 to a data pattern
                'write_columns': [], # starting indices of grouped field names, 0ind
            }
            self.set_structure()

        # Data object and output worksheet
        self.data = {} # actually contains the data
        self.sheet = test_book.active

        # Read the file and output to Excel
        # TODO: Move to user control
        self.scrape()
        self.write()
        ''' '''

    def __str__( self ):
        return 'Data file with name: ' + self.name

    def set_meta( self ):
        os.system( 'cls' if os.name == 'nt' else 'clear' )
        print( 'Please enter the primary key for this data set.' )
        while True:
            pk = input( '>> ' )
            if not pk or not re.search( r'\S', pk ):
                print( '\nPlease enter a non-blank name.\n' )
            else:
                break
        self.meta['pk'] = r'' + re.escape( pk )

        os.system( 'cls' if os.name == 'nt' else 'clear' )

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
            try:
                nextline = self.raw[i+1]
            except IndexError:
                nextline = None

            if Scanner.is_pk( self.meta['pk'], line ): # upon finding the PK...
                point = not point # we've entered or exited a data point

                # If the PK contains a colon, then include the text after that colon as a data point
                if re.search( r':', line ) and not self.meta['pk_inline']:
                    self.structure['keys'].append( 'PRIMARY KEY' )
                    self.structure['cells'].append( [ (
                        len( line.split( ':' )[0] ) + 1,
                        len( line ) - 1
                    ), ] )
                    self.structure['set_ids'].append( f_id )
                    self.structure['fix'].append( False )
                    self.meta['pk_inline'] = True
                    # print( self.structure['cells'][0] )
                    # input( type( self.structure['cells'][0] ) )
                    f_id += 1

            if point and Scanner.is_hrule( line, nextline ): # if we find a horizontal rule
                ( k, c ) = Scanner.get_pattern( self.raw, i )

                # store the horizontal rule pattern
                if k not in self.structure['keys']:
                    # input( 'Appending {} on line {}'.format( c, i ) )
                    self.structure['keys'].append( k )
                    self.structure['cells'].append( c )
                    self.structure['set_ids'].append( f_id )

                    # does the data need to be align-fixed?
                    if Scanner.is_data( self.raw[i+1] ):
                        '''
                        Data can sometimes be shifted one line down.
                        If this happens, an align-fix is needed.
                        If an align-fix is needed, data will be parsed from
                        the next line down, rather than the line that the data
                        is normally on.
                        '''
                        if re.search( self.raw[i+1][2:].strip(), self.raw[i+2] ):
                            self.structure['fix'].append( True )
                        else:
                            self.structure['fix'].append( False )

                    f_id += 1

                    ''' DISPLAYED IN UI '''
                    sample_data['locs'].append( str( i ) )
                    sample_data['hrules'].append( self.raw[i] )
                    sample_data['entries'].append( self.raw[i+1] )

        if not self.presets:
            ''' UI '''
            os.system( 'cls' if os.name == 'nt' else 'clear' )

            x = offset = 0 # offset index, to be used if pk_inline is True
            for i, key in enumerate( self.structure['keys'] ):
                x = i + offset

                self.structure['field_names'].append( [] )

                for j, cell in enumerate( self.structure['cells'][i] ):
                    if key != 'PRIMARY KEY':
                        print( 'All fields must be named to properly store data from this raw file.' )
                        print( 'Please provide names for each field in the file "' + self.meta['name'] + '".' )
                        print( 'Example data is shown below.\n' )
                        print( '(If necessary, please browse through the raw file to determine appropriate names.)\n' )

                        ( f, l ) = cell

                        print( '#####################################################################\n')
                        print( 'Displaying data from LINE ' + sample_data['locs'][x] + ':\n' )
                        print( '> ' + key, end = '' )
                        print( '> ' + sample_data['hrules'][x], end = '' )
                        print( '> ' + sample_data['entries'][x], end = '' )
                        print( '> ' + ' ' * int( f ) + '^' )
                        print( '\n#####################################################################\n')

                        print( 'Please provide a name for this unnamed field.' )
                        print( '(Names must consist of non-blank characters.' )
                        print( ' For any given line of data, all field names must be unique.)\n' )
                        print( '>\t' + self.structure['keys'][i][f:l] )
                        print( '>\t' + sample_data['hrules'][x][f:l] )
                        print( '>\t' + sample_data['entries'][x][f:l] + '\n' )

                        while True:
                            name = input( '>> ' )
                            if not name or not re.search( r'\S', name ):
                                print( '\nPlease enter a non-blank name.\n' )
                            elif name in self.structure['field_names'][i]:
                                print( '\nThat name is already in use for this line. Please enter a unique name.\n' )
                            else:
                                break
                    else:
                        name = self.meta['pk']
                        offset -= 1

                    self.structure['field_names'][i].append( name )
                    print( '\n' )
                    os.system( 'cls' if os.name == 'nt' else 'clear' )
            ''' /UI '''

        ''' SAVING PRESETS '''
        while True:
            print( 'Save your entries as a new preset? (Y/N)')
            save = input( '>> ' ).lower()
            if save == 'y':
                Presets.append_pdata( self.meta, self.structure )
                Presets.save_pdata()
                break
            elif save == 'n' :
                break
            else:
                print( 'Please enter a valid input.' )

        os.system( 'cls' if os.name == 'nt' else 'clear' )


    def scrape( self ):
        '''
        Look for data according to self.structure and store it in _self.data_.
        '''
        n = -1 # member index
        for i, line in enumerate( self.raw ):
            try:
                nextline = self.raw[i+1]
            except IndexError:
                nextline = None

                # push the last data_set to member
                try:
                    member['**SET_ID: ' + set_key] = data_set
                except UnboundLocalError: pass

                # push the last member to self.data
                self.data[n] = member

            if Scanner.is_pk( self.meta['pk'], line ):
                # create a new data member
                if n != -1:
                    # push previous data set into current member
                    try:
                        print( 'Found PK. Pushing data_set to member[]... (MEMBER: {}, SET_ID: {})'.format( n, set_key ) )
                        member['**SET_ID: ' + set_key] = data_set
                    except UnboundLocalError: pass
                    # carry on in case we haven't found a data_set yet

                    ''' zip the previously parsed data point'''
                    self.data[n] = member
                    if n != -1:
                        print( 'Pushing member[] to self.data[]... (MEMBER: {}, SET_ID: {})'.format( n, set_key ) )

                print( 'Clearing member dict.' )
                member = {}
                print( 'Clearing data_set dict.' )
                try: data_set = {}
                except UnboundLocalError: pass

                n += 1
                print( 'Now populating member[{}]'.format( n ) )

                if self.meta['pk_inline']:
                    print( 'Found PK: {}'.format( line ) )
                    i = self.structure['cells'][0][0][0]
                    j = len( line ) - 1
                    member['**SET_ID: 0'] = {
                        '**ENTRY#: 0': {
                            self.meta['pk']: line[i:j].strip()
                        }
                    }

            if Scanner.is_hrule( line, nextline ):
                # push previous data set into current member
                try:
                    if n != -1:
                        print( '(Found HRULE. Pushing data_set to member[]... (MEMBER: {}, SET_ID: {})'.format( n, set_key ) )
                    member['**SET_ID: ' + set_key] = data_set
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

                # sort each set into its own dict
                set_key = str( id_to_map )

                # if there is already an entry in this set under the same hrule pattern:
                try:
                    if member['**SET_ID: ' + set_key]:
                        print( 'Found a duplicate with set key: {}'.format( set_key ) )
                        data_set = member['**SET_ID: ' + set_key]
                        c = len( data_set )
                    else:
                        c = 0
                        data_set = {}
                except KeyError:
                    c = 0
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
                        Scanner.parse( 'DS' + self.raw[i+1], cell )
                        for cell in cells_to_map
                    ]
                else:
                    this_data = [
                        Scanner.parse( line, cell )
                        for cell in cells_to_map
                    ]

                # push data into appropriate set
                data_set['**ENTRY#: {}'.format( c )] = dict( zip( this_names, this_data ) )

                c += 1 # number of times data has been pushed under one hrule

        print( json.dumps( self.data, indent = 2) )

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
            self.sheet.cell( row = 1, column = i+1 ).fill =     PatternFill( 'solid', fgColor = "D0D0D0" )

        current_row = 2
        for i, m in enumerate( self.data ):
            current_row += Scanner.transcribe( self.data[m], self.sheet, current_row, self.structure['write_columns'] )

        test_book.save( 'test.xlsx' )

class Scanner():
    @staticmethod
    def is_hrule( line, nextline ):
        ''' Number of spaced horizontal rules defines the number of fields for a given member '''
        if nextline:
            return re.search( r'-+', line ) and not re.match( r'^DS', line ) and re.match( r'^DS', nextline )

        return False

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
            set_id = int( re.search( r'\d+', s ).group() )
            set_len = len( member[s] ) if len( member[s] ) > set_len else set_len

            for e in member[s]:
                entry_num = int( re.search( r'\d+', e ).group() )
                for i, f in enumerate( member[s][e] ):
                    r = cur + entry_num
                    c = wc[set_id] + i + 1

                    cell = ss.cell( row = r, column = c )
                    cell.value = member[s][e][f]
                    if c == 1:
                        cell.fill = PatternFill( 'solid', fgColor = "FFFF00" )

        return set_len

''' testing presets '''
Presets.load_pfile()

''' for use with test file '''
test_book = Workbook()
file_rp_all = RawFile( 'SERVICE_LIST.TXT' )
