import re, os, json
from openpyxl import Workbook

class File():
    ''' All relevant parsing guidelines, including structure, etc.
        are set on initialization '''

    def __init__( self, name, pk ):
        self.name = name
        self.raw = self.set_raw_data() # assign raw data
        self.structure = {
            'keys': [], # verifies if a scanned line of hrules is unique
            'cells': [], # stores all hrule patterns
            'field_names': [], # user-defined field names corresponding to patterns
            'set_ids': [] # list of integers corresponding 1:1 to a data pattern
        }
        self.data = {} # actually contains the data
        self.pk = r'' + re.escape( pk )
        self.sheet = test_book.active

        self.set_structure()
        self.scrape()

    def __str__( self ):
        return 'Data file with name: ' + self.name

    def set_raw_data( self ):
        rd = None
        with open( self.name ) as file:
            rd = file.readlines()
        return rd

    def set_structure( self ):
        ''' Assign each data field plus group names to a dict "structure".
            Use the pattern of the horizontal rules
            PLUS the line immediately above each line of horizontal rules
            to uniquely identify each row of data. '''

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

                ''' Store the horizontal rule pattern '''
                if k not in self.structure['keys']:
                    self.structure['keys'].append( k )
                    self.structure['cells'].append( c )
                    self.structure['set_ids'].append( f_id )
                    f_id += 1

                    ''' DISPLAYED IN UI '''
                    sample_data['locs'].append( str( i ) )
                    sample_data['hrules'].append( self.raw[i] )
                    sample_data['entries'].append( self.raw[i+1] )

        ''' DEBUG '''
        # print( 'Largest # of data fields found between lines: ' + str( largest ) + '.' )
        #
        # for n, pattern in enumerate( self.structure['cells'] ):
        #     print( self.structure['cells'][n], end='\n\n' )
        #
        # print( len( self.structure['cells'] ) )
        ''' /DEBUG '''

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
        ''' Look for data according to self.structure and store it in 'data'. '''
        n = -1 # member index
        for i, line in enumerate( self.raw ):
            if Scanner.is_pk( self.pk, line ):
                # create a new data member
                if n != -1:
                    # print( len( this_names ) )
                    # print( len( this_data ) )
                    ''' zip the previously parsed data point'''
                    self.data[n] = member

                member = {}
                n += 1

            if Scanner.is_hrule( line ):
                # push the previous data set into current member
                try:
                    member['**SET_ID: ' + set_key] = data_set
                except UnboundLocalError:
                    # set_key is undefiend or data_set is empty, continue
                    pass

                # retrieve the structure index determined by the closest key
                for j, key in enumerate( self.structure['keys'] ):
                    if self.raw[i-1] == key:
                        index = j

                # grab field data from self.structure
                names_to_map = self.structure['field_names'][index]
                cells_to_map = self.structure['cells'][index]
                id_to_map = self.structure['set_ids'][index]

                # reset number of times data has been pushed under one hrule
                c = 0

                # sort each set into its own dict
                set_key = str( id_to_map )
                data_set = {}

            if Scanner.is_data( line ):
                this_names = names_to_map
                this_data = [
                    Scanner.parse( line, cell ) for cell in cells_to_map
                ]

                # push data into appropriate set
                data_set['**ENTRY#: ' + str( c )] = dict( zip( this_names, this_data ) )

                c += 1 # number of times data has been pushed under one hrule

                # print( json.dumps( data_set, indent = 2 ) )

        # test_fields = self.field_names
        print( self.structure['field_names'] )
        test_fields = [ field for group in self.structure['field_names'] for field in group  ]
        for i, field in enumerate( test_fields ):
            self.sheet.cell( row = 1, column = i+1 ).value = field
        test_book.save( 'test.xlsx' )
        # print( json.dumps( self.data, indent = 2 ) )


class Scanner():
    @staticmethod
    def is_hrule( line ):
        ''' Number of spaced horizontal rules defines the number of fields for a given member '''
        return re.search( r'-+', line ) and not re.match( r'^DS', line )

    @staticmethod
    def is_pk( pk, line ):
        return re.search( pk, line )

    @staticmethod
    def get_pattern( content, index ):
        p = re.finditer( r'-+', content[index] )
        this_key = content[index - 1]
        this_pattern = []

        while True:
            try:
                this_pattern.append( next( p ).span() )
            except StopIteration: break

        return ( this_key, this_pattern )

    @staticmethod
    def is_data( line ):
        ''' if "DS" is present at the start of a line, then that line contains data'''
        return re.match( r'^DS', line )

    @staticmethod
    def parse( line, cell ):
        ( i, j ) = cell
        return line[i:j].strip() if re.match( r'\S', line[i:j] ) else ' '

test_book = Workbook()
file_rp_all = File( 'RP_ALL.txt', 'PAD' )
