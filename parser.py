import re, json

class File():
    ''' All relevant parsing guidelines, including structure, etc. are set on initialization '''
    def __init__( self, name, pk ):
        self.name = name
        self.raw = self.set_raw_data() # assign raw data
        self.structure = {
            'keys': [], # verifies if a scanned line of hrules is unique
            'cells': [], # stores all hrule patterns
            'field_names': [] # user-defined field names corresponding to patterns
        }
        self.data = {} # actually contains the data
        self.pk = r'' + re.escape( pk )

        self.get_structure()

    def __str__( self ):
        return 'Data file with name: ' + self.name

    def get_structure( self ):
        ''' Assign each data field plus group names to a dict "structure". Use the pattern of the horizontal rules PLUS the line immediately above each line of horizontal rules to uniquely identify each row of data. '''
        sample_data = {
            'locs': [],
            'hrules': [],
            'entries': []
        }

        largest = begin = end = found_point = hrule_count = hrule_max = 0
        for i, line in enumerate( self.raw ):
            if re.search( self.pk, line ): # upon finding the PK...
                found_point = not found_point # we've entered or exited a data point

                if found_point:
                    begin = i # indicate the beginning line number of this data point

                else: # if we are EXITING a data point...
                    end = i
                    if hrule_count > hrule_max:
                        largest = ( begin, end ) # line numbers of largest data point
                        hrule_max = hrule_count
                    hrule_count = 0 # reset count

            if found_point and Scanner.is_hrule( line ): # if we find a horizontal rule
                ''' While inside a data point, count the number of horizontal rules '''
                hrule_count += 1
                ( k, p ) = Scanner.get_pattern( self.raw, i )

                ''' Store the horizontal rule pattern '''
                if k not in self.structure['keys']:
                    self.structure['keys'].append( k )
                    self.structure['cells'].append( p )
                    sample_data['locs'].append( str( i ) )
                    sample_data['hrules'].append( self.raw[i] )
                    sample_data['entries'].append( self.raw[i+1] )
                # else:
                #     print( 'I know this pattern' )


        ''' DEBUG '''
        # print( 'Largest # of data fields found between lines: ' + str( largest ) + '.' )
        #
        # for n, pattern in enumerate( self.structure['cells'] ):
        #     print( self.structure['cells'][n], end='\n\n' )
        #
        # print( len( self.structure['cells'] ) )
        ''' DEBUG '''

        ''' UI '''
        print( 'All fields must be named to properly store data from this raw file.' )
        print( 'Please provide names for each data field in the file "' + self.name + '".' )
        print( 'Example data is shown below.\n' )
        print( '(If necessary, please browse through the raw file to determine appropriate names.)\n' )
        print( '########################################################################\n')
        for i, key in enumerate( self.structure['keys'] ):
            print( 'LINE ' + sample_data['locs'][i] + ' contains the following data:\n' )
            print( key, end = '' )
            print( sample_data['hrules'][i], end = '' )
            print( sample_data['entries'][i] )
            print( '###\n' )

        print( '\n########################################################################\n')
        ''' UI '''

    def set_raw_data( self ):
        rd = None
        with open( self.name ) as file:
            rd = file.readlines()
        return rd

class Scanner():
    @staticmethod
    def is_hrule( line ):
        ''' Number of spaced horizontal rules defines the number of fields for a given member '''
        return re.search( r'-+', line ) and not re.match( r'DS', line )

    @staticmethod
    def get_pattern( content, index ):
        p = re.finditer( r'-+', content[index] )
        this_key = content[index - 1]
        this_pattern = []

        while True:
            try:
                this_pattern.append( next( p ).span() )
            except StopIteration:
                break

        return ( this_key, this_pattern )

    @staticmethod
    def is_data():
        ''' if "DS" is present at the start of a line, then that line contains data'''
        return re.match( r'^DS', line )

    # @staticmethod
    # def parse( lin

file_rp_all = File( 'RP_ALL.txt', 'PAD' )
