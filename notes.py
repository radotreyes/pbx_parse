# with open( 'RP_ALL.txt' ) as file:
#     content = file.readlines()
#
#     if content:
#         print( 'I can read this' )
#         print( 'Type is: ' + str( type( content ) ) )
#         print( content[0][0] )
#         print( 'First datum is: ' + content[0] )
#         print( 'Second datum is:' + content[13] )
#         print( content[0][8] == ' ' )
import re
import json

'''
- readlines() stores data line by line
- both rows and columns are zero-indexed
- blank chars are spaces
'''
has_values = [ 7, 13, 14, 15 ] # relative lines in file containing data
has_values_test = [ 7, ]
cursor = [
        [ 3, 13, 23, 25, 28, 33, 37, 45, 53, 55, 57, 65, 67, 70, 73 ],
    ]
print( 'Validating file...', end = ' ' )
with open( 'RP_ALL.txt' ) as file:
    print( 'OK.' )
    content = file.readlines()

    def is_hrule( line ):
        return re.search( r'-+', line ) and not re.match( r'DS', line )

    def is_data( line ):
        return re.match( r'^DS', line )

    def parse( line ):
        ( i, j ) = line
        value = content[n][i:j]
        return value.strip() if re.match( r'\S', value ) else ' '

    ''' Scope is entire file '''
    data = {} # data gathered from parse
    member = -1 # single member/row, to be used as a 1:1 key in 'data'
                # starts at -1 so that when a member is found, iteration starts at 0
    pk = r'PAD' # DEBUG: id that signifies a new data member
    group_keys = []

    for i, line in enumerate( content ):
        ''' Scope is the line being scanned '''
        if re.search( pk, line ): # found new data point
            ''' reset group index and create a new data member '''
            member += 1; group = -1
            data[member] = {}; entry = data[member]
        if is_hrule( line ): # the line is a horizontal rule
            group += 1; entry[group] = []
            n = i + 1 # start looking at the following lines
            while True:
                if is_data( content[n] ): # if there is data
                    # get positions of the headers
                    data_map = []
                    hrule_iter = re.finditer( r'-+', line )
                    while( True ):
                        try:
                            data_map.append( next( hrule_iter ).span() )
                        except StopIteration:
                            if ( group + 1 ) > len( group_keys ):
                                print( 'Found unnamed data fields. Please name these fields before continuing.' )
                                group_keys.append( [
                                    input( 'Enter a label for this value: ' + parse(n) + '\n>> '  ) for n in data_map
                                ] )
                            entry[group].extend( [
                                parse(n) for n in data_map
                            ] )
                            break
                    n += 1
                else: # continue this until no more data is found
                    break
            break # DEBUG

    print( json.dumps( data[0], indent=2 ) )
    print( group_keys )
