from main import *

prompt = lambda x: askstring( 'Input requred', x )
confirm = lambda x: messagebox.showinfo( 'Confirm', x )
choose = lambda x: messagebox.askquestion( 'Input required', x )

true = lambda x: True if x == True or x == 'yes' else False
false = lambda x: True if x != True or x == False or x == None or x == 'no' else False
# else:
#     pr = cf = ch = print
#
#     t = lambda x: True if x == True or x == 'y' else False
#     f = lambda x: True if x != True or x == False or x == None or x == 'n' else False
#
# return ( pr, cf, ch, t, f )
