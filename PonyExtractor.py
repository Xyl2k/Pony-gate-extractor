import argparse
import validators

def get_gate(f):
    gate = ''
    pe = f.read()
    print len(pe)
    if len(pe) >= 63000 and len(pe) <= 100000:
        i = pe.find('YUIPWDFILE0YUIPKDFILE0YUICRYPTED0YUI1.0') - 3
 
        if i > 0:
            while pe[i] != '\x00' and i >= 0:
                gate = pe[i] + gate
                i   -= 1
 
    return gate
 
parser = argparse.ArgumentParser(description='Extract Pony binary gate.')
parser.add_argument('FILE', type=argparse.FileType('rb'), help='Pony binary')
args = parser.parse_args()
 
gate = get_gate(args.FILE)
 
if validators.url(gate):
    print gate
else:
    print 'Gate not found!'