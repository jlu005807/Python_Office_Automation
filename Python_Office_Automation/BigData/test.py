records= [
 ('foo',1,2),
 ('bar','hello'),
 ('foo',3,4),
 ]

def do_foo(x,y):
        print('foo', x, y)
def do_bar(s):
        print('bar', s)
for tag,*args in records:
    if tag== 'foo':
        do_foo(*args)
    elif tag =='bar':
        do_bar(*args)

import heapq
nums = [1, 8, 2, 23, 7,-4, 18, 23, 42, 37, 2]
print(heapq.nlargest(3, nums)) # Prints [42, 37, 23]
print(heapq.nsmallest(3, nums)) # Prints [-4, 1, 2]