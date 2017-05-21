#! python3

"""
Write a function that combines two lists by alternatingly taking elements.
For example: given the two lists [a, b, c] and [1, 2, 3],
the function should return [a, 1, b, 2, c, 3].
"""

list1 = ['a', 'b', 'c', 'd', 'e', 'f', 'g']
list2 = ['1', '2', '3', '4', '5', '6', '7']

def com(x, y):
    newlist = []
    xlen = len(x)
    ylen = len(y)
    esc = True
    tmpctr = 0
    while esc is True:
        if xlen > ylen:
            listlen = xlen
            esc = False
        else:
            listlen = ylen
            esc = False
    for i in range(0, listlen):
        while tmpctr % 2 == 0:
            xval = x[i]
            newlist.append(xval)
            tmpctr += 1
        yval = y[i]
        newlist.append(yval)
        tmpctr +=1
    return newlist


"""
newlist = []
x = list1
y = list2
xlen = len(x)
ylen = len(y)
esc = True
tmpctr = 0
while esc is True:
    if xlen > ylen:
        listlen = xlen
        esc = False
    else:
        listlen = ylen
        esc = False
for i in range(0, listlen):
    while tmpctr % 2 == 0:
        xval = x[i]
        newlist.append(xval)
        tmpctr += 1
    yval = y[i]
    newlist.append(yval)
    tmpctr +=1
print(newlist)
"""
