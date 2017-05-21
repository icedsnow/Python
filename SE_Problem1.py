#! python3

list1 = []

for i in range(0, 15):
    list1.append(i)

def forsolve(list):
    a = 0
    listlen = len(list)
    for i in range(0, listlen):
        a += list1[i]
    return a

def whilesolve(list):
    b = 0
    tmpctr = 0
    listlen = len(list)
    while listlen > 0:
        b += list1[tmpctr]
        tmpctr += 1
        listlen -= 1
    return b

def recurse(list):
        return sum(list[0:])

def sum(list):
   if len(list) == 1:
        return list[0]
   else:
        return list[0] + sum(list[1:])
