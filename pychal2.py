#! python3

import os

#Find Rare Characters

chardict = {}

tmpctr = 0

with open(r'C:\Users\Admin\Documents\Python\specialcharscount.txt') as f:
    while True:
        text = f.read(1)
        if not text:
            print("EOF")
            break
        tmpctr = 0
        if text in chardict:
            tmpctr = tmpctr + 1
            #gets value from dictionary
            charval = chardict.get(text, 0)
            #adding dictionary value + 1
            valupdate = int(charval) + tmpctr
            #finding the text key in dictionary and updating its value
            chardict.update({text : valupdate})

        else:
            tmpctr = tmpctr + 1
            #doesn't exist, add to dictionary with value of 1
            chardict.update({text : tmpctr})

#sort chardict by value
chardictsort = [(k, chardict[k]) for k in sorted(chardict, key=chardict.get, reverse=True)]
