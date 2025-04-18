#!/usr/bin/env python
import pdfkit
import sys
import os

# dir_list = os.listdir("./Customers/")

for x in os.listdir("./Customers/"):
    if x.endswith(".html"):
        # Prints only text file present in My Folder
        print(x)
        split_tup = os.path.splitext(x)
        # filepath = "./Customers/"+x
        # pdfpath = "./Customers/"
        pdfkit.from_file("./Customers/"+ x, "./Customers/"+split_tup[0]+".pdf")


    # print("Python >> Hello, " + name)
    #  return "You are " + name