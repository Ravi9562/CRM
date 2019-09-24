import re
import time

Input = open("Body.txt", "r")
Output = open("Details.txt", "w+")


#C:\Users\tejaw\PycharmProjects\CRM\Collecting_Details_Without_Duplicates.py
def Searching_strings():
    print('Collecting Details')
    for line in Input:
        if re.match("(.*)(Phone:)(.*)", line):
            print(line, file=Output)
        if re.match('(.*)(Email:)(.*)', line):
            print(line, file=Output)
        if re.match('(.*)(Name:)(.*)', line):
            print(line, file=Output)
        if re.match('(.*)(Company:)(.*)', line):
            print(line, file=Output)
        if re.match('(.*)(Product:)(.*)', line):
            print(line, file=Output)
        if re.match('(.*)(Topic:)(.*)', line):
            print(line, file=Output)


    Output.close()
    Input.close()


    #Below script is to eliminate duplicate values

    lines_seen = set() #holds lines already seen
    outfile = open("Accurate.txt", "w+")
    for line in open('Details.txt', "r"):
        if line not in lines_seen: # not a duplicate
            outfile.write(line)
            lines_seen.add(line)
    outfile.close()

if __name__ == "__main__":
    while True:
        Searching_strings()
        time.sleep(300)
