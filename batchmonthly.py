#batch process extractrides

import os
import monthlystat

def main():
    filenames=[]
    for file in os.listdir('.'):
        if file.endswith(".xlsx"):
            print("Processing: " + file)
            monthlystat.main(file)

    #filenames='completedridenov18to24.xlsx'

if __name__== "__main__":
    main()
