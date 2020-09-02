#access file excel file
#process data from hirna completed rides

import pandas as pd
import numpy as np
from datetime import datetime
from datetime import datetime

#edit here
city = ["Davao"] #list of cities
trip_thres= 5 #in minute, threshold for trip time
repeat_thres = 30 # in minutes, threshold for repeat count
# repeat cancel is the number of repeat cancel per passenger
promo = 'REBATE30'
#do not edit after this

# labels
# RR: Requested Ride
# CR: Completed Ride

def main(*inputname):
    for arg in inputname: #accept multiple input
        loc=arg

        shname = 'Booking details report'
        #shname = 'Sheet1'
        print("Loading Document..")

        sh1 = pd.read_excel(loc, sheet_name=shname,skiprows=3) # turn data into dataframes. might need to pick columns due to size

        shloc=sh1['Pickup location'] # save pickup location

        # save data into excel file
        ofile = "output_" + loc  # save filename
        print("Creating Excel File: " + ofile)
        writer = pd.ExcelWriter(ofile)

        for ct in range(0, len(city)):  # loop per city
            tcount=0;
            print('Processing City: ' + city[ct])
            shcityreq = sh1[(shloc.str.contains(city[ct]) == True)]  # requests
            shcitycom = sh1[(shloc.str.contains(city[ct]) == True) & (sh1['Paid by'] == "Cash")]  # completed rides

            #find company name
            shcompany = shcityreq['Company'].drop_duplicates().reset_index(
                drop=True)  # get unique company names, creates a series

            # get unique dates

            shdates = pd.to_datetime(shcityreq['Pickup time'], format='%m/%d/%Y %I:%M %p')
            shdatesuni = pd.DataFrame(zip(shdates.apply(lambda x: x.month), shdates.apply(lambda x: x.year)),
                                      columns=['month', 'year'])
            shdatesuni = shdatesuni[['month', 'year']].drop_duplicates()  # get unique month-year combination
            shdatesuni.sort_values(by=['year', 'month'])  # sort by ascending, priority by year

            # create storage
            mlabel = list()
            mlabel.append('Total')
            for k in range(0, shdatesuni['month'].count()):
                mlabel.append('M' + str(shdatesuni.iloc[k, 0]) + 'Y' + str(shdatesuni.iloc[k, 1]))  # create month label

            passRR = pd.DataFrame(np.zeros((shcompany.size, shdatesuni['month'].count() + 1)), columns=mlabel,
                                  index=shcompany.values.tolist())  # include total column
            passCR = passRR.copy()  # Completed Ride

            print('Processing Rides for: ' + city[ct])
            for pname in range(0, shcompany.count()):
                rep_count = 0  # initialize repeat count
                # get completed rides
                if(tcount>100):
                    print(datetime.now().strftime("%H:%M:%S"))
                    tcount=0
                else:
                    tcount=+1
                shcomp = shcitycom[shcitycom['Company'] == shcompany[pname]]  # get match passenger name, completed
                if not shcomp.empty:
                    shdtrip = pd.to_datetime(shcomp['Pickup time'], format='%m/%d/%Y %I:%M %p')
                    shdtrip = pd.DataFrame(zip(shdtrip.apply(lambda x: x.month), shdtrip.apply(lambda x: x.year)),
                                           columns=['month', 'year'])
                    for k in range(0, shdtrip['month'].count()):
                        cname = 'M' + str(shdtrip.iloc[k, 0]) + 'Y' + str(shdtrip.iloc[k, 1])
                        passCR.loc[shcompany[pname], cname] += 1
                # get requested rides
                shreqt = shcityreq[shcityreq['Company'] == shcompany[pname]]  # get match passenger name, requested
                if not shreqt.empty:
                    shdtrip = pd.to_datetime(shreqt['Pickup time'], format='%m/%d/%Y %I:%M %p')
                    shdtrip = pd.DataFrame(zip(shdtrip.apply(lambda x: x.month), shdtrip.apply(lambda x: x.year)),
                                           columns=['month', 'year'])
                    for k in range(0, shdtrip['month'].count()):
                        cname = 'M' + str(shdtrip.iloc[k, 0]) + 'Y' + str(shdtrip.iloc[k, 1])
                        passRR.loc[shcompany[pname], cname] += 1

            #Compute for total
            passCR['Total'] = passCR.iloc[:, 1:].sum(axis=1)
            passRR['Total'] = passRR.iloc[:, 1:].sum(axis=1)

            # save into Excel Passenger Data
            print('Writing Passenger Data for ' + city[ct])
            passCR.to_excel(writer, 'OperatorCR_' + city[ct], index=True)
            passRR.to_excel(writer, 'OperatorRR_' + city[ct], index=True)

        writer.save()

        


if __name__== "__main__":
    #loc='completedrides.xlsx'
    #loc='completedridedec2to8.xlsx'
    #loc = 'testpromo.xlsx'

    main('completedrides.xlsx')
    #main('testdata.xlsx')
