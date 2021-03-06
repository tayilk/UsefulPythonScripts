#This script is meant to be used with .csv file output from the
#Office 365 User Activity report generated from the 365 admin reports dashboard.
#It spits out 4 .csv files, which may seem like much, but I wanted to sort users by
#Activity (in Exchange specifically), Inactivity, Non-Licensed, and All.  
# Just run this script in the same directory as the report, and change the 
# #file variable so that it matches your file name

#Make sure to clean up the data from the report so UPN, Dipslay Name, Has Exchange,
#and Exchange Activity Date are the first four columns of the .csv in that order.  

import csv

#Opening Data source and creating output files
file = open('Office365ActiveUserDetail.csv')
csv_f = csv.reader(file)
csv_active = csv.writer(open("ActiveUsers.csv", "w", newline='')) 
csv_inactive = csv.writer(open("InactiveUsers.csv", "w", newline=''))
csv_nolicense = csv.writer(open("NonLicensedUsers.csv", "w", newline=''))
csv_alldata = csv.writer(open("Office365ActiveUserDatailSorted.csv", "w", newline=''))
licensed = ""
mailbox = ""

fields = ["Display Name", "Mailbox", "Licensed", "Last Active Date"]

#writing fields columns to our output files
csv_inactive.writerow(fields)
csv_active.writerow(fields)
csv_nolicense.writerow(fields)
csv_alldata.writerow(fields)

#declaring arrays to hold our row data
inactiverows = []
activerows = []
nolicenserows = []
alldatarows = []

#clearing fields row
next(csv_f)

for row in csv_f:

    displayname = row[1]
    licensed = row[2]
    mailbox = row[0]
    outrow = [displayname, mailbox, licensed, row[3]]

    if licensed == "TRUE":
        
        if row[3] == "":
            outrow[3] = "No Activity"
            inactiverows.append(outrow)
            alldatarows.append(outrow)
            print(mailbox + "is inactive")
        else:
            activerows.append(outrow)
            alldatarows.append(outrow)
            print(mailbox + " last active " + row[3])
    else:
        nolicenserows.append(outrow)
        alldatarows.append(outrow)

#sorting data (alphabetical)
inactiverows.sort()
activerows.sort()
nolicenserows.sort()
alldatarows.sort()

#output
for irow in inactiverows:
    csv_inactive.writerow(irow)
    print(irow)

for arow in activerows:
    csv_active.writerow(arow)
    print(arow)

for lrow in nolicenserows:
    csv_nolicense.writerow(lrow)
    print(lrow)

for aldatrow in alldatarows:
    csv_alldata.writerow(aldatrow)


        
    
