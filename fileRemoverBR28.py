#refer to BR28 in the BRD to understand why this file exists

#after you've read BR28, i know this isn't finished, the only thing remaining is to change the file to an "output.xlsx" file

import os
from datetime import datetime

created = os.stat('Data.sqlite').st_ctime
now = datetime.now()
dateTimeFile = datetime.fromtimestamp(created)

dateFile = dateTimeFile.strftime("%m/%d/%Y")

dateNow = now.strftime("%m/%d/%Y")

if(dateNow == dateFile):
    os.remove('Data.sqlite')
    print("fileRemoved")
