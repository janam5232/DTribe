import zipfile as zp
import glob as gb
import os

#getting the current directories locaiton (where the data will be)
path = os.getcwd()

#accessing the zip file
file = gb.glob(os.path.join(path, "*.zip"))

#extracting the zip file
zip = zp.ZipFile(file[0])
zip.extractall("C:\\Users\\janam\\Desktop\\")

zip.close()