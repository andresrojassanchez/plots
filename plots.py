import numpy as np
import pandas as pd
import pygsheets
import matplotlib.pyplot as plt
import itertools
import pylab
import os

#gc = pygsheets.authorize(outh_file = '/home/andres/Documents/plots/creds.json')

cols = 
files = [file for file in os.listdir(os.curdir) if file.endswith(".csv")]
print(files)

for file in files:

    df = pd.read_csv(file, skiprows = 11)
    filename = os.path.splitext(file)[0]
    filename = df
    

print(df)
#gc.create('plots')
#sh = gc.open('plots')
#wks = sh[0]
#wks.set_dataframe(df, (1,1))