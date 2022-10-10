# ForestInventoryCalculator
When I was an apprentice at Sylvamo Brasil, I developed a forest inventory calculator that replaced the old one they had.

Todas as instalações e bibliotecas necessárias:
Libraries and downloads:
pip install tqdm
pip install pandas
pip install matplotlib
pip install sqlalchemy
pip install pyinstaller
pip install cx-oracle
pip install openpyxl

import os
import sys
import math
import tkinter
import subprocess
import sqlalchemy
import numpy as np
import pandas as pd
from numpy import exp
from tqdm import tqdm
from tqdm import trange
from datetime import datetime
import matplotlib.pyplot as plt
from tkinter.filedialog import askopenfilename
from sqlalchemy.exc import IntegrityError, DatabaseError

Visual Studio build tools (C++)
1º https://visualstudio.microsoft.com/pt-br/downloads/
2º Install the build tools (C++)

Database connection made with: sqlalchemy.create_engine("oracle+cx_oracle://user:password@db")

IF YOU WANT TO VIEW THE FULL CODE, PLEASE GO TO Code Model (Non-Funcional).py: https://github.com/LucasDominguesTressoldi/ForestInventoryCalculator/blob/main/Code%20Model%20(Non-Funcional).py
IF YOU WANT TO RUN AN EXAMPLE, PLEASE COPY MainApp.py FILE: https://github.com/LucasDominguesTressoldi/ForestInventoryCalculator/blob/main/MainApp.py

The program works in 3 steps but the most important is the first one, in which the tree process is run, so in the example only the first calculations are made
The program is able to process and insert new data into the database, or process and update the database (if there is already data from the same date)

Restart function means restarting the program

At the end of the steps, an Excel file is generated with the results of the previous calculations

With the folium library it is possible to show a map from Google Maps to see where the forest is located and if everything is ok
Google Maps: folium.TileLayer(tiles='https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}', attr='Google', name='Google Earth Satellite', overlay=True, control=True).add_to(mapa)

You can generate an .exe with the command below:
pyinstaller --onefile .\NameOfFile.py --> The file will be available in a folder called 'dist', in the same directory as the code
