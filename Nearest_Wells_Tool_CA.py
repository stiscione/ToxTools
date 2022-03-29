import tkinter as tk
import xlrd
import xlsxwriter
import pandas as pd
import sqlite3
from math import sin, cos, sqrt, atan2, radians
from tkinter import filedialog
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

root= tk.Tk()
#name title of program
root.title("Nearby Well Distance Tool")

#declare file name string var
file_var=tk.StringVar()
db_var=tk.StringVar()
dis_var=tk.IntVar()

# setting the windows size
root.geometry("400x300")


#cas lookup function
def nearest_well ():
    inputfile = file_var.get()
    inputdb = db_var.get()
    conn = sqlite3.connect(inputdb)
    cur = conn.cursor()
    writer = pd.ExcelWriter('out.xlsx', engine='xlsxwriter')
    wb = xlrd.open_workbook(inputfile)
    sheet = wb.sheet_by_index(0)
    distance_threshold = dis_var.get()
    # approximate radius of earth in miles
    R = 3959.999
    for row in range(sheet.nrows):
        site_id = sheet.cell_value(row,0)
        site_name = sheet.cell_value(row,1)
        site_lat = sheet.cell_value(row,2)
        site_long = sheet.cell_value(row,3)
        # with open('well_output.csv', 'w', newline='') as csvfile:
        # csvwriter = csv.writer(csvfile, delimiter=',',
        #                         quotechar='|', quoting=csv.QUOTE_MINIMAL)
        # csvwriter.writerow(['globalid', 'lat', 'long', 'direction'])
        site_lat_r = radians(site_lat)
        site_long_r = radians(site_long)
        df = pd.DataFrame(columns = ['globalid', 'locationid', 'lat', 'long','direction','distance','value','fieldptclass','date'])
        #will query all wells meeting criteria we are interested in and iterate through them and calculating distance from site coordinates
        for well in cur.execute('SELECT globalID, locationID, latitude, longitude, sum(CASE WHEN qualifier = "=" THEN value ELSE 0 END), date, fieldptclass FROM pfas_ca WHERE matrix = "Liquid" AND longitude IS NOT NULL AND latitude IS NOT NULL GROUP BY latitude, longitude, fieldptclass, date'):
            globalid = well[0]
            locationid = well[1]
            if well[2]=="" or well[2]==None or well[3] == "" or well[3] == None:
                continue
            well_lat = float(well[2])
            well_long = float(well[3])
            #print (well_lat, well_long)
            well_lat_output = well[2]
            well_long_output = well[3]
            value = well[4]
            date = well[5]
            fieldptclass = well[6]
            date = datetime.strptime(date, '%m/%d/%Y')
            if site_lat > well_lat:
                north_south = "south"
            else:
                north_south  = "north"
            if site_long > well_long:
                east_west = "west"
            else:
                east_west = "east"
            direction = north_south + '-' + east_west
            # site_lat = radians(site_lat)
            # site_long = radians(site_long)
            well_lat_r = radians(well_lat)
            well_long_r = radians(well_long)
            dlon = well_long_r - site_long_r
            dlat = well_lat_r - site_lat_r
            a = sin(dlat / 2)**2 + cos(well_lat_r) * cos(site_lat_r) * sin(dlon / 2)**2
            c = 2 * atan2(sqrt(a), sqrt(1 - a))
            distance = R * c
            if distance <= float(distance_threshold):
                data = pd.DataFrame([[globalid, locationid, well_lat_output, well_long_output, direction,distance,value, fieldptclass, date]], columns = ['globalid', 'locationid', 'lat', 'long','direction','distance','value','fieldptclass','date'])
                df = df.append(data)
            else:
                continue
        df.to_excel (writer,sheet_name=site_id, index=False)
    writer.save()
    cur.close()

def find_file ():
    filetypes = (
        ('excel files', '*.xls'),
        ('All files', '*.*')
    )
    file_var = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    file_entry.insert(0, file_var)

def find_db ():
    filetypes = (
        ('sqlite files', '*.sqlite'),
        ('All files', '*.*')
    )
    db_var = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    db_entry.insert(0, db_var)

#label for entry fields
file_label = tk.Label(root, text = 'filename', font=('calibre',10, 'bold'))
db_label = tk.Label(root, text = 'database name', font=('calibre',10,'bold'))
dis_label = tk.Label(root, text = 'distance (miles)', font=('calibre',10,'bold'))

#entry space for file and db
file_entry = tk.Entry(root, text = file_var, textvariable = file_var, font=('calibre',10,'normal'))
db_entry = tk.Entry(root, text = db_var, textvariable = db_var, font=('calibre',10,'normal'))
dis_entry = tk.Entry(root, text = dis_var, textvariable = dis_var, font=('calibre',10,'normal'))

#find file button
file_button =tk.Button(root, text='...',command=find_file)
db_button =tk.Button(root, text='...',command=find_db)

# creating a button using the widget
# Button that will call the function
run_button = tk.Button(root, text='Run',command=nearest_well)

# placing the label and entry in the required position using grid method
file_label.grid(row=1,column=0)
file_entry.grid(row=1,column=1)
file_button.grid(row=1,column=2)
db_label.grid(row=2,column=0)
db_entry.grid(row=2,column=1)
db_button.grid(row=2,column=2)
dis_label.grid(row=3,column=0)
dis_entry.grid(row=3,column=1)
run_button.grid(row=4,column=1)


# performing an infinite loop
# for the window to display
root.mainloop()
