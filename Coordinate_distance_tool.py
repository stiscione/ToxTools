import tkinter as tk
import xlrd
import xlsxwriter
import pandas as pd
from math import sin, cos, sqrt, atan2, radians
from tkinter import filedialog
import warnings
warnings.filterwarnings('ignore')

root= tk.Tk()
#name title of program
root.title("Nearby Well Distance Tool")

#declare file name string var
file1_var=tk.StringVar()
file2_var=tk.StringVar()
dis_var=tk.IntVar()

# setting the windows size
root.geometry("600x200")


#cas lookup function
def distance_measure ():
    inputfile1 = file1_var.get()
    inputfile2 = file2_var.get()
    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
    wb1 = xlrd.open_workbook(inputfile1)
    wb2 = xlrd.open_workbook(inputfile2)
    sheet1 = wb1.sheet_by_index(0)
    sheet2 = wb2.sheet_by_index(0)
    distance_threshold = dis_var.get()

    # approximate radius of earth in miles
    R = 3959.999
    header_row1 = True
    for row1 in range(sheet1.nrows):
        #skip header row
        if header_row1 is True:
            header_row1 = False
            continue
        site1_id = sheet1.cell_value(row1,0)
        site1_info = sheet1.cell_value(row1,1)
        site1_lat = sheet1.cell_value(row1,2)
        site1_long = sheet1.cell_value(row1,3)
        site1_lat_r = radians(site1_lat)
        site1_long_r = radians(site1_long)

        df = pd.DataFrame(columns = ['identifier1', 'identifier2', 'lat', 'long','direction','distance'])

        #iterate through all rows in input 2 to look for distances under threshold to output
        header_row2 = True
        for row2 in range(sheet2.nrows):
            #skip header row
            if header_row2 is True:
                header_row2 = False
                continue
            identifier1 = sheet2.cell_value(row2,0)
            identifier2  = sheet2.cell_value(row2,1)
            if sheet2.cell_value(row2,2)=="" or sheet2.cell_value(row2,2)==None or sheet2.cell_value(row2,3) == "" or sheet2.cell_value(row2,3) == None:
                continue
            site2_lat = float(sheet2.cell_value(row2,2))
            site2_long = float(sheet2.cell_value(row2,3))
            site2_lat_output = sheet2.cell_value(row2,2)
            site2_long_output = sheet2.cell_value(row2,3)
            if site1_lat > site2_lat:
                north_south = "south"
            else:
                north_south  = "north"
            if site1_long > site2_long:
                east_west = "west"
            else:
                east_west = "east"
            direction = north_south + '-' + east_west
            site2_lat_r = radians(site2_lat)
            site2_long_r = radians(site2_long)
            dlon = site2_long_r - site1_long_r
            dlat = site2_lat_r - site1_lat_r
            a = sin(dlat / 2)**2 + cos(site2_lat_r) * cos(site1_lat_r) * sin(dlon / 2)**2
            c = 2 * atan2(sqrt(a), sqrt(1 - a))
            distance = R * c
            if distance <= float(distance_threshold):
                data = pd.DataFrame([[identifier1, identifier2, site2_lat_output, site2_long_output, direction,distance]], columns = ['identifier1', 'identifier2', 'lat', 'long','direction','distance'])
                df = df.append(data)
            else:
                continue
        df.to_excel (writer,sheet_name=site1_id[0:10], index=False)
    writer.save()

def find_file1 ():
    filetypes = (
        ('excel files', '*.xls'),
        ('All files', '*.*')
    )
    file1_var = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    file1_entry.insert(0, file1_var)

def find_file2 ():
    filetypes = (
        ('excel files', '*.xls'),
        ('All files', '*.*')
    )
    file2_var = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    file2_entry.insert(0, file2_var)

#label for entry fields
file1_label = tk.Label(root, text = 'Input 1', font=('calibre',10, 'bold'))
text_label = tk.Label(root, text = '----->', font=('calibre',10, 'bold'))
file2_label = tk.Label(root, text = 'Input 2', font=('calibre',10,'bold'))
dis_label = tk.Label(root, text = 'distance (miles)', font=('calibre',10,'bold'))

#entry space for file1 and file2
file1_entry = tk.Entry(root, text = file1_var, textvariable = file1_var, font=('calibre',10,'normal'))
file2_entry = tk.Entry(root, text = file2_var, textvariable = file2_var, font=('calibre',10,'normal'))
dis_entry = tk.Entry(root, text = dis_var, textvariable = dis_var, font=('calibre',10,'normal'))

#find file button
file1_button =tk.Button(root, text='...',command=find_file1)
file2_button =tk.Button(root, text='...',command=find_file2)

# creating a button using the widget
# Button that will call the function
run_button = tk.Button(root, text='Run',command=distance_measure)

# placing the label and entry in the required position using grid method
file1_label.grid(row=1,column=0)
file1_entry.grid(row=1,column=1)
file1_button.grid(row=1,column=2)
text_label.grid(row=1,column=3)
file2_label.grid(row=1,column=4)
file2_entry.grid(row=1,column=5)
file2_button.grid(row=1,column=6)
dis_label.grid(row=3,column=0)
dis_entry.grid(row=3,column=1)
run_button.grid(row=4,column=3)


# performing an infinite loop
# for the window to display
root.mainloop()
