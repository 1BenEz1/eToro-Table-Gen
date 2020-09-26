import openpyxl # To open xlsx files
import pandas as pd # To process the data
from tkinter import Tk 
from tkinter.filedialog import askopenfilename # Nice GUI
import dataframe_image as dfi # To save the DataFrame

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
path = askopenfilename() # show an "Open" dialog box and return the path to the selected file


obj = openpyxl.load_workbook(path) # The Excel file
ws = obj.active # The Excel work sheet

# Fetching the relevant info
table = []
for row in ws.iter_rows(10,20):
    r = []
    for cell in row[2:7]:
        r.append(cell.value)
    table.append(r)

# Creating the DataFrame
data = table[1:]
df = pd.DataFrame(data=data, columns = table[0])

# Method to dye values: RED if negative, GREEN elsewise
def color_green(x):
    if x < 0.0:
        color = 'red'
    else:
        color = 'green'
    return 'color: %s' % color

# Styling the DF to match the template
styled_df = df = df.style.set_table_styles(
                [{'selector':'th',
                 'props': [('background','#36304a'),
                           ('color','white'),
                           ('text-align','center')]},
                 {'selector':'th',
                   'props':[('max-width', '135px')]},

                 {'selector':'tr:nth-of-type(odd)',
                 'props': [('background', '#d9d9d9')]},

                 {'selector':'tr:nth-of-type(even)',
                 'props': [('background-color', '#f9f9f9')]},
                 ])\
        .set_properties(subset= ['Instument Name','Etoro Symbol'], **{'text-align': 'left'})\
        .set_properties(subset= ['#', '% Volume Above 20 Day Average', "Yesterday's % Change" ], **{'text-align': 'center'})\
        .format({'#': lambda x: '%.0f'%(x),
                '% Volume Above 20 Day Average': lambda x: '%.2f'%(x*100) +'%',
                "Yesterday's % Change": lambda x: '%.2f'%(x*100) +'%' })\
        .applymap(color_green, subset=pd.IndexSlice[:,["Yesterday's % Change"]])\
        .hide_index()

# Exporting the styled DF as image
dfi.export(styled_df,f'{path[:-5]}.png')