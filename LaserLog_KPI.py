# -*- coding: utf-8 -*-
"""
Created on Tue May 24 14:45:39 2022

    Test script for manipulating and visualising data collected from Laser Drill department for KPI purposes

@author: alexs, joshuaj
"""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import ticker, image
from matplotlib.offsetbox import (TextArea, AnnotationBbox, OffsetImage)
import mplcursors
import datetime as dt

"""
-------------------------
DATA MANIPULATION SECTION
-------------------------
"""

def convToTimeDelta(x):
    vals = x.split(':')
    vals = [int(val) for val in vals]
    out = dt.timedelta(minutes=vals[0], seconds=vals[1])
    return out.total_seconds()

def jobTotal(row, src_1, src_2, targ):
    row[targ] = int(row[src_1]) * int(row[src_2])
    return row

# Center align DataFrame column titles
pd.set_option('display.colheader_justify', 'center')

# Read data file into the .py
# combi_plog => CombiDrill production log
combi_plog = pd.read_excel('C:/Users/joshuaj/Laser_Log.xlsx',
                         parse_dates = ['Start-Date'])

combi_plog = combi_plog.dropna(how = 'all')

# Select the columns we care about for calculating a day total
combi_plog = combi_plog[['Start-Date', 'Work-Order', 'Panel-Count', 'Hole-Count', 'Drill-Time']]

# Rename the columns to something easier to type since they will likely be used often
rename_dict = {'Start-Date': 'Complete', 'Work-Order': 'WO', 'Panel-Count': 'Pnls', 'Drill-Time': 'Elapsed'}
combi_plog = combi_plog.rename(columns = rename_dict)

# Add a column for total holes in a job and a column for total time in a job
combi_plog['Th'] = None
combi_plog['Tt'] = None

# Apply the getSeconds and jobTotal functions to the DataFrame to store the job total holes and time elapsed in Th and Tt
combi_plog['TimeSec'] = combi_plog['Elapsed'].apply(convToTimeDelta)
combi_plog = combi_plog.apply(jobTotal, src_1 = 'Pnls', src_2 = 'Hole-Count', targ = 'Th', axis = 1)
combi_plog = combi_plog.apply(jobTotal, src_1 = 'Pnls', src_2 = 'TimeSec', targ = 'Tt', axis = 1)

# Group the production log DataFrame by date so that we can create a new DataFrame with aggregate data from the totals 
#   for each day
by_date = combi_plog.groupby(by = 'Complete')

# Create a new empty DataFrame for the daily aggregate data that we will be pulling from the combi_plog grouped by
#   completion date
aggr = pd.DataFrame(
    {
         'Date': [],
         'Panels': [],
         'Holes': [],
         'Time': [],
    }
)

# Convert certain columns to data types that we want them in
aggr['Panels'] = aggr['Panels'].astype(int)
aggr['Holes'] = aggr['Holes'].astype(int)
aggr['Time'] = aggr['Time'].astype(float)

# Create a Series that we can use within the aggregation loop to append to the aggr DataFrame
day_total = pd.Series(dtype = 'object')
day_total['Date'] = None
day_total['Panels'] = None
day_total['Holes'] = None
day_total['Time'] = None

# Keep the data from the previous row to check if the work order is the same so panels are not added multiple times
prev_row = pd.Series(dtype = 'object')

last_ind = -1

# Loop through the groups that are vertically delimited by completion date
for key, group in by_date:
    
    # Reset the daily totals for hole ct, drill time, and panel ct
    day_holes = 0
    day_time = 0
    day_pnls = 0
    
    last_ind += group.count()[0]
    
    # Loop through each row in the group to add up the hole ct, drill time, and panel ct per work order
    for index, row in group.iterrows():
        day_holes += row['Th']
        day_time += row['Tt']
        
        if(index > 0):
           if(prev_row['WO'] != row['WO']):
               day_pnls += row['Pnls']
        else:
            day_pnls = row['Pnls']
        
        prev_row = row
    
    # Debug prints
    #print('Job #: ' + str(count) + '\tDate: ' + str(key))
    #print('Holes: \t\t' + str(day_holes))
    #print('Drill time: ' + str(day_time))
    #print('Panels: \t' + str(day_pnls) + '\n')
    
    # Put values into the day_total Series then append the series to the aggr DataFrame
    day_total['Date'] = key
    day_total['Panels'] = day_pnls
    day_total['Holes'] = day_holes
    day_total['Time'] = day_time
    aggr = aggr.append(day_total, ignore_index = True)

# Logic by which to aggregate the resampled data
logic = {'Holes'  : 'sum',
         'Time'  : 'sum',
         'Panels'   : 'sum'}

#day_totals = aggr

#day_totals['H/M'] = day_totals['Holes'] / (day_totals['Time'] / 60)
#day_totals['H/M wo skive'] = day_totals['Holes'] / ((day_totals['Time'] - 60) / 60)

aggr = aggr.set_index('Date').resample('W-SUN').apply(logic).reset_index()

"""
----------------
PLOTTING SECTION
----------------
"""

# Set the graphical theme for the plot
plt.rcParams.update(plt.rcParamsDefault)
plt.style.use('seaborn-notebook')

# Function used to make the scatter point annotations look nicer
# 1_30_23 - changed x[sel.index] to x.iloc[sel.index] to supplement the lesser data
def vias_annotation(sel):
    sel.annotation.set(position = (-75, 115), anncoords = 'offset points', va = 'top', ha = 'left', fontsize = 16)
    sel.annotation.set_text('Date: {}\nμVias: {:,}'.format(x.iloc[sel.index].strftime('%B %#d'), int(sel.target[1])))
    sel.annotation.arrow_patch.set(arrowstyle = '-|>', color = 'black', linewidth = 2)
    
    bb = sel.annotation.get_bbox_patch()
    bb.set(fc = 'lightsteelblue', ec = 'black', alpha = 1, lw = 2)
    bb.set_boxstyle('round', pad = 0.8)
    
def pnls_annotation(sel):
    sel.annotation.set(position = (-15, 75), anncoords = 'offset points', va = 'top', ha = 'left', fontsize = 16)
    sel.annotation.set_text('Date: {}\nPanels: {:,}'.format(x.iloc[sel.index].strftime('%B %#d'), int(sel.target[1])))
    sel.annotation.arrow_patch.set(arrowstyle = '-|>', color = 'black', linewidth = 2)
    
    bb = sel.annotation.get_bbox_patch()
    bb.set(fc = 'lightsteelblue', ec = 'black', alpha = 1, lw = 2)
    bb.set_boxstyle('round', pad = 0.8)

# Set x, x_fmtd, y, and v to the aggregate data
# 1_30_22 - added [-16:] otherwise all data will be graphed
x = aggr['Date'][-16:]
x_fmtd = x.dt.strftime('%m-%d')[-16:]
y = aggr['Holes'][-16:]
v = aggr['Panels'][-16:]

# Create a figure that contains a subplot with 2 axes
fig, (ax1, ax2) = plt.subplots(2, figsize = (17, 11), dpi = 80, sharex = True)
fig.canvas.manager.set_window_title('KPI Laser Drill {}'.format(dt.date.today()))

plt.subplots_adjust(hspace = 0.1)

# Add the Somacis logo to the figure
logo = image.imread('C:/Users/joshuaj/Logo.png', format = 'png')
imgbox = OffsetImage(logo, zoom = 0.5)
imbbox = AnnotationBbox(imgbox,
                        xy = (0.05, 0.035),
                        xycoords = ('axes fraction', 'figure fraction'))

ax2.add_artist(imbbox)



# Plot both the scatter points and line chart for the x, y data so that we can select certain parts of the graph
scatter = ax1.scatter(x_fmtd, y, color = 'cornflowerblue', alpha = 1, zorder = 4)

line = ax1.plot_date(x_fmtd, y,
                     linewidth = 2,
                     linestyle = '-',
                     color = 'cornflowerblue',
                     label = 'μVias Drilled')

otc_vias = ax1.plot(x_fmtd, [1440000] * y.size,
                              linewidth = 2,
                              linestyle = '--',
                              color = 'green',
                              label = 'Expected',
                              alpha = 0.8)

average_vias = ax1.plot(x_fmtd, [y.mean()] * y.size,
                        linewidth = 2,
                        linestyle = '--',
                        color = 'orange',
                        label = 'Average',
                        alpha = 0.8)

otc_panels = ax2.plot(x_fmtd, [70] * v.size,
                              linewidth = 2,
                              linestyle = '--',
                              color = 'green',
                              label = 'Expected',
                              alpha = 0.8)

average_panels = ax2.plot(x_fmtd, [v.mean()] * v.size,
                        linewidth = 2,
                        linestyle = '--',
                        color = 'orange',
                        label = 'Average',
                        alpha = 0.8)

ax1.legend(loc = 'upper left', facecolor = 'lightsteelblue')

bar = ax2.bar(x_fmtd, v, width = .35, color = 'cornflowerblue', label = 'v')

# Rotate each x-axis label to make it more readable
for item in ax2.get_xticklabels():
    item.set_rotation(30)

ax1.set(ylim = (0, max(y) + 100000))
ax2.set(ylim = (0, max(v) + 20))

fig.suptitle('KPI {} - Laser Drill'.format(dt.date.today().year), fontsize = 20)

ax1.set_title('Microvia Throughput', size = 16)
ax2.set_title('Panel Throughput', size = 16)
ax1.set_ylabel('μVia Count', size = 16)
ax2.set_xlabel('Week\n(Numbers reflect values up to label date)', size = 16)
ax2.set_ylabel('Panel Count', size = 16)

ax1.grid(which = 'Major', axis = 'both')
ax2.grid(which = 'Major', axis = 'y')
ax2.set_axisbelow(True)

for c in ax2.containers:
    ax2.bar_label(c, fontsize = 12)

mplcursors.cursor(scatter, hover = True).connect('add', vias_annotation)
mplcursors.cursor(bar, hover = True).connect('add', pnls_annotation)

# Formatting function to make y-values show K and M for thousand and million
mkfunc = lambda x, pos: '%1.1fM' % (x * 1e-6) if x >= 1e6 else '%1.1fK' % (x * 1e-3) if x >= 1e3 else '%1.1f' % x
mkformatter = ticker.FuncFormatter(mkfunc)
ax1.yaxis.set_major_formatter(mkformatter)




# Create an annotation that includes the instructions for annotating the scatter points
text = 'Hover over a scatter point to display exact data.\nRight click annotation to remove.'
offsetbox = TextArea(text, textprops = dict(fontsize = 10))

instructions_abb = AnnotationBbox(offsetbox,
                                  xy = (0.1, 0.905),
                                  xycoords = ('axes fraction', 'figure fraction'),
                                  bboxprops = dict(fc = 'lightsteelblue',
                                                   ec = 'black',
                                                   alpha = 1,
                                                   lw = 2,
                                                   boxstyle = 'round'))

ax1.add_artist(instructions_abb)




text3 = 'Date viewed: ' + dt.date.today().strftime('%B %#d \'%y')
ax1.annotate(text3, xy = (0.8, 0.895), xycoords = 'figure fraction', fontsize = 12)



# Show all of the plot elements added thus far
plt.show() 
