#!/usr/bin/python3

# MIT License
# 
# Copyright (c) 2017 Josh Simonot
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# 

import sys
import datetime
import plotly
import plotly.graph_objs as go
from plotly.graph_objs import *

import locale
locale.setlocale( locale.LC_ALL, '' )

#=============================
def printHelpInfo():
	print( )
	print( "build-dashboard" )
	print( "  -> reads account data and builds an analysis dashboard" )
	print( )
	print( "Command-line syntax:" )
	print( "  python3 build-dashboard.py [data_file] [color_theme]")
	print( )
	print( "Color theme options:")
	print( "  -> Light, Dark, Blue, Umber")
	print( )
	print( "Data file:")
	print( "  -> expecting an Excel file (xls/xlsx), with tabular data in the following format:")
	print( )
	print( "                  A       |  B   |       C       | ... |       N       |")
	print( "          ---------------------------------------------------------------")
	print( "    ROW1:|                |      |               |     |               |")
	print( "    ROW2:|                |      |    JAN_YEAR   | ... |    DEC_YEAR   |")
	print( "    ROW3:| ACCOUNT_1_NAME |      | ACCOUNT_VALUE | ... | ACCOUNT_VALUE |")
	print( "    ROW4:|                |      | CONTRIBUTIONS | ... | CONTRIBUTIONS |")
	print( "    ROW5:| ACCOUNT_2_NAME |      | ACCOUNT_VALUE | ... | ACCOUNT_VALUE |")
	print( "    ROW6:|                |      | CONTRIBUTIONS | ... | CONTRIBUTIONS |")
	print( "    ROW5:| ACCOUNT_3_NAME |      | ACCOUNT_VALUE | ... | ACCOUNT_VALUE |")
	print( "    ROW6:|                |      | CONTRIBUTIONS | ... | CONTRIBUTIONS |")
	print( "    ...etc...")
	print( )
	print( "  **NOTE: ACCOUNT_VALUE is the Market Value at the end of month,")
	print( "          CONTRIBUTIONS are the change in Book Value over the course of the month.")
	print( )
	print( "  Where the sheet name is YEAR (eg, 2015, 2016, 2017, etc...)")
	print( "  There must also be a sheet named 'Info' that lists all account names and specifies a start and end date:")
	print( )
	print( "             A  |        B       |   C  |    D   |      E     |")
	print( "          ---------------------------------------------------------------")
	print( "    ROW1:|      |                |      |        |            |")
	print( "    ROW2:|      |                |      |        |            |")
	print( "    ROW3:|      | ACCOUNT_1_NAME |      | Start: | START_DATE |")
	print( "    ROW4:|      | ACCOUNT_2_NAME |      |  End:  |  END_DATE  |")
	print( "    ROW5:|      | ACCOUNT_3_NAME |      |        |            |")
	print( "    ROW6:|      | ACCOUNT_4_NAME |      |        |            |")
	print( "    ROW7:|      | ACCOUNT_5_NAME |      |        |            |")
	print( "    ...etc...")
	
#=============================
if len(sys.argv) < 2: 
	print( "command-line argument required: filepath to investment data file" )
	printHelpInfo()
	exit()

filepath = sys.argv[1]

if ".xls" not in filepath:
	print( "Unexpected file format: expecting an Excel (.xls/.xlsx) document" )
	printHelpInfo()
	exit()
	
colourTheme = "light"
if len(sys.argv) >= 3:
	colourTheme = sys.argv[2].lower()
	
if "umber" == colourTheme:
	pageColour = '#3C3333'
	labelColour = '#C1AC8D'
	borderColour = '#222222'
	
elif "blue" == colourTheme:
	pageColour = '#D6F3FF'
	labelColour = '#00163D'
	borderColour = 'AAAAAA'
	
elif "dark" == colourTheme:
	pageColour = '#222222'
	labelColour = '#C1AC8D'
	borderColour = '#262626'

else: #default: light
	pageColour = '#EDEEF4'
	labelColour = '#3C0000'
	borderColour = 'DCDDE3'

#read data from excel document (see script: utils.py)
from utils import loadDataFromExcel
data = loadDataFromExcel(filepath)

#calculate secondary data (per account monthly gains, and growth rate)
accounts = data["accounts"]
for acctname, accountdata in accounts.items():
	gains = []
	growth = []
	value = accountdata["value"]
	cw = accountdata["cw"]
	
	accountdata["gains"] = gains
	accountdata["growth"] = growth
	
	#first recorded month
	gains.append( 0 )
	growth.append(0)
	
	for date in range( 1, len(data["dateRange"]) ):
		gains.append( value[date] - value[date-1] - cw[date] )
		
		#equation assumes all contributions are made in the middle of the month:
		div = value[date-1] + 0.5*cw[date]
		growth.append(0 if div==0 else (gains[date] / div) )

#calculate account aggregates (total value, total cw, total gains, weighted average growth)
aggregates = {}
data["aggregates"] = aggregates

n = len(data["dateRange"])
totvalue = [0]*n
totcw = [0]*n
totgains = [0]*n
wavgrowth = [0]*n  #(weighted average growth)

aggregates["totvalue"] = totvalue
aggregates["totcw"] = totcw
aggregates["totgains"] = totgains
aggregates["wavgrowth"] = wavgrowth

for acctname, accountdata in accounts.items():
	for date in range( 0, len(data["dateRange"]) ):
		totvalue[date] += accountdata["value"][date]
		totcw[date] += accountdata["cw"][date]
		totgains[date] += accountdata["gains"][date]
		wavgrowth[date] += accountdata["growth"][date] * accountdata["value"][date]
		
for date in range( 0, len(data["dateRange"]) ):
	wavgrowth[date] = 0 if totvalue[date]==0 else (wavgrowth[date]/totvalue[date])

#calculate cumulatives (cumulative cw, gains)
cumulatives = {}
data["cumulatives"] = cumulatives

n = len(data["dateRange"])
cumcw = [0]*n
cumgains = [0]*n
cumcwgains = [0]*n

cumulatives["cw"] = cumcw
cumulatives["gains"] = cumgains
cumulatives["cw+gains"] = cumcwgains

cumcw[0] = totcw[0]
cumgains[0] = totgains[0]
cumcwgains[0] = cumcw[0] + cumgains[0]

for date in range( 1, len(data["dateRange"]) ):
	cumcw[date] = cumcw[date-1] + totcw[date]
	cumgains[date] = cumgains[date-1] + totgains[date]
	cumcwgains[date] = cumcw[date] + cumgains[date]

#==================================
graph_title = "Investment Overview"
layout = Layout(
	title = graph_title,
    paper_bgcolor=pageColour,
    plot_bgcolor='rgba(0,0,0,0)',
	font=dict(color=labelColour),
	xaxis=dict( showgrid=True, gridcolor='#111111' ),
	yaxis=dict( showgrid=True, range=[0, max(totvalue)] ),
)

cw_text = [ locale.currency( y, grouping=True ) for y in cumcw]
gains_text = [ locale.currency( y, grouping=True ) for y in cumgains]
val_text = [ locale.currency( y, grouping=True ) for y in totvalue]

plotdata = []
plotdata.append( go.Scatter(x=data["dateRange"], y=cumcw ,name="Contributions (since {})".format( data["dateRange"][0].strftime("%b %Y") ),fill='tonexty',text=cw_text, hoverinfo='text+name') ) 
plotdata.append( go.Scatter(x=data["dateRange"], y=cumcwgains ,name="Gains (since {})".format( data["dateRange"][0].strftime("%b %Y")),fill='tonexty',text=gains_text, hoverinfo='text+name') ) 
plotdata.append( go.Scatter(x=data["dateRange"], y=totvalue ,name="Total Value",fill='tonexty',text=val_text, hoverinfo='text+name') ) 

fig = Figure(data=plotdata, layout=layout)	

overview_graph = "investment-overview.html"
plotly.offline.plot(fig, filename=overview_graph, auto_open=False, show_link=False)

#==================================
graph_title = "Account Distribution"
layout = Layout(
	title = graph_title,
    paper_bgcolor=pageColour,
    plot_bgcolor='rgba(0,0,0,0)',
	font=dict(color=labelColour),
	xaxis=dict( showgrid=False ),
	yaxis=dict( showgrid=False ),
)

plotdata = []
text = []
for acctname, accountdata in accounts.items():
	text = "{0:.2f}%".format( 100*accountdata["value"][-1] / totvalue[-1] )
	if accountdata["value"][-1] != 0:
		plotdata.append( go.Bar(x=[' '], y=[accountdata["value"][-1]] ,name=acctname, text=[text], hoverinfo='text+name',) ) 	

fig = Figure(data=plotdata, layout=layout)	

dist_graph = "account-distribution.html"
plotly.offline.plot(fig, filename=dist_graph, auto_open=False, show_link=False)

#==================================
graph_title = "Growth Rates"
layout = Layout(
	title = graph_title,
    paper_bgcolor=pageColour,
    plot_bgcolor='rgba(0,0,0,0)',
	font=dict(color=labelColour),
	xaxis=dict( showgrid=True, gridcolor='#111111' ),
	yaxis=dict( showgrid=True ),
)

plotdata = []
for acctname, accountdata in accounts.items():
	text = [ "{0:.4f}%".format(100*y) for y in accountdata["growth"]]
	plotdata.append( go.Scatter(x=data["dateRange"], y=accountdata["growth"] ,name=acctname, text=text, hoverinfo='text+name') ) 

acctname = "Weighted Average"
text = [ "{0:.4f}%".format(100*y) for y in aggregates["wavgrowth"]]
plotdata.append( go.Scatter(x=data["dateRange"], y=aggregates["wavgrowth"] ,name=acctname, text=text, hoverinfo='text+name',line = dict(
        color = ('rgb(245, 50, 50)'),
        width = 6,
        dash = 'dash',
		shape='spline')) ) 	
	
fig = Figure(data=plotdata, layout=layout)	

rates_graph = "account-growth.html"
plotly.offline.plot(fig, filename=rates_graph, auto_open=False, show_link=False)


#see script: utils.py
from utils import builddashboard
dashboard_filepath = builddashboard("dashboard.html", overview_graph, dist_graph, rates_graph, borderColour)

import webbrowser
webbrowser.open(dashboard_filepath, new=1)


