from matplotlib import patheffects
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import (MultipleLocator)
from matplotlib import colors
import numpy as np
from fpdf import FPDF # fpdf2
import os

SetColors = ['#d6afd2','#f2cadf', '#936daf', '#d6afd2','#f2cadf', '#936daf', '#f2cadf']
BaskFont = r'D:\Python\fonts\arial.ttf'
CalFont = BaskFont
NameCol = 'Names'                                                            # Variable to store the name column
file = ("2 green end term 1 2025")

# Getting data from filename to be used in pdf
grade, gcolor, session, term, termNum, year = file.split()
gcolor = gcolor.capitalize()
session = session.capitalize()
term = term.capitalize()
grade = grade + ' ' + gcolor
term = term + ' ' + termNum
examInformation = f"{grade} {session} {term} {year}"
examInformationseparated = grade, session, term, year
file = f'src\{file}.xlsx'

# Setting up the directories to store the final pdf's and the graphs used in them
Pdfdir = f'{examInformation} PDFs\\'
Datadir = f'{examInformation} Data\\'
if not os.path.exists(Pdfdir): os.makedirs(Pdfdir)       # Create a new directory because it does not exist 
if not os.path.exists(Datadir): os.makedirs(Datadir)     # Create a new directory because it does not exist 

# Convert excel file into a pandas database and clean it
data = pd.read_excel(file)                                                  # All data from the original excel spreadsheet
data.columns = data.columns.str.strip()                                     # Remove whitespace from column names
cleanData = data
NumStudents = len(cleanData.iloc[:,0])
subList = list(data.columns.str.strip())
subList = [subject.capitalize() for subject in subList]
cleanData.columns = subList
subList = subList[1:] + ['Total Marks', 'Average']

def hex_to_rgb(value):                                                      # Function to convert a hex value to rgb
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))
Purple1 = hex_to_rgb('#f2cadf')

# Calculating student total, average and ACM
for i in range(NumStudents) :
    SplitFormatedNames = cleanData.iloc[i,0]                                            # Get student names on column 0 row i
    cleanData = cleanData.set_index(NameCol)                                   # Setting the name column as the index
    
    # Calculating total marks and average for a student
    cleanData.loc[SplitFormatedNames, 'Total Marks'] = cleanData[subList[:-2]].iloc[i].sum()                              # Sum the valuesof specific subjects in a row to get the total marks and insert the result in the total marks column
    cleanData.loc[SplitFormatedNames, 'Average'] = round(cleanData[subList[:-2]].iloc[i].mean())                          # same as above but for mean
    cleanData = cleanData.reset_index(NameCol)                                                                  # Re-setting index

refreceValues = cleanData[subList].iloc[0,]
cleanData = cleanData.drop(cleanData.index[0])                                        # Remove first "OUT OFF" row
cleanData = cleanData.sort_values(by="Total Marks", ascending=False)
cleanData['Position'] = list(range(1,NumStudents))
NumStudents -= 1
studentData = cleanData[subList + ['Position']]                                                                            # Dataframe containing only the marks for each student

# Finding the class average for each subject
SubjectAverages = []
for i in subList:
    SubjectAverages.append(round(cleanData[i].mean(),1))
SubjectAverages = pd.Series(SubjectAverages, index=subList)

# Turning the student data int 3 columns subject, score and frequency the purpose is to make a dataframe for the scatter plot
FrequencyTable = pd.DataFrame()
for column in range(len(subList[:-2])) :
    StudentDataSorted = sorted(studentData[subList[column]])
    freak = studentData[subList[column]].value_counts()
    for row in range(NumStudents):
        val = 26 * column
        FrequencyTable.loc[val+row,'Subject'] = subList[column]
        FrequencyTable.loc[val+row,'Score'] = studentData.iloc[row,column]
        FrequencyTable.loc[val+row,'freak'] = freak[studentData.iloc[row,column]]

plt.rcParams['font.size'] = 40

# Creating charts
def CreateBarChart(XAxisLables, singleStudentDataXaxis, ClassDataXaxis, exportFoldernameBarChart, UniqueFileName, chartsize=1):    # Create bar chart
    Figure, axis = plt.subplots(figsize=(32, 21), dpi=40)
    # Figure.patch.set_alpha(0.5)

    #Bar chart style------
    axis.grid(True, linestyle='solid', color = "white")     # Show grid behind bars
    axis.set_facecolor('#f2cadf')                           # Set plot backgroud
    plt.rcParams['font.size'] = 40
    # axis.patch.set_alpha(0.6)                           

    # Bar chart customization
    ticks = np.arange(0,101,5)                                     # Setting Y axis values to start from 0 increase in steps of five and end at 100
    axis.set_ylim(0,100)         
    axis.set_yticks(ticks)
    axis.set_yticklabels(ticks, rotation=0, ha='right', fontsize=37)
    axis.xaxis.set_minor_locator(MultipleLocator(1))
    axis.tick_params('both', length=30, width=1, which='major')
    axis.set_axisbelow(True)

    # Plot Bar chart------------------------------------------------------------
    if chartsize == 2 : 
        axis.bar(XAxisLables, singleStudentDataXaxis, width=-0.4, align='edge', label='Student marks', color = "#936daf")    # Students Data
        axis.bar(XAxisLables, ClassDataXaxis, width=0.4, align='edge', label='Class avearge', color = "#d6afd2")   # Class average
    if chartsize == 1 : 
        axis.bar(XAxisLables, singleStudentDataXaxis, label='Marks', color = "#936daf")    # Students Data
    axis.legend(bbox_to_anchor = (1.01,1.15), loc='upper right')                                                    # Legend

    # Figure.savefig(f'{exportFoldernameBarChart}\\{UniqueFileName} bar.png', facecolor=Figure.get_facecolor())    # Save bar chart
    Figure.savefig(f'{exportFoldernameBarChart}\\{UniqueFileName} bar.png', bbox_inches='tight', pad_inches = 0)    # Save bar chart

def CreatePieChartBSH(dataForPiechart, Colors, exportFolderPieChart, UniqueFileName):  # Create a pie chart with the best subject highlighted
    displayData = []
    colorsMost = Colors
    colorsBest = ['none'] * len(dataForPiechart)
    plt.rcParams['font.size'] = 24

    # Find and highlight best subject
    expd = [0] * len(dataForPiechart)
    BestSubjectPositon = np.array(dataForPiechart).argmax()
    expd[BestSubjectPositon] = -0.05
    colorsBest[BestSubjectPositon] = Colors[BestSubjectPositon]
    # colorsMost[BestSubjectPositon] = 'none'

    # Customize anotated boxes
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="<|-",facecolor='black', linewidth=2), bbox=bbox_props, zorder=0, va="center")

    # Create pie chart
    Figure1, ax =  plt.subplots(figsize=(14, 7), dpi = 70)
    wedges, text = ax.pie(dataForPiechart, wedgeprops=dict(width=0.5), startangle=90, colors=colorsMost, counterclock=False, shadow=False)
    
    # Annotating wedges--------
    for i, p in enumerate(wedges):
        # Getting the Arc midpoint of each wedge
        ang = (p.theta2 - p.theta1)/2. + p.theta1   # Wedge midpoint angle
        y = np.sin(np.deg2rad(ang))                 # Calculating the y coordinats of arc midpoint
        x = np.cos(np.deg2rad(ang))                 # Calculating the x coordinats of arc midpoint

        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        connectionstyle = f"angle, angleA=0, angleB={ang}"
        kw["arrowprops"].update({"connectionstyle": connectionstyle})
        # Annotating wedges that are not the best subject
        if not(colorsBest[i] == 'none'):
            x *= 1.05
            y *= 1.05
        displayData.insert(i, f'{dataForPiechart.index[i]}: {dataForPiechart.iloc[i]} marks\n{round(dataForPiechart.iloc[i]/dataForPiechart.sum()*100,1)}%')  # Box data
        ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y), horizontalalignment=horizontalalignment, **kw)              # Create annotaion box 

    # Save the charts
    Figure1.savefig(f'{exportFolderPieChart}\\{UniqueFileName} pie.png', transparent = True)     # Pie chart

def CreatePieChartLMIS(dataForPiechart, Colors, exportFolderPieChart, UniqueFileName):  # Create a pie chart with the lost marks indicated separate
    displayData = []
    colorsMost = []
    lostmarkscolor = '#c23616'
    datawithLostmarks = []
    colorsBest = ['none'] * len(dataForPiechart)
    plt.rcParams['font.size'] = 24

    # Find and highlight best subject
    expd = [0] * len(dataForPiechart)
    BestSubjectPositon = np.array(dataForPiechart).argmax()
    expd[BestSubjectPositon] = -0.05
    colorsBest[BestSubjectPositon] = Colors[BestSubjectPositon]

    # Lost marks with subject
    for j, i in enumerate(dataForPiechart):
        datawithLostmarks.extend([i, 100-i])
        colorsMost.extend([Colors[j], lostmarkscolor])

    # Create pie chart
    Figure1, ax =  plt.subplots(figsize=(14, 7), dpi = 70)
    wedges, text = ax.pie(datawithLostmarks, wedgeprops=dict(width=0.5), startangle=90, colors=colorsMost, counterclock=False, shadow=False)
    
    # Customize chart
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="<|-",facecolor='black', linewidth=2), bbox=bbox_props, zorder=0, va="center")

    # Annotating wedges--------
    for i, p in enumerate(wedges):
        if not(colorsMost[i] == lostmarkscolor):
            # Getting the Arc midpoint of each wedge
            i /= 2
            i = int(i)
            ang = (p.theta2 - p.theta1)/2. + p.theta1   # Wedge midpoint angle
            y = np.sin(np.deg2rad(ang))                 # Calculating the y coordinats of arc midpoint
            x = np.cos(np.deg2rad(ang))                 # Calculating the x coordinats of arc midpoint

            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = f"angle, angleA=0, angleB={ang}"
            kw["arrowprops"].update({"connectionstyle": connectionstyle})
            # Annotating wedges that are not the best subject
            # if not(colorsBest[i] == 'none'):
            #     x *= 1.05
            #     y *= 1.05
            displayData.insert(i, f'{dataForPiechart.index[i]}: {dataForPiechart.iloc[i]} marks\n{round(dataForPiechart.iloc[i]/dataForPiechart.sum()*100,1)}%')  # Box data
            print(x,y)
            print(1.35*np.sign(x), 1.3*y)
            ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y), horizontalalignment=horizontalalignment, **kw)              # Create annotaion box 

    # Save the charts
    Figure1.savefig(f'{exportFolderPieChart}\\{UniqueFileName} pie.png', bbox_inches='tight', pad_inches = 0, transparent = True)     # Pie chart

def CreatePieChartLMIT(dataForPiechart, maxScore, chartColors, exportFolderPieChart, UniqueFileName):  # Create a pie chart with the lost marks incated together LMIT
    displayData = []
    colorsMost = []
    lostmarkscolor = '#c23616'
    datawithLostmarks = []
    plt.rcParams['font.size'] = 24

    datawithLostmarks.extend(dataForPiechart)
    datawithLostmarks.append(maxScore-dataForPiechart.sum())
    colorsMost.extend(chartColors[:len(dataForPiechart)])
    colorsMost.append(lostmarkscolor)

    gap = 0.3
    usedAngle = 90
    multiplier = 1.35
    lablexlocation = 3
    extend = False
    if dataForPiechart.sum() < maxScore*3/10 : usedAngle = ((dataForPiechart.sum()/maxScore)*90/0.5)+3; multiplier = 2.5
    labelyrange = [round(gap*x*-1,1) for x in range(-round(len(dataForPiechart)/2), round(len(dataForPiechart)/2))]
    
    # Create pie chart
    Figure1, ax =  plt.subplots(figsize=(16, 7), dpi = 70)
    wedges, text = ax.pie(datawithLostmarks, startangle=usedAngle, colors=colorsMost, counterclock=False, shadow=False)
    
    # Customize chart
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="-",facecolor='black', linewidth=2), bbox=bbox_props, zorder=0, va="center")
    if extend :ax.set_xlim(-1,4); ax.set_ylim(-1.5,1.5)

    # Annotating wedges--------
    for i, p in enumerate(wedges):
        if not(colorsMost[i] == lostmarkscolor):
            # Getting the Arc midpoint of each wedge
            ang = (p.theta2 - p.theta1)/2. + p.theta1   # Wedge midpoint angle
            y = np.sin(np.deg2rad(ang))                 # Calculating the y coordinats of wedge midpoint
            x = np.cos(np.deg2rad(ang))                 # Calculating the x coordinats of wedge midpoint
            curry = 1.5*(y)
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = f"angle, angleA=0, angleB={ang} "
            kw["arrowprops"].update({"connectionstyle": connectionstyle})
            displayData.insert(i, f'{dataForPiechart.index[i]}: {dataForPiechart.iloc[i]} marks')  # Box data

            if extend:
                ax.annotate('', xy=(x,y), xytext=(multiplier-x, curry), horizontalalignment="left", **kw)              # Create annotaion box 
                connectionstyle = f"angle, angleA=0, angleB={0-ang}"
                kw["arrowprops"].update({"connectionstyle": connectionstyle})
                ax.annotate(displayData[i], xy=(multiplier-x, curry), xytext=(lablexlocation, labelyrange[i]), horizontalalignment="left", **kw)
                continue
            ax.annotate(displayData[i], xy=(x,y), xytext=(multiplier*np.sign(x), curry), horizontalalignment=horizontalalignment, **kw)              # Create annotaion box 
    plt.tight_layout()
    # Save the charts
    Figure1.savefig(f'{exportFolderPieChart}\\{UniqueFileName} pie.png', transparent = True)     # Pie chart

def CreateScatterplot(scatterTableToPlot, exportFoldernameScatter, UniqueFileName):
    Figure, axes =  plt.subplots(figsize=(30, 20), dpi=60)

    # Style scatter plot
    plt.rcParams['font.size'] = 40
    axes.grid(True, linestyle='solid', color = "black")
    axes.set_axisbelow(True)
    axes.yaxis.set_minor_locator(MultipleLocator(1))
    axes.set_title("Spread of marks per subject size indicates frequency")
    axes.set_xlabel("Subjects",fontsize=30)
    axes.set_ylabel("Marks",fontsize=30)

    # axes.bar(subList[:-2], SubjectAverages[:-2], width=0.4, align='edge', label='Class avearge', color = "#d6afd2")   # Class average
    # axes.plot(subList[:-2], SubjectAverages[:-2],"_", lw = 100)
    axes.scatter(scatterTableToPlot['Subject'],scatterTableToPlot['Score'],  s=((scatterTableToPlot['freak']+1)*8)**2,  c=-(scatterTableToPlot['freak']-18)**6) # Plot scatter plot
    axes.scatter(subList[:-2], SubjectAverages[:-2], s = 6000, c="red", lw=5, marker='_', label="Class Average")
    plt.tight_layout()

    Figure.savefig(f'{exportFoldernameScatter}\\{UniqueFileName} scat.png', transparent = True)                 # Save scatter plot
    # Figure.savefig(f'{exportFoldernameScatter}\\{UniqueFileName} scat.png', bbox_inches='tight', pad_inches = 0, transparent = True)                 # Save scatter plot

def CreateHistogram(HistogramData, exportFoldernameHistogram, UniqueFileName):
    figt4,axes = plt.subplots(figsize= (25,15), dpi=70)

    # Variables
    HisHeight = round(HistogramData.value_counts().max() + 1)
    n_bins = HistogramData.value_counts().sort_index().index[:]
    xlabledata = [round(lable) for lable in n_bins]
    xticcks = np.arange(0,101,2)
    yticcks = np.arange(0,HisHeight,1)
    
    # Style
    plt.rcParams['font.size'] = 30
    axes.set_title("frequency of marks across all subjects")
    axes.set_xlabel("Marks")
    axes.set_ylabel("Frequency")
    axes.grid(True, linestyle='solid', color = "lightgrey")
    axes.set_axisbelow(True)
    axes.set_ylim(0,HisHeight)
    axes.set_yticks(yticcks)
    axes.set_yticklabels(yticcks, rotation=0, ha='right', fontsize=20)
    axes.yaxis.set_minor_locator(MultipleLocator(1))
    axes.set_xticks(xticcks)
    axes.set_xlim(xlabledata[0]-1,xlabledata[-1]+1)
    axes.set_xticklabels(xticcks, rotation=0, ha='center', fontsize=20)

    N, bins, patches = axes.hist(HistogramData, bins = n_bins)
    
    # Setting color
    fracs = ((N**(1 / 2)) / N.max())
    norm = colors.Normalize(fracs.min(), fracs.max())
    for thisfrac, thispatch in zip(fracs, patches):
        color = plt.cm.viridis(norm(thisfrac))
        thispatch.set_facecolor(color)
    plt.tight_layout()

    figt4.savefig(f'{exportFoldernameHistogram}\\{UniqueFileName} histogram.png', transparent = True)   # Histogram
    # figt4.savefig(f'{exportFoldernameHistogram}\\{UniqueFileName} histogram.png', bbox_inches='tight', pad_inches = 0, transparent = True)   # Histogram

def IndividualData(Studentstouse, Studentdatframe, ClassAverages, subjectstaught, numberofgraphs, exportfoldernameBar2, UniqueFileName):
    remeinder = Studentstouse % numberofgraphs
    studentset = round((Studentstouse - remeinder)/numberofgraphs)
    alignBars='center'
    if remeinder != 0 : numberofgraphs += 1
    lastLoopChecker = studentset
    barWidth = 0.9/studentset
    ind = np.arange(len(subjectstaught))
    spread = [(x*barWidth) for x in range(-round(studentset/2), round(studentset/2))]
    if lastLoopChecker % 2 == 0 : alignBars='edge'

    fig, axis = plt.subplots(numberofgraphs, figsize=(40,60), dpi=30)
    plt.rcParams['font.size'] = 80
    axis[0].set_title("All student data marked by color")
    plt.rcParams['font.size'] = 30
    for l in range(numberofgraphs):
        axis[l].grid(True, linestyle='solid', color = "lightgrey")
        axis[l].set_axisbelow(True)
        # spread = [0, -barWidth, barWidth, -2*barWidth, 2*barWidth, -3*barWidth, 3*barWidth, -4*barWidth, 4*barWidth, -5*barWidth, 5*barWidth]
        if l == numberofgraphs-1 and remeinder!= 0 : lastLoopChecker = remeinder
        for i in range(lastLoopChecker):
            studentlist = Studentdatframe.iloc[l*studentset + i,1:len(subjectstaught)+1]
            axis[l].bar(ind+spread[round(studentset/2)-round(lastLoopChecker/2) + i], studentlist, label=Studentdatframe.iloc[l*studentset + i,0], width=barWidth, align=alignBars)
        axis[l].scatter(subjectstaught, ClassAverages, s = 20000, c="red", lw=5, marker='_')
        axis[l].scatter(subjectstaught, ClassAverages, s = 1000, c="red", lw=5, marker='_', label="Class Average")
        axis[l].legend(bbox_to_anchor=(1, 1.05), loc='upper left')
    plt.tight_layout()
    fig.savefig(f'{exportfoldernameBar2}\\{UniqueFileName} classBarCombined.png')    # Save bar chart

def totalMarksTogether(Dataframe,maxvalue,exportfoldername,UniqueFileName):
    fig, ax = plt.subplots(figsize=(35,45),dpi=30)

    ax.grid(True, linestyle='solid', color = "white")     # Show grid behind bars
    ax.set_facecolor('#f2cadf')                           # Set plot backgroud
    ax.set_axisbelow(True)

    ticks = np.arange(0,maxvalue,10)                                     # Setting Y axis values to start from 0 increase in steps of five and end at 100
    people = np.arange(0, len(Dataframe.iloc[:,0]))

    ax.set_yticks(ticks)
    ax.set_xticks(people)
    ax.set_xticklabels(Dataframe.iloc[:,0], rotation=45, ha='right', fontsize=37)

    for x in range(100, round(maxvalue),100) :
        ax.plot(people,[x]*len(Dataframe.iloc[:,0]), c="red", lw=5, marker='_')
    ax.bar(Dataframe.iloc[:,0],Dataframe.iloc[:,-3], label='Student marks', color = "#936daf")    # Students Data
    plt.tight_layout()

    fig.savefig(f'{exportfoldername}\\{UniqueFileName} totalMarks.png')

def CreatePDF(studentDataToDisplay, fullNamesSeparated, FullNamesCombines, ImportDatalocation, charts, exportfolderPDF, Sessiondata, includePosition=True):  # Creating PDF's
    # Setup
    pdf = FPDF('P', 'mm', format='A4')                                      # Create PDF object, set orientation to potrait and units to milimeters
    pdf.add_page()                                                          # Add page to pdf
    pdf.set_margins(left=20, top=25, right=20)                              # Setup margins
    pdf.set_author('Mrs Theuri')
    pdf.set_producer('Ilmdl')
    pdf.set_title(f'{FullNamesCombines} Exam data analysis')
    pdf.set_lang('English')

    # Variables
    borderState = False                                                 # Controls text border visibility
    DisplayNames = f'{fullNamesSeparated[0]} {fullNamesSeparated[1]}'   # Student first and second
    rectangleHeight = 20

    # Functions
    def roundRec(Posx, Posy, height, width, roundness) :                    # Draw a rectangle rounded on one side
        pdf.set_fill_color(r=242, g=202, b=223)                                             # Setting the fill color for a shape about to be drawn
        pdf.rect(x=Posx, y=Posy, w=width-(roundness/2), h=height, style="F")                # 

        pdf.circle(x=Posx+width-roundness, y=Posy, r=roundness, style="F")
        pdf.circle(x=Posx+width-roundness, y=Posy+height-roundness, r=roundness, style="F")
        
        pdf.rect(x=Posx+width-(roundness/2), y=Posy+(roundness/2), w=roundness/2, h=height-roundness, style="F")

    # Header
    pdf.set_font('Helvetica', '', 24)                                                                                                     # Set font and size
    pdf.multi_cell(130, 10, f'{DisplayNames}', border=borderState, new_y='Top')                                                                     # Place names
    pdf.set_font('Helvetica', '', 15)                                                                                                                  # Set font and size
    pdf.multi_cell(40, 6, f'{Sessiondata[1]} {Sessiondata[2]}\n{Sessiondata[3]}' , border=borderState, new_y='Next', new_x='LMARGIN', align='R')    # Place exam session, term number and year
    
    # Draw line
    row1names = pdf.get_string_width(DisplayNames) + 30        # Get string width based on font
    pdf.set_draw_color(r=242, g=202, b=223)
    pdf.set_line_width(1)
    pdf.line(x1=20, y1=35, x2=20+row1names, y2=35)

    # Place data
    pdf.cell(45, 13, "", border=borderState, new_y='Next', new_x='LMARGIN')                                                             # Spacing
    if not(includePosition) : rectangleHeight *= 2/3
    rectangleX = pdf.get_x()
    rectangleY = pdf.get_y() - 2
    roundRec(rectangleX, rectangleY, rectangleHeight, 42, 10) # Make rectange
    pdf.set_font('helvetica', '', 12)                                                                                                   # Set font and size
    if not(includePosition) :
        pdf.cell(45, 5.5, f'Mean score: {studentDataToDisplay.iloc[-1]}', border=borderState, new_y='Next', new_x='LMARGIN')                     # Place mean
        pdf.cell(45, 5.5, f'Total score: {studentDataToDisplay.iloc[-2]}', border=borderState, new_y='Next', new_x='LMARGIN')                    # Place Total marks
    else:
        pdf.cell(45, 5.5, f'Mean score: {studentDataToDisplay.iloc[-2]}', border=borderState, new_y='Next', new_x='LMARGIN')                     # Place mean
        pdf.cell(45, 5.5, f'Total score: {studentDataToDisplay.iloc[-3]}', border=borderState, new_y='Next', new_x='LMARGIN')                    # Place Total marks
    if includePosition : pdf.cell(45, 5.5, f'Position: {studentDataToDisplay.iloc[-1]}', border=borderState, new_y='Next', new_x='LMARGIN')  # Place position
    
    # Insert charts
    for insertchart in charts :
        pdf.image(f'{ImportDatalocation}\\{FullNamesCombines} {insertchart}.png', w=pdf.epw) # Place pie chart
        pdf.cell(45, 3, "", border=borderState, new_y='Next', new_x='LMARGIN')                                                             # Spacing

    pdf.output(f'{exportfolderPDF}\\{FullNamesCombines}.pdf') # Output pdf

#----------------------------------------------------------------------Creating student PDF---------------------------------------------------------------------------------------
chartsToUse = ["pie", "bar", "scat", "histogram", "classBarCombined", "totalMarks"]
for i in range(1):
    fullNames = cleanData.iloc[i,0]                                                         # Student names
    SplitFormatedNames = [singleNames.capitalize() for singleNames in fullNames.split()]            # Student names as a list with only the firt letter capital
    SingleStudentData = studentData.iloc[i]
    stringNames = " ".join(SplitFormatedNames)
    insertCharts = [f'{Datadir}\\{stringNames} bar.png']

    CreateBarChart(subList[:-2], SingleStudentData[:-3], SubjectAverages[:-2], Datadir, stringNames, 2)
    CreatePieChartLMIT(SingleStudentData[:-3], refreceValues.iloc[-2], SetColors, Datadir, stringNames)
    CreatePDF(SingleStudentData, SplitFormatedNames, stringNames, Datadir, chartsToUse[:2], Pdfdir, examInformationseparated)
    plt.close()
    print(f"student {i} done")

# Teacher stuff
CreateBarChart(subList[:-2], SubjectAverages[:-2], SubjectAverages[:-2], Datadir, examInformation)
CreatePieChartLMIT(SubjectAverages[:-2], refreceValues.iloc[-2], SetColors, Datadir, examInformation)
CreateScatterplot(FrequencyTable, Datadir, examInformation)
CreateHistogram(FrequencyTable['Score'], Datadir, examInformation)
IndividualData(NumStudents, cleanData, SubjectAverages[:-2],subList[:-2], 4, Datadir, examInformation)
totalMarksTogether(cleanData,refreceValues.iloc[-2], Datadir, examInformation)
CreatePDF(SubjectAverages, examInformationseparated, examInformation, Datadir, chartsToUse, Pdfdir, examInformationseparated, False)
# plt.show()