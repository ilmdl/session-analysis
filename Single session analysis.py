from matplotlib import patheffects
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import (MultipleLocator)
from matplotlib import colors
from matplotlib.ticker import PercentFormatter
import numpy as np
from fpdf import FPDF
import os

SetColors = ['#d6afd2','#f2cadf', '#936daf', '#d6afd2','#f2cadf', '#936daf', '#f2cadf']
subList = ['ENGLISH', 'KISWA', 'MATHS', 'HYGIENE', 'ENVIRON', 'C.R.E', 'ACM', 'Total Marks', 'Average']
BaskFont = r'C:\Users\ILMBL\AppData\Local\Microsoft\Windows\Fonts\Baskerville-Regular.ttf'
CalFont = r'C:\Windows\Fonts\Calibri.ttf'
NumStudents = 28
positionVariance = ['PSTN','POSITION']
NameCol = 'NAME'                                                            # Variable to store the name column
file = ("1 RED OPENER TERM 1 2022")
studPosition = positionVariance[0]

# Getting data from filename to be used in pdf
grade, gcolor, session, term, termNum, year = file.split()
gcolor = gcolor.capitalize()
session = session.capitalize()
term = term.capitalize()
grade = grade + ' ' + gcolor
term = term + ' ' + termNum
file = f'C:/Users/ILMBL/Documents/Projects/Programing/Python/Source file/{file}.xlsx'
# print(grade, gcolor, session, term, termNum, year, file, sep="\n")

# Setting up the directories to store the final pdf's and the graphs used in them
Pdfdir = f'PDFs {grade} {session} {term} {year}/'
Datadir = f'Data {grade} {session} {term} {year}/'
if not os.path.exists(Pdfdir): os.makedirs(Pdfdir)       # Create a new directory because it does not exist 
if not os.path.exists(Datadir): os.makedirs(Datadir)     # Create a new directory because it does not exist 

# Convert excel file into a pandas database and clean it
data = pd.read_excel(file)                                                  # All data from the original excel spreadsheet
data.columns = data.columns.str.strip()                                     # Remove whitespace from column names
cleanData = data
cleanData = cleanData.drop(cleanData.index[-1])
cleanData = cleanData.drop(cleanData.index[-1])
if not(subList[6] in cleanData.columns ): cleanData.loc[:,subList[6]] = 0   # Adding column for calculated subject if it doesn't exist
def hex_to_rgb(value):                                                      # Function to convert a hex value to rgb
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))
Purple1 = hex_to_rgb('#f2cadf')

# Calculating student total, average and ACM
for i in range(NumStudents) :
    SplitFormatedNames = cleanData.iloc[i,0]                                            # Get student names on column 0 row i
    cleanData = cleanData.set_index(NameCol)                                   # Setting the name column as the index
    # Calculating ACM
    cleanData.loc[SplitFormatedNames, 'ACM'] = cleanData.loc[SplitFormatedNames,'ART/MUSIC'] + cleanData.loc[SplitFormatedNames,'PSYCHO']  # Combine art and psycho score and put it into the ACM column
    # Calculating total marks and average for a student
    cleanData.loc[SplitFormatedNames, 'Total Marks'] = cleanData[subList[:7]].iloc[i].sum()                              # Sum the valuesof specific subjects in a row to get the total marks and insert the result in the total marks column
    cleanData.loc[SplitFormatedNames, 'Average'] = round(cleanData[subList[:7]].iloc[i].mean())                          # same as above but for mean
    cleanData = cleanData.reset_index(NameCol)                                                                  # Re-setting index

refreceValues = cleanData[subList].iloc[0,]
cleanData = cleanData.drop(cleanData.index[0])                                        # Remove first "OUT OFF" row
cleanData = cleanData.sort_values(by="Total Marks", ascending=False)
cleanData['Position'] = list(range(1,NumStudents))
NumStudents -= 1
studentData = cleanData[subList[:7]]                                                                            # Dataframe containing only the marks for each student

# Finding the class average for each subject
SubjectAverages = []
for i in subList:
    SubjectAverages.append(round(cleanData[i].mean(),1))


# Output
Fig = list(range(NumStudents))
fig1 = list(range(NumStudents))

def CreatePairBarChart(XAxisLables, singleStudentDataXaxis, ClassDataXaxis):    # Create bar chart
    Figure, axis = plt.subplots(figsize=(35, 20), dpi=40)
    Figure.patch.set_alpha(0.5)

    #Bar chart style------
    axis.grid(True, linestyle='solid', color = "white")     # Show grid behind bars
    axis.set_facecolor('#f2cadf')                           # Set plot backgroud
    # axis.patch.set_alpha(0.6)                           

    # Bar chart customization
    ticks = np.arange(0,101,5)                                     # Setting Y axis values to start from 0 increase in steps of five and end at 100
    axis.set_ylim(0,100)                                                
    axis.set_yticks(ticks, lenght=30)
    axis.set_yticklabels(ticks, rotation=0, ha='right', fontsize=37)
    axis.set_xticks([0,1,2,3,4,5,6], lenght=30)
    axis.set_xticklabels(XAxisLables, fontsize=50)
    axis.xaxis.set_minor_locator(MultipleLocator(1))
    axis.tick_params('both', length=5, width=1, which='major')
    axis.set_axisbelow(True)

    # Plot Bar chart------------------------------------------------------------
    axis.bar(XAxisLables, singleStudentDataXaxis, width=-0.4, align='edge', label='Student marks', color = "#d6afd2")    # Students Data
    axis.bar(XAxisLables, ClassDataXaxis, width=0.4, align='edge', label='Class avearge', color = "#936daf")   # Class average
    axis.legend(bbox_to_anchor = (1.01,1.15), loc='upper right')                                                    # Legend

    Figure.savefig(f'Data {grade} {session} {term} {year}\{SplitFormatedNames} bar.png', facecolor=Figure.get_facecolor())    # Save bar chart

def CreatePieChartBSH(StudentDatatoplot, Colors):  # Create a pie chart with the best subject highlighted
    displayData = []
    colorsMost = Colors
    colorsBest = ['none'] * 7
    plt.rcParams['font.size'] = 24

    # Find and highlight best subject
    expd = [0] * len(StudentDatatoplot)
    BestSubjectPositon = np.array(StudentDatatoplot).argmax()
    expd[BestSubjectPositon] = -0.05
    colorsBest[BestSubjectPositon] = Colors[BestSubjectPositon]
    colorsMost[BestSubjectPositon] = 'none'
    
    # Create pie chart
    Figure1, ax =  plt.subplots(figsize=(20, 7), dpi = 50)
    wedges, text = ax.pie(StudentDatatoplot, wedgeprops=dict(width=0.5), startangle=90, colors=colorsMost, counterclock=False, shadow=False)
    wedges1 , text1 = ax.pie(StudentDatatoplot, wedgeprops=dict(width=0.6), startangle=90, colors=colorsBest, counterclock=False, radius=1.1, explode=expd)
    
    # Customize chart
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="<|-",facecolor='black', linewidth=2),
            bbox=bbox_props, zorder=0, va="center")

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
        displayData.insert(i, f'{StudentDatatoplot.index[i]}: {StudentDatatoplot[i]} marks\n{round(StudentDatatoplot[i]/StudentDatatoplot.sum()*100,1)}%')  # Box data
        ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y), horizontalalignment=horizontalalignment, **kw)              # Create annotaion box 

    # Save the charts
    Figure1.savefig(f'Data {grade} {session} {term} {year}\{SplitFormatedNames} pie.png', transparent = True)     # Pie chart

def CreatePDF (cleanData):
    # Setup
    pdf = FPDF('P', 'mm', format='A4')                                      # Create PDF object, set orientation to potrait and units to milimeters
    pdf.add_page()                                                          # Add page to pdf
    pdf.add_font('calibr', '', CalFont)                                     # Add font
    pdf.add_font('Baskerville-Regular', '', BaskFont)                       # Add font
    pdf.set_margins(left=20, top=25, right=20)                              # Setup margins
    pdf.set_author('Mrs Theuri')
    pdf.set_producer('Ilmdl')
    pdf.set_title(f'{" ".join(SplitFormatedNames)} {session} {term} {year} Exam data analysis')
    pdf.set_lang('English')

    # Variables
    borderState = False                                                  # Controls text border visibility
    DisplayNames = f'{SplitFormatedNames[0]} {SplitFormatedNames[1]}'   # Student first and second

    # Functions
    def roundRec(Posx, Posy, height, width, roundness) :                    # Draw a rectangle rounded on one side
        pdf.set_fill_color(r=242, g=202, b=223)                                             # Setting the fill color for a shape about to be drawn
        pdf.rect(x=Posx, y=Posy, w=width-(roundness/2), h=height, style="F")                # 

        # pdf.set_fill_color(r=0, g=202, b=223)
        pdf.circle(x=Posx+width-roundness, y=Posy, r=roundness, style="F")
        pdf.circle(x=Posx+width-roundness, y=Posy+height-roundness, r=roundness, style="F")
        
        # pdf.set_fill_color(r=0, g=100, b=0)
        pdf.rect(x=Posx+width-(roundness/2), y=Posy+(roundness/2), w=roundness/2, h=height-roundness, style="F")

    # Header
    pdf.set_font('Baskerville-Regular', '', 24)                                                         # Set font and size
    pdf.multi_cell(130, 10, f'{DisplayNames}', border=borderState, new_y='Top')                           # Place names
    pdf.set_font('calibr', '', 15)                                                                      # Set font and size
    row1names = pdf.get_string_width(DisplayNames) - 3                                                    # Get string width based on font
    pdf.multi_cell(40, 6, f'{session} {term}\n{year}' , border=borderState, new_y='Next', new_x='LMARGIN', align='R')   # Place exam session, term number and year
    
    # Draw line
    pdf.set_draw_color(r=242, g=202, b=223)
    pdf.set_line_width(1)
    pdf.line(x1=20, y1=35, x2=25+row1names, y2=35)

    # Place data
    roundRec(20,48,20,42,10) # Make rectange
    pdf.set_font('helvetica', '', 12)                                                                           # Set font and size
    pdf.cell(45, 13, "", border=borderState, new_y='Next', new_x='LMARGIN')                                     # Spacing
    pdf.cell(45, 5.5, f'Mean score: {cleanData[subList[8]][1]}', border=borderState, new_y='Next', new_x='LMARGIN') # Place mean
    pdf.cell(45, 5.5, f'Total score: {cleanData[subList[7]][1]}', border=borderState, new_y='Next', new_x='LMARGIN')  # Place Total marks
    pdf.cell(45, 5.5, f'Position: {cleanData["Position"][1]}', border=borderState, new_y='Next', new_x='LMARGIN')  # Place position
    
    # Insert charts
    pdf.image(f'Data {grade} {session} {term} {year}\{SplitFormatedNames} pie.png', 20, 65, w=pdf.epw) # Place pie chart
    pdf.image(f'Data {grade} {session} {term} {year}\{SplitFormatedNames} bar.png', 20, 160, w=pdf.epw) # Place bar chart

    pdf.output(f'PDFs {grade} {session} {term} {year}\{" ".join(SplitFormatedNames)} {session} {term} {year}.pdf') # Output pdf

#----------------------------------------------------------------------Creating student PDF---------------------------------------------------------------------------------------

# Student data to print on pdf and to use for loc--------------------------------------------------------------
fullNames = cleanData.iloc[1,0]                                                         # Student names
SplitFormatedNames = [singleNames.capitalize() for singleNames in fullNames.split()]            # Student names as a list with only the firt letter capital
SingleStudentData = studentData.iloc[1]

CreatePairBarChart([subject.capitalize() for subject in subList[:7]], SingleStudentData, SubjectAverages[:7])
CreatePieChartBSH(SingleStudentData, SetColors)
CreatePDF(cleanData)

plt.show()