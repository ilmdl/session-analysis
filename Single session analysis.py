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
subList = ['ENGLISH', 'KISWA', 'MATHS', 'HYGIENE', 'ENVIRON', 'C.R.E', 'ACM']
BaskFont = r'C:\Users\ILMBL\AppData\Local\Microsoft\Windows\Fonts\Baskerville-Regular.ttf'
CalFont = r'C:\Windows\Fonts\Calibri.ttf'
NumStudents = 26
files = ['1 RED OPENER TERM 1 2022','1 RED MID TERM 1 2022'] 
positionVariance = ['PSTN','POSITION']
NameCol = 'NAME'                                                            # Variable to store the name column
file = files[1]
studPosition = positionVariance[1]

# Title data collection
print(len(file.split()))
grade, gcolor, session, term, termNum, year = file.split()
gcolor = gcolor.capitalize()
session = session.capitalize()
term = term.capitalize()
grade = grade + ' ' + gcolor
term = term + ' ' + termNum
file = f'C:/Users/ILMBL/Documents/Projects/Programing/Python/Source file\{file}.xlsx'
print(grade, gcolor, session, term, termNum, year, file, sep="\n")

# Directory setup
Pdfdir = f'PDFs {grade} {session} {term} {year}/'
Datadir = f'Data {grade} {session} {term} {year}/'
PdfdirExist = os.path.exists(Pdfdir)
DatadirExist = os.path.exists(Datadir)
if not PdfdirExist:
    # Create a new directory because it does not exist 
    os.makedirs(Pdfdir)

if not DatadirExist:
    # Create a new directory because it does not exist 
    os.makedirs(Datadir)

# Import data using pandas
data = pd.read_excel(file)                                                  # All data from the original excel spreadsheet
data.columns = data.columns.str.strip()                                     # Remove whitespace from column names
cleanData = data.drop(data.index[0])

# cleanData = cleanData.head(NumStudents)                                          # Dataframe containing only the data for each student data
if not(subList[6] in cleanData.columns ):
    cleanData.loc[:,subList[6]] = 0
studentData = cleanData[subList]                                            # Dataframe containing only the marks for each student
print(cleanData)
# studentProfile = cleanData[cleanData[NameCol] == 'ELLA BAYARDE ESSOME']   # Table only containind 1 row of data based on the name provided
students = cleanData.NAME                                                   # list of items on the column Name (As a series)
headval = cleanData[NameCol] == "TOTAL"
headval = cleanData.loc[cleanData[NameCol] == "TOTAL", :].index
print('***', "this is head val ", headval, '***')

def hex_to_rgb(value):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))
Purple1 = hex_to_rgb('#f2cadf')
print(Purple1, "\n\n\n")
print(cleanData)

# Setting the total marks and average for each student
for i in range(NumStudents) :
    fullNames = cleanData.iloc[i,0]                                            # Name for specific student
    cleanData = cleanData.set_index(NameCol)                                   # Setting index

    # Adding values ot ACM
    art = cleanData.loc[fullNames,'ART/MUSIC'] # Get the art score
    psycho = cleanData.loc[fullNames,'PSYCHO'] # Get the Pscho score
    acm = art + psycho
    cleanData.loc[fullNames, 'ACM'] = acm
    studentData = cleanData[subList]

    totalMarks = studentData.iloc[i].sum()                                     # Total marks for the list of subects provide
    averageMarks = round(studentData.iloc[i].mean())                                  # Average marks using the mean function
    cleanData.loc[fullNames, 'Total Marks'] = totalMarks
    cleanData.loc[fullNames, 'Average'] = averageMarks
    cleanData = cleanData.reset_index(NameCol)                   # Re-setting index

# Finding the class average for each subject
classEngAverage = round(cleanData[subList[0]].mean())
classKisAverage = round(cleanData[subList[1]].mean())
classMatAverage = round(cleanData[subList[2]].mean())
classHygAverage = round(cleanData[subList[3]].mean())
classEnvAverage = round(cleanData[subList[4]].mean())
classCREAverage = round(cleanData[subList[5]].mean())
classAMCAverage = round(cleanData[subList[6]].mean())
classAverage = [classEngAverage, classKisAverage, classMatAverage, classHygAverage, classEnvAverage, classCREAverage, classAMCAverage]
classTotal = cleanData.loc[:, 'Total Marks'].sum()
classTotalAverage = cleanData.loc[:, 'Total Marks'].mean()
classSubAverage = cleanData.loc[:, 'Average'].mean()

# Output
fig = list(range(NumStudents))
fig1 = list(range(NumStudents))
displayData = []




# ----------------------------------------------------------------------Creating student PDF---------------------------------------------------------------------------------------
for j in range(2) :
    plt.rcParams['font.size'] = 24
    # Student data to use for each PDF--------------------------------------------------------------
    print('student number: {}'.format(j))
    fullNames = cleanData.iloc[j,0]                                # Student name
    numNames = len(fullNames.split())
    if numNames > 2 :
        first, middle, last = fullNames.split()
        first = first.capitalize()
        middle = middle.capitalize()
        last = last.capitalize()
    else :
        first, middle = fullNames.split()
        first = first.capitalize()
        middle = middle.capitalize()
        last = " "
    subjects = studentData.iloc[j]
    
    plt.rcParams['font.size'] = 40
    # Bar chart data------------------------------------------------------------
    fig1[j], ax2 =  plt.subplots(figsize=(30, 20), dpi=100)
    fig1[j].patch.set_alpha(0.5)
    # ax2.invert_yaxis()

    #Bar chart style------
    ax2.grid(True, linestyle='solid', color = "white")
    ax2.set_facecolor('#f2cadf')
    ax2.patch.set_alpha(0.6)

    # Bar chart customization
    ticcks = np.arange(0,101,5)
    ax2.set_ylim(0,100)
    ax2.set_yticks(ticcks, lenght=30)
    ax2.set_yticklabels(ticcks, rotation=0, ha='right', fontsize=37)
    ax2.set_xticks([0,1,2,3,4,5,6], lenght=30)
    ax2.set_xticklabels(subList, fontsize=50)
    ax2.xaxis.set_minor_locator(MultipleLocator(1))
    ax2.tick_params('both', length=5, width=1, which='major')
    ax2.set_axisbelow(True)

    # Plot Bar chart------------------------------------------------------------
    ax2.bar(subList, subjects, width=-0.4, align='edge', label='Student marks', color = "#d6afd2")
    ax2.bar(subList, classAverage, width=0.4, align='edge', label='Class avearge', color = "#936daf")
    ax2.legend(bbox_to_anchor = (1.01,1.15), loc='upper right')


    # -------------------------------------------------------------------------Creating pie chart------------------------------------------------------------------------------
    # Pie chart data------------------------------------------------------------
    colorsMost = ['none', 'none', 'none', 'none', 'none', 'none', 'none']
    colorsBest = ['none', 'none', 'none', 'none', 'none', 'none', 'none']
    plt.rcParams['font.size'] = 24
    plt.tight_layout()
    # Annotation box data
    for i in range(7) :
        displayData.insert(i, '{}: {} marks\n{}%'.format(subList[i],subjects[i],round(subjects[i]/subjects.sum()*100,1)))
    pAngle = 0
    fig[j], ax =  plt.subplots(figsize=(16, 8), subplot_kw=dict(aspect="equal"), dpi = 120)
    
    # Best subject calculations
    maxVal = subjects.max(0)
    expd = [0,0,0,0,0,0,0]
    for k in range(7) :
        print(k)
        CurrColor = SetColors[k]
        colorsMost[k] = CurrColor
        if subjects[k] == maxVal :
            expd[k] = -0.05
            colorsBest[k] = CurrColor
            print(" ",k)
            colorsMost[k] = 'none'
    wedges , text = ax.pie(subjects, wedgeprops=dict(width=0.5), startangle=90, colors=colorsMost, counterclock=False, shadow=False)
    wedges1 , text1 = ax.pie(subjects, wedgeprops=dict(width=0.6), startangle=90, colors=colorsBest, counterclock=False, radius=1.1, explode=expd)
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    kw = dict(arrowprops=dict(arrowstyle="<|-",facecolor='black', linewidth=2),
            bbox=bbox_props, zorder=0, va="center")

    for patch in wedges:
        patch.set_path_effects([patheffects.SimpleLineShadow(),
        patheffects.Normal()])

    # Annotating wedges--------
    print(f'Annotation for student {j}')
    for i, p in enumerate(wedges):
        print('wedge:', i)
        # Annotating wedges that are not the best subject
        if colorsBest[i] == 'none' :
            print("Not best subject")
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            # print("Angle:", ang)
            # print("Angle1:", p.theta1)
            # print("Angle2:", p.theta2)
            y = np.sin(np.deg2rad(ang))
            x = np.cos(np.deg2rad(ang))
            # print('x:', x)
            # print('y:', y)
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = "angle,angleA=0,angleB={}".format(ang)
            kw["arrowprops"].update({"connectionstyle": connectionstyle})

            print(displayData[i])
            if ((abs(ang - pAngle) < 15 and not(i == 0)) ) :
                bbox_props["lw"] = 2
                ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.5*y),
                        horizontalalignment=horizontalalignment, **kw)
                # print("gap is small","\n")
            else :
                bbox_props["lw"] = 0.75
                ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y),
                            horizontalalignment=horizontalalignment, **kw)
    # Anotating for best subject
    for i, p in enumerate(wedges1):
        print('wedge:', i)
        # Anotating for best subject
        if not(colorsBest[i] == 'none') :
            print("Best subject")
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            # print("Angle:", ang)
            # print("Angle1:", p.theta1)
            # print("Angle2:", p.theta2)
            y = np.sin(np.deg2rad(ang))* 1.05
            x = np.cos(np.deg2rad(ang))* 1.05
            # print('x:', x)
            # print('y:', y)
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = "angle,angleA=0,angleB={}".format(ang)
            kw["arrowprops"].update({"connectionstyle": connectionstyle})

            print(displayData[i])
            # print("previouse Angle:", pAngle)
            # print("Math:", ang-pAngle)
            # print(abs(ang - pAngle) < 15)
            if ((abs(ang - pAngle) < 15 and not(i == 0)) ) :
                bbox_props["lw"] = 2
                ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.5*y),
                        horizontalalignment=horizontalalignment, **kw)
            else :
                bbox_props["lw"] = 0.75
                ax.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y),
                            horizontalalignment=horizontalalignment, **kw)
            pAngle = ang
            print('Done', '\n')
    # Save the charts
    fig[j].savefig(f'Data {grade} {session} {term} {year}\{fullNames} pie.png', transparent = True)     # Pie chart
    fig1[j].savefig(f'Data {grade} {session} {term} {year}\{fullNames} bar.png', facecolor=fig1[j].get_facecolor())    # Bar chart

    # ----------------------------------------------------------------------Creating PDF-----------------------------------------------------------------------------
    # PDF data--------------------------------
    # Setup PDF
    pdf = FPDF('P', 'mm')                                           # Create PDF object, set orientation to potrait and units to milimeters
    pdf.add_page()                                                  # Add page to pdf
    pdf.add_font('calibr', '', CalFont)
    pdf.add_font('Baskerville-Regular', '', BaskFont)
    pdf.set_margins(left=20, top=25, right=20)                      # Setup margins
    width = 210-40                                                  # Width of document
    cleanData = cleanData.set_index(NameCol)                        # Setting the name colomn as the index used fort locating the average, total and position data
    printNames = f'{first} {middle}'                                # Student first and second
    printNames2 = f'{last}'                                         # Students last name
    printAverage = cleanData.loc[fullNames, 'Average']              # Student average score to print
    printTotal = cleanData.loc[fullNames, 'Total Marks']            # Student total to print
    printPosition = int(cleanData.loc[fullNames, studPosition])     # Student position to print
    printSession = f'{session} {term}\n{year}'                      # Session to print
    pieChart = f'Data {grade} {session} {term} {year}\{fullNames} pie.png'   # Pie chart to insert
    barChart = f'Data {grade} {session} {term} {year}\{fullNames} bar.png'   # Bar chart to insert
    cleanData = cleanData.reset_index(NameCol)                      # Reseting the index
    borderState = False                                             # Variable to allow for quich control of border visibility
    # with pdf.local_context(fill_opacity=0, stroke_opacity=0.5) :    # Draw a rectangel at the margin and make it tranparent
    # pdf.set_draw_color(1)
    # pdf.rect(x=20, y=25, w=210-40, h=297-50, style="D")
    
    # Put first and second name and exam session
    pdf.set_font('Baskerville-Regular', '', 24)                               # Set font for PDF
    pdf.multi_cell(130, 10, f'{printNames} \n{printNames2}', border=borderState, new_y='Top')
    row1names = pdf.get_string_width(printNames) - 3
    row2names = pdf.get_string_width(printNames2) - 3
    print("Row 1", row1names)
    print("Row 2", row2names)
    pdf.set_font('calibr', '', 15)                               # Set font for PDF
    pdf.multi_cell(40, 6, printSession, border=borderState, new_y='Next', new_x='LMARGIN', align='R')
    if numNames > 2 :
        pdf.set_draw_color(r=242, g=202, b=223)
        pdf.set_line_width(1)
        pdf.line(x1=20, y1=(25+45)/2, x2=25+row1names, y2=(25+45)/2)
        pdf.line(x1=20, y1=45, x2=25+row2names, y2=45)
    else :
        pdf.set_draw_color(r=242, g=202, b=223)
        pdf.set_line_width(1)
        pdf.line(x1=20, y1=(25+45)/2, x2=70, y2=(25+45)/2)
    
    # Scores and positions
    # Rounded rectangele
    def roundRec(Posx, Posy, height, width, roundness) :
        pdf.set_fill_color(r=242, g=202, b=223)                                           # Setting the fill color for a shape about to be drawn
        pdf.rect(x=Posx, y=Posy, w=width-(roundness/2), h=height, style="F")

        # pdf.set_fill_color(r=0, g=202, b=223)
        pdf.circle(x=Posx+width-roundness, y=Posy, r=roundness, style="F")
        pdf.circle(x=Posx+width-roundness, y=Posy+height-roundness, r=roundness, style="F")
        
        # pdf.set_fill_color(r=0, g=100, b=0)
        pdf.rect(x=Posx, y=Posy+(roundness/2), w=width, h=height-roundness, style="F")
    roundRec(20,48,20,42,10)

    # Put Mean
    pdf.set_font('helvetica', '', 12)
    pdf.cell(45, 13, "", border=borderState, new_y='Next', new_x='LMARGIN') # Spacing
    pdf.cell(45, 5.5, 'Mean score: {}'.format(printAverage), border=borderState, new_y='Next', new_x='LMARGIN')
    
    # Put Total marks
    pdf.cell(45, 5.5, 'Total score: {}'.format(printTotal), border=borderState, new_y='Next', new_x='LMARGIN')
    
    # Put Position
    pdf.cell(45, 5.5, 'Position: {}'.format(printPosition), border=borderState, new_y='Next', new_x='LMARGIN')
    
    # Insert charts
    # Put in pie chart
    pdf.image(pieChart, 20, 65, w=width)
    # Put in bar chart
    pdf.image(barChart, 20, 160, w=width)

    pdf.set_author('Mrs Theuri')
    pdf.set_producer('Brighton Academy')
    pdf.set_title(f'{fullNames} {session} {term} {year} Exam data analysis')
    pdf.set_lang('English')
    # Output pdf
    pdf.output(f'PDFs {grade} {session} {term} {year}\{fullNames} {session} {term} {year}.pdf')
    
# ----------------------------------------------------------------------End of PDF creation--------------------------------------------------------------------------------

pdft = FPDF
pdft = FPDF('P', 'mm')
pdft.add_page()
pdft.add_font('calibr', '', CalFont)
pdft.add_font('Baskerville-Regular', '', BaskFont)

# Pie chart data-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
colorsMost = ['none', 'none', 'none', 'none', 'none', 'none', 'none']
colorsBest = ['none', 'none', 'none', 'none', 'none', 'none', 'none']
plt.rcParams['font.size'] = 24

# Annotation box data-------------
for i in range(7) :
    displayData.insert(i, f'{subList[i]}: {classAverage[i]} marks\n{round(classAverage[i]/sum(list(classAverage))*100,1)}%')
pAngle = 0
figt, axt = plt.subplots(figsize=(16, 8), subplot_kw=dict(aspect="equal"), dpi = 120)
plt.tight_layout()

maxVal = max(list(classAverage))
expd = [0,0,0,0,0,0,0]
for k in range(7) :
    print(k)
    CurrColor = SetColors[k]
    colorsMost[k] = CurrColor
    if classAverage[k] == maxVal :
        expd[k] = -0.05
        colorsBest[k] = CurrColor
        print(" ",k)
        colorsMost[k] = 'none'
wedges , text = axt.pie(classAverage, wedgeprops=dict(width=0.5), 
                        startangle=90, colors=colorsMost, counterclock=False, 
                        shadow=False)
wedges1 , text1 = axt.pie(classAverage, wedgeprops=dict(width=0.6), 
                        startangle=90, colors=colorsBest, counterclock=False, 
                        radius=1.1, explode=expd)
bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
kw = dict(arrowprops=dict(arrowstyle="<|-",facecolor='black', linewidth=2),
        bbox=bbox_props, zorder=0, va="center")


# Annotating wedges----------------
for i, p in enumerate(wedges):
    print('wedge:', i)
    # Annotating wedges that are not the best subject
    if colorsBest[i] == 'none' :
        print("Not best subject")
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        # print("Angle:", ang)
        # print("Angle1:", p.theta1)
        # print("Angle2:", p.theta2)
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        # print('x:', x)
        # print('y:', y)
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        connectionstyle = "angle,angleA=0,angleB={}".format(ang)
        kw["arrowprops"].update({"connectionstyle": connectionstyle})

        print(displayData[i])
        if ((abs(ang - pAngle) < 15 and not(i == 0)) ) :
            bbox_props["lw"] = 2
            axt.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.5*y),
                    horizontalalignment=horizontalalignment, **kw)
            # print("gap is small","\n")
        else :
            bbox_props["lw"] = 0.75
            axt.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y),
                        horizontalalignment=horizontalalignment, **kw)
# Anotating best subject-----------
for i, p in enumerate(wedges1):
    print('wedge:', i)
    # Anotating for best subject
    if not(colorsBest[i] == 'none') :
        print("Best subject")
        ang = (p.theta2 - p.theta1)/2. + p.theta1
        # print("Angle:", ang)
        # print("Angle1:", p.theta1)
        # print("Angle2:", p.theta2)
        y = np.sin(np.deg2rad(ang))* 1.05
        x = np.cos(np.deg2rad(ang))* 1.05
        # print('x:', x)
        # print('y:', y)
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
        connectionstyle = "angle,angleA=0,angleB={}".format(ang)
        kw["arrowprops"].update({"connectionstyle": connectionstyle})

        print(displayData[i])
        # print("previouse Angle:", pAngle)
        # print("Math:", ang-pAngle)
        # print(abs(ang - pAngle) < 15)
        if ((abs(ang - pAngle) < 15 and not(i == 0)) ) :
            bbox_props["lw"] = 2
            axt.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.5*y),
                    horizontalalignment=horizontalalignment, **kw)
        else :
            bbox_props["lw"] = 0.75
            axt.annotate(displayData[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.3*y),
                        horizontalalignment=horizontalalignment, **kw)
        pAngle = ang
        print('Done', '\n')

# Save figure
figt.savefig(f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} pie.png', transparent = True)   # Pie chart


# Scatter plot---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
figt2, axt2 =  plt.subplots(figsize=(20, 20), dpi=120)
yaxisData = np.arange(0,101,1)
scatterTable = pd.DataFrame()


# Turning the student data int 3 columns subject, score and frequency the purpose is to make a dataframe for the scatter plot
for k in range(7) :
    scattersorted = sorted(studentData[subList[k]])
    freak = studentData[subList[k]].value_counts()
    for l in range(NumStudents) :
        val = 26 * k
        scatterTable.loc[val+l,'Subject'] = subList[k]
        scatterTable.loc[val+l,'Score'] = studentData.iloc[l,k]
        scatterTable.loc[val+l,'freak'] = freak[studentData.iloc[l,k]]

# Plot the scatter plot
axt2.grid(True, linestyle='solid', color = "black")
axt2.set_axisbelow(True)
axt2.yaxis.set_minor_locator(MultipleLocator(1))
axt2.set_title("Spread of marks per subject size indicates frequency")
axt2.set_xlabel("Subjects",fontsize=30)
axt2.set_ylabel("Marks",fontsize=30)
axt2.scatter(scatterTable['Subject'],scatterTable['Score'], 
            s=((scatterTable['freak']+1)*8)**2, 
            c=-(scatterTable['freak']-18)**6)

# Save figure
figt2.savefig(f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} scat.png', transparent = True)   # Scatter plot

# Bar graph-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
plt.rcParams['font.size'] = 15
# Bar chart data---------------------------
figt3, axt3 =  plt.subplots(figsize=(15, 10), dpi=100)
figt3.patch.set_alpha(0.5)
# ax2.invert_yaxis()

#Bar chart style------
axt3.grid(True, linestyle='solid', color = "white")
axt3.set_axisbelow(True)
axt3.set_facecolor('#f2cadf')
axt3.patch.set_alpha(0.6)

# Bar chart customization
ticcks = np.arange(0,101,5)
axt3.set_ylim(0,100)
axt3.set_yticks(ticcks, lenght=300)
axt3.set_yticklabels(ticcks, rotation=0, ha='right', fontsize=20)
axt3.set_xticks([0,1,2,3,4,5,6], lenght=30)
axt3.set_xticklabels(subList, fontsize=20)
axt3.xaxis.set_minor_locator(MultipleLocator(1))
axt3.tick_params('both', length=5, width=1, which='major')
axt3.set_facecolor('#f2cadf')
axt3.patch.set_alpha(0.6)

# Plot Bar chart------------------------------------------------------------
# axt3.bar(subList, subjects, width=-0.4, align='edge', label='Student marks', color = "#d6afd2")
axt3.bar(subList, classAverage, width=0.8, align='center', label='Class avearge', color = "#936daf")
axt3.legend(bbox_to_anchor = (1.01,1.1), loc='upper right')

# Save figure
figt3.savefig(f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} bar.png', facecolor=figt3.get_facecolor())   # bar chart

# Histogram chart------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
figt4,axt4 = plt.subplots(figsize= (20,20), dpi=100)
HisHeight = round(scatterTable['Score'].value_counts().max() + 1)
n_bins = len(scatterTable['Score'].value_counts())
axt4.grid(True, linestyle='solid', color = "lightgrey")
axt4.set_axisbelow(True)
xticcks = np.arange(0,101,5)
yticcks = np.arange(0,HisHeight,1)
axt4.set_ylim(0,HisHeight)
axt4.set_yticks(yticcks, lenght=1001)
axt4.set_yticklabels(yticcks, rotation=0, ha='right', fontsize=20)
axt4.set_xlim(0,100)
axt4.yaxis.set_minor_locator(MultipleLocator(1))
axt4.set_xticks(xticcks, lenght=50)
axt4.set_xticklabels(xticcks, rotation=0, ha='center', fontsize=20)
axt4.hist(scatterTable['Score'],bins = n_bins)
axt4.set_title("frequency of marks across all subjects")
axt4.set_xlabel("Marks")
axt4.set_ylabel("Frequency")
N, bins, patches = axt4.hist(scatterTable['Score'],bins = n_bins)
 
# Setting color
fracs = ((N**(1 / 2)) / N.max())
norm = colors.Normalize(fracs.min(), fracs.max())

for thisfrac, thispatch in zip(fracs, patches):
    color = plt.cm.viridis(norm(thisfrac))
    thispatch.set_facecolor(color)

# Save figure
figt4.savefig(f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} hist.png', transparent = True)   # Histogram

# Stuff to put on the pdf---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
pdft.set_font('Baskerville-Regular', '', 24)
pdft.set_font('calibr', '', 15) 
pdft.set_margins(left=20, top=25, right=20)
Classpie = f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} pie.png'
Classbar = f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} bar.png'
Classscat = f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} scat.png'
Classhist = f'Data {grade} {session} {term} {year}\{grade} {session} {term} {year} hist.png'
pdft.set_font('Baskerville-Regular', '', 24)                               # Set font for PDF
printTitle = f'{grade} {session}'
printTitle2 = f'{term} {year}'
row1names = pdft.get_string_width(printTitle)+1
row2names = pdft.get_string_width(printTitle2)+1
width = 210-40
message = 'This is the first message'
borderState = False

# pdft.set_draw_color(1)
# pdft.rect(x=20, y=25, w=210-40, h=297-50, style="D")
    
# Put first and second name and exam session
pdft.multi_cell(130, 10, f'{printTitle} \n{printTitle2}', border=borderState, new_y='Top')
pdft.set_font('calibr', '', 15)                               # Set font for PDF
pdft.multi_cell(40, 6, (f'{session} {term} \n{year}'), border=borderState, new_y='Next', new_x='LMARGIN', align='R')

pdft.set_draw_color(r=242, g=202, b=223)
pdft.set_line_width(1)
pdft.line(x1=20, y1=(25+45)/2, x2=20+row1names, y2=(25+45)/2)
pdft.line(x1=20, y1=45, x2=20+row2names, y2=45)

# Scores and positions
# Rounded rectangele
def roundRec(Posx, Posy, height, width, roundness) :
    pdft.set_fill_color(r=242, g=202, b=223)                                           # Setting the fill color for a shape about to be drawn
    pdft.rect(x=Posx, y=Posy, w=width-(roundness/2), h=height, style="F")

    # pdf.set_fill_color(r=0, g=202, b=223)
    pdft.circle(x=Posx+width-roundness, y=Posy, r=roundness, style="F")
    pdft.circle(x=Posx+width-roundness, y=Posy+height-roundness, r=roundness, style="F")
    
    # pdf.set_fill_color(r=0, g=100, b=0)
    pdft.rect(x=Posx, y=Posy+(roundness/2), w=width, h=height-roundness, style="F")
roundRec(20,48,20,42,10)

# Put Mean
pdft.set_draw_color(0)
pdft.set_line_width(0.1)
pdft.set_font('helvetica', '', 12)
pdft.cell(45, 15.4, "", border=borderState, new_y='Next', new_x='LMARGIN') # Spacing
pdft.cell(45, 5.5, f'Mean score: {round(classSubAverage,2)}', border=borderState, new_y='Next', new_x='LMARGIN')

# Put Total marks
pdft.cell(45, 5.5, f'Total score: {round(classTotalAverage,2)}', border=borderState, new_y='Next', new_x='LMARGIN')

# Put Position
# pdft.cell(45, 5.5, 'Position: {}'.format(printPosition), border=borderState, new_y='Next', new_x='LMARGIN')

# Insert charts
# Put in pie chart
pdft.image(Classpie, 20, 68, w=width)
# Put in bar chart
pdft.image(Classbar, 20, 160, w=width)

pdft.set_author('Mrs Theuri')
pdft.set_producer('Brighton Academy')
pdft.set_title(f'{grade} {session} {term} {year} Exam data analysis')
pdft.set_lang('English')

pdft.add_page()
# pdft.set_draw_color(1)
# pdft.rect(x=20, y=25, w=210-40, h=297-50, style="D")
# Insert charts
# Put in scatter plot
pdft.image(Classscat, 20, 10, w=width)

pdft.add_page()
pdft.set_draw_color(1)
pdft.rect(x=20, y=25, w=210-40, h=297-50, style="D")
# Put in histogram
pdft.image(Classhist, 20, 10, w=width)


pdft.output(f'PDFs {grade} {session} {term} {year}\{grade} {session} {term} {year} analysis.pdf')
# Display data
print(cleanData)

# plt.show()