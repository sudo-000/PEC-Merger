import json
import os
import time

from openpyxl.chart.axis import ChartLines

os.system("md \".\\Selected\" >nul")
os.system("start \"Selected\" \"Explorer.exe\" \".\\Selected\\\"")
print('\033[2J\033[0;0H', end='')
print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
print('\033[35mInstructions: ')
print('Charts stored in PEC format takes up less space than RPE format with the same amount of information. ')
print('You cannot edit PEC strait away in Ex:Phiedit, but you can convert them into PEC when you\'ve finished editing them. ')
print('To convert a RPE format chart to PEC, load it up in Ex:Phiedit and click "Export old format" in settings. ')
print('Put all the charts you want to merge into the "Selected" folder. (PEC format only!)')
print('You may rename the chart files to re-order them. ')
print('You can also duplicate files if you want to overlap it with its self. \n')
input('\033[32mPress the return key when done...\033[0m\033[1m')
#Process the user's selection
os.system("md \".\\Selected\\Result\" >nul & attrib \".\\Selected\\Result\" +s +h & dir \".\\Selected\\*.*\" /b /o:n > \".\\Selected\\Result\\.txt\"")
if os.path.getsize(".\\Selected\\Result\\.txt") <= 0: #File empty?
    os.system("del \".\\Selected\" /f /s /q & rmdir \".\\Selected\" /s /q")
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mERROR: No charts are being selected! \nExiting now...')
    time.sleep(5)
    quit()
else: #The list file is not empty, GOOD! Let's move on.
    result = open(".\\Selected\\Result\\.txt", mode='r', encoding='UTF-8')
    result = result.read().splitlines()
    #'result' now holds the list of files that exist in '.\Selected'
    opened_charts = []
    merge_bpm = False
    #Configure some options:
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mOptions: ')
    loop = True
    while loop:
        selection = input('\033[35mDo you want to merge BPM events? [Y/N]')  # Merge BPM?
        if selection.upper() == 'Y' or selection.upper() == 'N':
            loop = False
            if selection.upper() == 'Y': merge_bpm = True
        else:
            print('\033[31mError: Invalid input! Please try again.')
    del loop
    del selection
    #Time to open up the actual chart files!
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mPlease wait...')
    print('Reading charts...')
    for filename in result:
        opened_charts.append(open(".\\Selected\\" + filename, mode='r', encoding='UTF-8').read().splitlines())
        print('Loaded ' + filename)
    del filename
    print('Merging begins...')
    final_chart = opened_charts[0]
    EventCount = 0
    TotalLines = 0
    ChartLines = 0
    ChartCount = 0
    for event in final_chart:
        if event.startswith("cp") or event.startswith("cm") or event.startswith("cd") or event.startswith("cr") or event.startswith("ca") or event.startswith("cf") or event.startswith("cv") or event.startswith("n1") or event.startswith("n2") or event.startswith("n3") or event.startswith("n4"):
            print('Processing event ' + str(EventCount) + ' on line ' + event.split(' ')[1] + ' in chart ' + result[ChartCount])
            lst_event = event.split(' ')
            ChartLines = max(int(lst_event[1]), ChartLines)
        EventCount += 1
    TotalLines += ChartLines + 1
    ChartCount += 1
    opened_charts.pop(0)
    while len(opened_charts) > 0:
        EventCount = 0
        ChartLines = 0
        for event in opened_charts[0]:
            if event.startswith("bp") and merge_bpm:
                final_chart.append(event)
            elif event.startswith("cp") or event.startswith("cm") or event.startswith("cd") or event.startswith("cr") or event.startswith("ca") or event.startswith("cf") or event.startswith("cv") or event.startswith("n1") or event.startswith("n2") or event.startswith("n3") or event.startswith("n4"):
                print('Processing event ' + str(EventCount) + ' on line ' + event.split(' ')[1] + ' in chart ' + result[ChartCount])
                lst_event = event.split(' ')
                ChartLines = max(int(lst_event[1]), ChartLines)
                lst_event[1] = str(int(lst_event[1]) + TotalLines)
                str_event = ""
                for item in lst_event:
                    str_event += item + ' '
                final_chart.append(str_event)
            elif event.startswith('#') or event.startswith('&'):
                final_chart.append(event)
            EventCount += 1
        TotalLines += ChartLines + 1
        ChartCount += 1
        opened_charts.pop(0)
    print('Writing to ' + '.\\Resources\\Result.PEC')
    final_output = open('.\\Resources\\Result.PEC', mode='w', encoding='UTF-8')
    for event in final_chart:
        final_output.write(event + '\n')
    del final_output
    del final_chart
    del opened_charts
    del EventCount
    del ChartCount
    del TotalLines
    del ChartLines
    del event
    del lst_event
    del str_event
    del merge_bpm
    del result
    os.system("del \".\\Selected\" /f /s /q & rmdir \".\\Selected\" /s /q")
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mAll done! The merged chart has been exported to .\\Resources\\Result.PEC')
    print('Thank you for using this program. ')
    time.sleep(5)
    quit()