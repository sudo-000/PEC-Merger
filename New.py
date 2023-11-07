import json
import os
import time
print('\033[2J\033[0;0H', end='')
print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
print('\033[35mReminder: ')
print('Charts stored in the old PEC format takes up much less space than the new RPE format with the same amount of information. ')
print('Sadly you cannot edit PEC strait away in Ex:Phiedit, but you can convert your charts into PEC when you\'ve finished editing them. ')
print('To convert a RPE format chart to PEC, load it up in Ex:Phiedit and click "Export old format" in settings, then a zip file will be exported to .Resources\\\n')
print('Note that doing so will cause your chart to loose all the formality that the old format don\'t support! ')
print('You must keep XY move events binded and you can only use only one event layer. ')
print('Paint events, story board, parent line, link UI, basier curve, note alpha, visible time, and a lot more will be fully ignored when converting. \n')
input('\033[32mPress the return key when ready...\033[0m\033[1m')
print('\033[2J\033[0;0H', end='')
print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
loop=True
while loop:
    print('\033[35mSelect an operation to execute:')
    print('\033[34m  1. Merge all charts')
    print('  2. Move selected charts by a number of bars')
    choice = input('\033[35mPlease enter your choice: \033[0m\033[1m')
    if choice == '1' or choice == '2':
        loop=False
    else:
        print('\033[31m\nError: Invalid input! Please try again.')
del loop
os.system("md \".\\Selected\" >nul & md \".\\All Charts\" >nul & md \".\\PEC Unzip\" >nul")
#Reading Resources\*.zip
os.system("md \".\\All Charts\\List\" >nul & attrib \".\\All Charts\\List\" +s +h & dir \".\\Resources\\*.zip\" /b /o:n > \".\\All Charts\\List\\.txt\"")
if os.path.getsize(".\\All Charts\\List\\.txt") <= 0: #File empty?
    os.system("del \".\\Selected\" /f /s /q & rmdir \".\\Selected\" /s /q")
    os.system("del \".\\All Charts\" /f /s /q & rmdir \".\\All Charts\" /s /q")
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mERROR: No PEC charts found! \nExiting now...')
    time.sleep(5)
    quit()
else: #The list file is not empty, GOOD! Let's move on.
    #Create temp files for selection
    Chartlist = open(".\\All Charts\\List\\.txt", mode='r', encoding='UTF-8')
    Chartlist = Chartlist.read().splitlines()
    for Charts in Chartlist:
        os.system("Powershell /command \"& {Expand-Archive -Path \\\".\\Resources\\" + Charts + "\\\" -DestinationPath \\\".\\PEC Unzip\\" + Charts + "\\\" -Force}")
        info = open(".\\PEC Unzip" + Charts, mode='r', encoding='UTF-8')
        info = info.read().splitlines()
        File = open(".\\All Charts\\"+info[1].lstrip("Path: ")+".txt", mode="w", encoding='UTF-8')
        File.write("# Song details: \n\n")
        for lines in info:
            File.write(lines + '\n')
        File.close()
    os.system("del \".\\All Charts\\List\" /f /s /q & rmdir \".\\All Charts\\List\" /s /q")
    del File
    del info
    del Charts
    del ChartlistLST
    #Call explorer.exe
    os.system("start \"Selected\" \"Explorer.exe\" \".\\Selected\\\" & start \"All Charts\" \"Explorer.exe\" \".\\All Charts\\\"")
    print('\033[2J\033[0;0H', end='')
    print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
    print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
    print('\033[35mInstructions: ')
    print('All your charts are being listed in the "All Charts" folder. ')
    print('Double click on the text file to view the song details. ')
    print('Drag all the charts you want to add into the "Selected" folder. ')
    print('Charts will be ordered with its file name in increasing alphabetical order. ')
    print('You may rename these text files to re-order them but you should never edit anything inside the file. ')
    print('You can also duplicate files if you want to overlap it with its self. \n')
    input('\033[32mPress the return key when done...\033[0m\033[1m')
    #Process the user's selection
    os.system("md \".\\Selected\\Result\" >nul & attrib \".\\Selected\\Result\" +s +h & dir \".\\Selected\\*.txt\" /b /o:n > \".\\Selected\\Result\\.txt\"")
    os.system("del \".\\All Charts\" /f /s /q & rmdir \".\\All Charts\" /s /q")
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
        merge_chart = False
        obfuscate = False
        super_obf = False
        move_amount = 0.0
        offset_per_file = 0.0
        #Configure some options:
        print('\033[2J\033[0;0H', end='')
        print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
        print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
        print('\033[35mOptions: ')
        loop = True
        while loop:
            selection = input('\033[35mDo you want to obfuscate this chart? [Y/N]') #Making XPE unable to read this chart, but various chart players should work fine.
            if selection.upper() == 'Y' or selection.upper() == 'N':
                loop = False
                if selection.upper() == 'Y': obfuscate = True
            else:
                print('\033[31mError: Invalid input! Please try again.')
        if obfuscate:
            loop = True
            while loop:
                selection = input('\033[35mEnable extreme obfuscate? (Warning: May slow down event process!) [Y/N]') #Modify more items, making the chart even harder to recover.
                if selection.upper() == 'Y' or selection.upper() == 'N':
                    loop = False
                    if selection.upper() == 'Y': super_obf = True
                else:
                    print('\033[31mError: Invalid input! Please try again.')
        if choice == '2':
            from fractions import Fraction
            loop = True
            while loop:
                selection = input('\033[35mAmount of bars to move: ')
                try:
                    move_amount = Fraction(selection)
                    loop = False
                except:
                    print('\033[31mError: Invalid input! Please try again.')
            loop = True
            while loop:
                selection = input('\033[35mOffset per file: ')
                try:
                    offset_per_file = Fraction(selection)
                    loop = False
                except:
                    print('\033[31mError: Invalid input! Please try again.')
            loop = True
            while loop:
                selection = input('\033[35mDo you want to merge these charts? [Y/N]') #Ask if the user wants to merge all the moved charts into one
                if selection.upper() == 'Y' or selection.upper() == 'N':
                    loop = False
                    if selection.upper() == 'Y': merge_chart = True
                else:
                    print('\033[31mError: Invalid input! Please try again.')
        if choice == '1' or merge_chart:
            loop = True
            while loop:
                selection = input('\033[35mDo you want to merge BPM events? [Y/N]') #Merge BPM?
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
        opened_chart_list = []
        for filename in result:
            data = open(".\\Selected\\" + filename, mode='r', encoding='UTF-8').read().splitlines()
            opened_chart_list.append(data)
            opened_charts.append(json.load(open('.\\PEC Unzip\\' + data[3].lstrip("Path: ") + '\\' + data[6].lstrip("Chart: "), mode='r', encoding='UTF-8')))
            print('Loaded ' + filename)
        del filename
        del data
        ChartCount = 0
        os.system("del \".\\Selected\" /f /s /q & rmdir \".\\Selected\" /s /q")
        if choice == '2':
            print('Time move begins...')
            chartlist_edit = open("Chartlist.txt", mode='a', encoding='UTF-8')
            for chart in opened_charts:
                if obfuscate:
                    chart["judgeLineGroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                    chart["META"] = json.loads('{"offset":' + str(chart["META"]["offset"]) + '}')
                ChartLines = 0
                EventCount = 0
                for line in chart["judgeLineList"]:
                    print('Processing line ' + str(ChartLines) + ' in chart ' + result[ChartCount])
                    if obfuscate:
                        line["Group"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                        line["bpmfactor"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                    else:
                        line["Group"] = 0
                    line["Name"] = "Created with Ex:Phiedit v5.0"
                    try:
                        line["eventLayers"]
                        go = True
                    except:
                        go = False
                    if go:
                        for layer in line["eventLayers"]:
                            if layer != None:
                                for event_type in layer:
                                    for event in layer[event_type]:
                                        Start = event["startTime"]
                                        End = event["endTime"]
                                        new_value = Fraction(Start[0] * Start[2] + Start[1], Start[2]) + move_amount
                                        event["startTime"] = [0, new_value.numerator, new_value.denominator]
                                        new_value = Fraction(End[0] * End[2] + End[1], End[2]) + move_amount
                                        event["endTime"] = [0, new_value.numerator, new_value.denominator]
                                        if super_obf: event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                        EventCount += 1
                    try:
                        line["extended"]
                        go = True
                    except:
                        go = False
                    if go:
                        for event_type in line["extended"]:
                            for event in line["extended"][event_type]:
                                Start = event["startTime"]
                                End = event["endTime"]
                                new_value = Fraction(Start[0] * Start[2] + Start[1], Start[2]) + move_amount
                                event["startTime"] = [0, new_value.numerator, new_value.denominator]
                                new_value = Fraction(End[0] * End[2] + End[1], End[2]) + move_amount
                                event["endTime"] = [0, new_value.numerator, new_value.denominator]
                                if super_obf: event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                EventCount += 1
                    if line["numOfNotes"] > 0:
                        for note in line["notes"]:
                            Start = note["startTime"]
                            End = note["endTime"]
                            new_value = Fraction(Start[0] * Start[2] + Start[1], Start[2]) + move_amount
                            note["startTime"] = [0, new_value.numerator, new_value.denominator]
                            new_value = Fraction(End[0] * End[2] + End[1], End[2]) + move_amount
                            note["endTime"] = [0, new_value.numerator, new_value.denominator]
                            EventCount += 1
                    ChartLines += 1
                    print('Processed event: ' + str(EventCount))
                if not merge_chart: #Write to separate files if the user didn't require merging
                    if move_amount < 0:
                        os.system('xcopy \".\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '\\\" \".\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '-' + str(-(float(move_amount.numerator) / float(move_amount.denominator))) + '\\\" /c /e /v /r /y /k /i /g')
                        print('Writing to ' + '.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '-' + str(-(float(move_amount.numerator) / float(move_amount.denominator))) + '\\' + opened_chart_list[ChartCount][6].lstrip("Chart: "))
                        final_output = open('.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '-' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + '\\' + opened_chart_list[ChartCount][6].lstrip("Chart: "), mode='w', encoding='UTF-8')
                    else:
                        os.system('xcopy \".\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '\\\" \".\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '+' + str(float(move_amount.numerator) / float(move_amount.denominator)) + '\\\" /c /e /v /r /y /k /i /g')
                        print('Writing to ' + '.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '+' + str(float(move_amount.numerator) / float(move_amount.denominator)) + '\\' + opened_chart_list[ChartCount][6].lstrip("Chart: "))
                        final_output = open('.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '+' + str(float(move_amount.numerator)/float(move_amount.denominator)) + '\\' + opened_chart_list[ChartCount][6].lstrip("Chart: "), mode='w', encoding='UTF-8')
                    json.dump(chart, final_output, ensure_ascii=False)
                    final_output.close()
                    chartlist_edit.write('\n#\n')
                    if move_amount < 0:
                        chartlist_edit.write(opened_chart_list[ChartCount][2] + ' -' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + 'bars\n')
                        chartlist_edit.write(opened_chart_list[ChartCount][3] + '-' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + '\n')
                    else:
                        chartlist_edit.write(opened_chart_list[ChartCount][2] + ' +' + str(float(move_amount.numerator)/float(move_amount.denominator)) + 'bars\n')
                        chartlist_edit.write(opened_chart_list[ChartCount][3] + '+' + str(float(move_amount.numerator)/float(move_amount.denominator)) + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][4] + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][5] + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][6] + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][7] + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][8] + '\n')
                    chartlist_edit.write(opened_chart_list[ChartCount][9])
                    if move_amount < 0:
                        info = open('.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '-' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + '\\info.txt', mode='w', encoding='UTF-8')
                        info.write('#\n')
                        info.write(opened_chart_list[ChartCount][2] + ' -' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + 'bars\n')
                        info.write(opened_chart_list[ChartCount][3] + '-' + str(-(float(move_amount.numerator)/float(move_amount.denominator))) + '\n')
                    else:
                        info = open('.\\Resources\\' + opened_chart_list[ChartCount][3].lstrip("Path: ") + '+' + str(float(move_amount.numerator)/float(move_amount.denominator)) + '\\info.txt', mode='w', encoding='UTF-8')
                        info.write('#\n')
                        info.write(opened_chart_list[ChartCount][2] + ' +' + str(float(move_amount.numerator)/float(move_amount.denominator)) + 'bars\n')
                        info.write(opened_chart_list[ChartCount][3] + '+' + str(float(move_amount.numerator)/float(move_amount.denominator)) + '\n')
                    info.write(opened_chart_list[ChartCount][4] + '\n')
                    info.write(opened_chart_list[ChartCount][5] + '\n')
                    info.write(opened_chart_list[ChartCount][6] + '\n')
                    info.write(opened_chart_list[ChartCount][7] + '\n')
                    info.write(opened_chart_list[ChartCount][8] + '\n')
                    info.write(opened_chart_list[ChartCount][9])
                ChartCount += 1
                move_amount += offset_per_file
            chartlist_edit.close()
            del chartlist_edit
            del ChartCount
            del ChartLines
            del Start
            del End
            del new_value
            del EventCount
            del layer
            del event_type
            del event
            del note
            del line
            if super_obf: del go
            if not merge_chart:
                del info
                del final_output
                del opened_charts
                del opened_chart_list
                del choice
                del obfuscate
                del merge_chart
                del merge_bpm
                del super_obf
                del move_amount
                del offset_per_file
                del result
                print('\033[2J\033[0;0H', end='')
                print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
                print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
                print('\033[35mAll done! The moved charts has been exported to .\\Resources\\')
                print('Thank you for using this program. ')
                time.sleep(5)
                quit()
        ChartCount = 0
        if choice == '1' or merge_chart:
            print('Merging begins...')
            final_chart = opened_charts[0]
            final_chart["judgeLineGroup"] = ["Default"]
            if obfuscate:
                final_chart["META"] = json.loads('{"offset":' + str(final_chart["META"]["offset"]) + '}')
                final_chart["judgeLineGroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
            TotalLines = 0
            ChartLines = 0
            EventCount = 0
            for line in final_chart["judgeLineList"]:
                print('Processing line ' + str(ChartLines) + ' in chart ' + result[ChartCount].rstrip('.txt'))
                if obfuscate:
                    line["Group"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                    if line["bpmfactor"] == 0: line["bpmfactor"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                    if super_obf:
                        try:
                            line["eventLayers"]
                            go = True
                        except:
                            go = False
                        if go:
                            for layer in line["eventLayers"]:
                                if layer != None:
                                    for event_type in layer:
                                        for event in layer[event_type]:
                                            event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                            EventCount += 1
                        try:
                            line["extended"]
                            go = True
                        except:
                            go = False
                        if go:
                            for event_type in line["extended"]:
                                for event in line["extended"][event_type]:
                                    event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                    EventCount += 1
                else:
                    line["Group"] = 0
                line["Name"] = "Created with Ex:Phiedit v5.0"
                ChartLines += 1
            TotalLines += ChartLines
            opened_charts.pop(0)
            ChartCount += 1
            while len(opened_charts) > 0:
                ChartLines = 0
                EventCount = 0
                if merge_bpm:
                    final_chart["BPMList"] += opened_charts[0]["BPMList"]
                opened_charts[0]["judgeLineGroup"] = ["Default"]
                for line in opened_charts[0]["judgeLineList"]:
                    print('Processing line ' + str(ChartLines) + ' in chart ' + result[ChartCount].rstrip('.txt'))
                    if obfuscate:
                        line["Group"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                        if line["bpmfactor"] == 0: line["bpmfactor"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                        if super_obf:
                            try:
                                line["eventLayers"]
                                go = True
                            except:
                                go = False
                            if go:
                                for layer in line["eventLayers"]:
                                    if layer != None:
                                        for event_type in layer:
                                            for event in layer[event_type]:
                                                event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                                EventCount += 1
                            try:
                                line["extended"]
                                go = True
                            except:
                                go = False
                            if go:
                                for event_type in line["extended"]:
                                    for event in line["extended"][event_type]:
                                        event["linkgroup"] = "卧槽，谱面居然被神奇的Ex:Phiedit混淆了！惊不惊喜意不意外？ | Surprise, this chart was amazingly obfuscated by the magical Ex:Phiedit! Are you amazed? | なんと、楽譜は魔法のEx:Phieditによって暗号化されていたのだ！ 驚いた？"
                                        EventCount += 1
                            print('Processed event: ' + str(EventCount))
                    else:
                        line["Group"] = 0
                    line["Name"] = "Created with Ex:Phiedit v5.0"
                    if line["father"] != -1:
                        line["father"] += TotalLines
                    ChartLines += 1
                final_chart["judgeLineList"] += opened_charts[0]["judgeLineList"]
                ChartCount += 1
                opened_charts.pop(0)
            os.system('xcopy \".\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '\\\" \".\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '_merged\\\" /c /e /v /r /y /k /i /g')
            print('Writing to ' + '.\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '_merged\\' + opened_chart_list[0][6].lstrip("Chart: "))
            final_output = open('.\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '_merged\\' + opened_chart_list[0][6].lstrip("Chart: "), mode='w', encoding='UTF-8')
            json.dump(final_chart, final_output, ensure_ascii=False)
            final_output.close()
            chartlist_edit = open("Chartlist.txt", mode='a', encoding='UTF-8')
            chartlist_edit.write('\n#\n')
            chartlist_edit.write(opened_chart_list[0][2]+' (Merged)\n')
            chartlist_edit.write(opened_chart_list[0][3]+'_merged\n')
            chartlist_edit.write(opened_chart_list[0][4]+'\n')
            chartlist_edit.write(opened_chart_list[0][5]+'\n')
            chartlist_edit.write(opened_chart_list[0][6]+'\n')
            chartlist_edit.write(opened_chart_list[0][7]+'\n')
            chartlist_edit.write(opened_chart_list[0][8]+'\n')
            chartlist_edit.write(opened_chart_list[0][9])
            chartlist_edit.close()
            info = open('.\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '_merged\\info.txt', mode='w', encoding='UTF-8')
            info.write('#\n')
            info.write(opened_chart_list[0][2] + ' (Merged)\n')
            info.write(opened_chart_list[0][3] + '_merged\n')
            info.write(opened_chart_list[0][4] + '\n')
            info.write(opened_chart_list[0][5] + '\n')
            info.write(opened_chart_list[0][6] + '\n')
            info.write(opened_chart_list[0][7] + '\n')
            info.write(opened_chart_list[0][8] + '\n')
            info.write(opened_chart_list[0][9])
            info.close()
            del chartlist_edit
            del info
            del final_output
            del final_chart
            del opened_charts
            del ChartCount
            del EventCount
            del TotalLines
            del ChartLines
            del choice
            del line
            if super_obf: del go
            del obfuscate
            del merge_chart
            del merge_bpm
            del super_obf
            del move_amount
            del offset_per_file
            del result
            print('\033[2J\033[0;0H', end='')
            print('\033[1m\033[32mWelcome to Happy_mimimix\'s PEC chart merger v2.1!')
            print('\033[31mNote: Bundled with Ex:Phiedit v5.0, distributing this program separately is not allowed.\n')
            print('\033[35mAll done! The merged chart has been exported to .\\Resources\\' + opened_chart_list[0][3].lstrip("Path: ") + '_merged\\' + opened_chart_list[0][6].lstrip("Chart: "))
            print('Thank you for using this program. ')
            del opened_chart_list
            time.sleep(5)
            quit()