import os
import subprocess
import math

import PySimpleGUI as sg
from mido import MidiFile

exe_path = "playbae.exe"  # playbae.exe path
hsb = ""  # bank path -p argument
music = ""  # file path, MIDI -f
output = ""  # WAV output path -o
samplerate = 44100  # sample rate output -mr
n_loops = 0  # number of loops for the file -l
volume = 100  # volume -v
reverb = 1  # set default reverb type -rv (0 to 11), default being none
max_voices = 64  # max number of voices

lp_list = list(range(0, 21))  # list of numbers for looping the midi track
vol_list = list(range(0, 201))  # list of numbers to set maximum volume
voice_list = list(range(1, 512))  # list of numbers to set maximum polyphony
mid_tuple = (("MID", "*.mid"), ("RMF", "*.rmf"),)
bnk_tuple = (("HSB", "*.hsb"), ("BSN", "*.bsn"), ("GM", "*.gm"),)
rvb_tuple = (("0: Default"),
             ("1: None"),
             ("2: Igor's Closet"),
             ("3: Igor's Garage"),
             ("4: Igor's Acoustic Lab"),
             ("5: Igor's Cavern"),
             ("6: Igor's Dungeon"),
             ("7: Small reflections (WebTV)"),
             ("8: Early reflections (variable verb)"),
             ("9: Basement (variable verb)"),
             ("10: Banquet hall (variable verb)"),
             ("11: Catacombs (variable verb)"),)


def secs_to_time(seconds):  # convert seconds to time format mm:ss
    seconds = seconds % (24 * 36000)
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60

    return (str(math.trunc(minutes)) + ':' + f"{math.trunc(seconds):02}")


def kill_task(subp):  # kill playbae.exe from subprocess to stop the music (only in Windows i think)
    os.popen('TASKKILL /PID ' + str(subp.pid) + ' /F 2>NUL')


def chk_box(q, tp):  # simple elseif for -q (quiet mode) and -2p (2-point interpolation) arguments
    if q and tp:
        return "-q -2p"
    elif q is False and tp:
        return "-2p"
    elif q and tp is False:
        return "-q"
    else:
        return ""


sg.theme('DarkBlack1')  # theme of the project


layout = [[sg.Text(text="///////////////  PLAYBAE  ///////////////")],
          [sg.Text("Choose a .mid: "), sg.Input('', key='-MID-', readonly=True), sg.FileBrowse(file_types=(mid_tuple))],
          [sg.Text("Choose a bank: "), sg.Input('', key='-BANK-', readonly=True), sg.FileBrowse(file_types=(bnk_tuple))],

          [sg.Button("Play", key='-PLAY-'), sg.Button("Stop", key='-STOP-'), sg.Text("Total Time:", key="-TEXT1-"),
           sg.Text("0:00", key="-TIMETRACK-")],

          [sg.Text("Sample Rate Output: "), sg.Input(samplerate, key='-SROUT-', size=(6, 1)), sg.Text("Number of loops: "),
           sg.Spin(lp_list, initial_value=n_loops, readonly=True, key='-LOOPS-', size=(3, 1))],

          [sg.Text("Max volume %: "), sg.Spin(vol_list, initial_value=volume, readonly=True, key='-VOL-', size=(3, 1)),
           sg.Text("Reverb:"), sg.Combo(values=rvb_tuple, size=(28, 1), default_value="1: None", key="-REVCOMBO-")],

          [sg.Text("--------- Additional Settings")],

          [sg.Text("Polyphony:"), sg.Spin(voice_list, initial_value=max_voices, readonly=True, key='-VOICES-', size=(3, 1)),
           sg.Checkbox("quiet mode", key="-QUIET-"), sg.Checkbox("2-point interpolation", key="-TWOPOINT-")],

          [sg.Button("Export WAV", key="-EXPORT-"), sg.Input("", key='-EXPWAV-', readonly=True), sg.FolderBrowse()]
          ]

window = sg.Window('playBAE GUI', layout, size=(540, 300))  # initialize and set the size of the window

while True:
    event, values = window.read()

    if event == '-STOP-':  # if the stop button is pressed
        try:
            kill_task(sp)
        except Exception:
            pass  # nothing - just to ignore annoying errors
        window["-TIMETRACK-"].update("0:00")

    if event == '-PLAY-':  # if the play button is pressed
        music = values["-MID-"]  # get the path of the midi file
        hsb = values["-BANK-"]  # get the path of the bank
        samplerate = values["-SROUT-"]  # get the value of the sample rate
        if music == "":  # if there's no path
            sg.popup("You must select a midi file")
        if hsb == "":  # if there's no path
            sg.popup("You must select a bank file")

        if samplerate.isnumeric() is False:  # if sample rate value is gibberish
            sg.popup("The sample rate is invalid")

        if music != "" and hsb != "" and samplerate.isnumeric():  # if everything is okay
            n_loops = values["-LOOPS-"]  # get loop value
            max_voices = values["-VOICES-"]  # get polyphony value
            qt = window.find_element("-QUIET-").get()  # get quiet mode status
            tp = window.find_element("-TWOPOINT-").get()  # get 2-point interpolation status

            mid_file = MidiFile(music)  # load the midi file
            totalSecs = mid_file.length  # get the total time in seconds
            window["-TIMETRACK-"].update(
                secs_to_time(totalSecs * (int(n_loops) + 1)))  # calculates the time multiplied by the loop

            volume = values["-VOL-"]  # get the volume value
            if values["-REVCOMBO-"] in rvb_tuple:  # get the index number from reverb's combo
                reverb = rvb_tuple.index(values["-REVCOMBO-"])
                window.find_element("-REVCOMBO-").update(set_to_index=reverb)

            # print(exe_path, "-p", hsb, "-f", music, "-mr", str(samplerate), "-l", str(n_loops), "-v", str(volume), "-rv",
            #       str(reverb), "-mv", str(max_voices), chk_box(qt, tp))

            # send the arguments to playbae.exe
            sp = subprocess.Popen([exe_path, "-p", hsb, "-f", music, "-mr", str(samplerate), "-l", str(n_loops), "-v",
                                   str(volume), "-rv", str(reverb), "-mv", str(max_voices), chk_box(qt, tp)])

    if event == "-EXPORT-":
        export_path = values["-EXPWAV-"]
        if export_path != "":  # if the export path is not empty

            # i duplicated the code because if i convert it into a function i can't kill the subprocess

            music = values["-MID-"]
            hsb = values["-BANK-"]
            samplerate = values["-SROUT-"]
            if music == "":
                sg.popup("You must select a midi file")
            if hsb == "":
                sg.popup("You must select a bank file")

            if samplerate.isnumeric() is False:
                sg.popup("The sample rate is invalid")

            if music != "" and hsb != "" and samplerate.isnumeric():
                filename = os.path.basename(music)  # get the filename from midi path
                filenoext = os.path.splitext(filename)[0]  # get the filename without extension
                n_loops = values["-LOOPS-"]
                max_voices = values["-VOICES-"]
                qt = window.find_element("-QUIET-").get()
                tp = window.find_element("-TWOPOINT-").get()
                mid_file = MidiFile(music)
                totalSecs = mid_file.length
                roundSecs = math.trunc(totalSecs)
                window["-TIMETRACK-"].update(
                    secs_to_time(totalSecs * (int(n_loops) + 1)))  # calculates the time multiplied by the loop

                volume = values["-VOL-"]
                if values["-REVCOMBO-"] in rvb_tuple:
                    reverb = rvb_tuple.index(values["-REVCOMBO-"])
                    window.find_element("-REVCOMBO-").update(set_to_index=reverb)

                # print(exe_path, "-p", hsb, "-f", music, "-mr", str(samplerate), "-l", str(n_loops), "-v", str(volume), "-rv",
                #       str(reverb), "-mv", str(max_voices), chk_box(qt, tp) + "-o", str(export_path + "/" + filenoext + ".wav"))

                # send the arguments to playbae.exe with -o (output) as .WAV
                sp = subprocess.Popen(
                    [exe_path, "-p", hsb, "-f", music, "-mr", str(samplerate), "-l", str(n_loops), "-v", str(volume), "-rv",
                     str(reverb), "-mv", str(max_voices), chk_box(qt, tp) + "-o", str(export_path + "/" + filenoext + ".wav")])
        else:  # otherwise
            sg.popup("You must specify a path folder")


    if event == sg.WINDOW_CLOSED:
        try:
            kill_task(sp)
        except Exception:
            pass  # nothing - just to ignore annoying errors
        break

window.close()
