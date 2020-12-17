#start des Python Projekts
#imports
import random
import pyttsx3  #Text to spech converter
import speech_recognition as sr #modul um die Stimme zu erkennen
import webbrowser #um Webseiten zu öffnen
import os
import datetime
import win32com.client as win32
import time
import concurrent.futures  #für Threading
from turtle import *

while True:     #kann nur mit dem "turn off" command beendet werden

    engine = pyttsx3.init()
    rate = engine.getProperty('rate')       #einstellungen für die engine
    engine.setProperty('rate', 142)         #geschwindikeit der stimme
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id) #auswählen der englischen stimme   #(funktioniert trotz warnung)


    def speak(audio):       #speak funktion definieren
        engine.say(audio)
        engine.runAndWait()

    def wishme():                                   #Begrüssung je nach Tageszeit
        myfile = open("user.txt")
        name = myfile.read()
        hour = int(datetime.datetime.now().hour)
        if hour >= 0 and hour < 12:
            speak("Good morning" + name + ",how can i help you?")          #begrüssung mit usernamen
        elif hour >= 12 and hour < 18:
            speak("Good afternoon" + name + "how can i help you?")
        else:
            speak("Good night" + name + "how can i help you?")


    def takecommand():  #eingabe durch stimme
        r = sr.Recognizer()
        with sr.Microphone() as source:  #öffnet mikrofon
            print("Hearing...")
            #printtinker("Hearing...")
            try:
                audio = r.listen(source, timeout=7, phrase_time_limit=10)  #Timeout nach 7 sekunden damit die Threads nicht unnötig gefüllt sind (verringert die respond time)
                print("Recognizing...")
                #printtinker("Recognizing...")
                query = r.recognize_google(audio, language='en-in')  #google zum erkennen
                print(f"User said: {query}\n")
            except Exception as e:
                print("Say that again please...")  #Wenn nicht verstanden "say that again pls"
                #printtinker("Say that again please...")
                return "None"
            except TimeoutError:   #Timeout error exeption damit das Programm nach timeout nicht abstürzt
                return "None"
            return query


    #anfang commands
    #Threading
    with concurrent.futures.ThreadPoolExecutor() as executor:       #es werden drei höhr Prozesse gestartet
        f1 = executor.submit(takecommand)
        time.sleep(2)                                               #diese zu einer unterschiedlichen Zeit
        f2 = executor.submit(takecommand)                           #damit möglichst zu jeder Zeit zugehört wird
        time.sleep(4)
        f3 = executor.submit(takecommand)
        query1 = f1.result()
        query2 = f2.result()
        query3 = f3.result()
    #sucht keyword
    if "hey Pai" in query1 or "hey Pai" in query2 or "hey Pai" in query3:  # in jeder query wird nach dem "keyword" geschaut
        del query1
        del query2  #löschen der "query" damit wieder für neue Eingabe frei ist
        del query3
        ##########
        wishme()  #begrüssung
        query = takecommand()  #erste abfrage nach keyword
        #google command
        if "Google" in query:
            speak("Opening google")
            webbrowser.open_new_tab("https://google.com")
        #youtube command
        if "YouTube" in query:
            speak("Opening Youtube...")
            webbrowser.open_new_tab("https://youtube.com")
        #jebaited command
        if "jebaited" in query:
            speak("Hahaha Jebaited!")
            webbrowser.open_new_tab("https://www.youtube.com/watch?v=dQw4w9WgXcQ")
        #email command
        if "email" in query:
            speak("opening email")
            os.startfile("outlook")
        #quote command
        if "quote" in query:
            line = random.choice(open("quotes.txt").readlines())
            speak(line)
        #write email
        if "spam" in query:
            speak("Spamming")
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = "Meinrad.Buergler@bbcag.ch"
            mail.Subject = "Hallo Meini"
            mail.Body = 'Message body'
            mail.Send()
        #ask for help
        if "ask" in query:
            speak("asking for help")
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = "Meinrad.Buergler@bbcag.ch"
            mail.Subject = "Ich Bitte dich um Hilfe"
            mail.Body = 'Message body'
            mail.Send()
        #user name
        if "user" in query:
            speak("Whats your name?")
            del query
            query = takecommand()
            MyFile = open("user.txt", "w")
            MyFile.write(query)
            MyFile.close()
        #tell time
        if "time" in query:
            time = datetime.datetime.now().strftime("%H:%M")
            speak("it is" + time)
        #help command öffnet befehlliste
        if "help" in query:
            os.startfile("help.txt")
        #singing command
        if "sing for me" in query:
            myFile = open("song.txt")
            MyText = myFile.read()
            speak(MyText)
        #turn off
        if "turn off" in query:
            quit()
        #music command
        if "music" in query:
            url = random.choice(open("song list.txt").readlines())
            speak("enjoy your music")
            webbrowser.open_new_tab(url)
        #note command
        if "note" in query:
            del query
            speak("what do you wanna note?")
            query = takecommand()
            myfile = open("note.txt", "w")
            myfile.write(query)
            myfile.close()
        #open notes/todos
        if "open list" in query:
            os.startfile("note.txt")
        #coinflip game
        if "coin flip" in query:
            del query
            speak("Welcome to the Coinflip game!")
            time.sleep(0.5)
            speak("if you want to pick Head say one, but if you want to pick tail say two")
            query = takecommand()
            side = [1, 2]
            chossed = "{}".format(random.choice(side))
            if query == chossed:
                speak("You won! congratulations!")
                time.sleep(0.1)
                speak("It was: " + chossed)
            if query != chossed:
                speak("No you lose, dont cry this can happen some times")
                time.sleep(0.1)
                speak("It was: " + chossed)
        #draw game
        if "draw" in query:
            speak("Opening draw site")
            try:
                width = 505
                height = 519
                screen = Screen()
                screen.setup(width, height)
                pensize(2)
                speed(10)
                up()
                goto(-250, -250)
                setheading(90)
                down()
                for i in range(4):
                    fd(110)
                    rt(90)
                up()
                goto(-245, -155)
                write("Commands: ", font=("Arial", 8, "bold"))
                goto(-245, -170)
                write("Bewegen: Pfeiltasten")
                goto(-245, -185)
                write("Stift hoch: u")
                goto(-245, -200)                              #für commandliste unten links
                write("Stift runter: d")
                goto(-245, -215)
                write("Schliessen: q, Esc")
                goto(-245, -230)
                write("Farbe wechseln: f")

                goto(0, 0)
                down()
                Screen()

                def key1():
                    fd(5)

                def key2():
                    lt(5)

                def key3():
                    rt(5)

                def key4():
                    bk(5)

                def key5():
                    up()

                def key6():
                    down()

                def key7():
                    Screen().bye()

                def key8():
                    Screen().bye()

                def key10():
                    colors = ["blue", "red", "black", "purple", "magenta", "orange"]
                    choose = "{}".format(random.choice(colors))
                    pencolor(choose)

                def onkeylisten():
                    onkey(key1, "Up")
                    onkey(key2, "Left")
                    onkey(key3, "Right")
                    onkey(key4, "Down")
                    onkey(key5, "u")
                    onkey(key6, "d")
                    onkey(key7, "q")
                    onkey(key8, "Escape")
                    onkey(key10, "f")

                listen(onkeylisten())
                mainloop()
            except Exception as e:
                print("Error! Try again later")
                time.sleep(5)
            #ende