#ALBINPRAVEEN
import speech_recognition as sr
import webbrowser #opening web
import wolframalpha #mathematical operations
import time
import os #file handling
import pyperclip #clipping
import wikipedia
import win32com.client as wincl
from datetime import datetime
import pafy
from bs4 import BeautifulSoup as bs
import requests
import sys
#import pyttsx3 as tts #text to speech_recognition
v=wincl.Dispatch("SAPI.SpVoice")
cl=wolframalpha.Client('YVH8AY-R8H93LQAJ2')
att=cl.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold=0.7
r.energy_threshold=500
shell=wincl.Dispatch("WScript.Shell")
#v.Speak('Hello! For a list of commands, please say "Keyword list"...')
v.Speak('At your service Sir!')
#print('Hello! For a list of commands, please say "Keyword list"...')

#List of commands
google = 'search for'
youtube='search YouTube for'
acad = 'academic search'
wkp = 'wiki results for'
rdds = 'read the copied text'
t='what is the time'
d='what is the date'
say='say'
copy='copy the text'
sav = 'save the text'
bkmk = 'bookmark this page'
vid = 'video for'
wtis = 'what is'
wtar = 'what are'
whis = 'who is'
whws = 'who was'
when = 'when'
where = 'where'
how = 'how'
lsp = 'silence please'
lsc = 'resume listening'
stoplst = 'stop listening'
sc = 'deep search'
calc='calculate'
keywd='keyword list'
calculat='open calculator'
paint = 'open paint'
playmusic='play'
dlmusic='download song'
dlvideo='download video'
mgntlink='magnet link for'


while True:
    with sr.Microphone() as source:
        try:
            v.Speak("How can I help you today?")
            print("Waiting for your command")
            audio = r.listen(source, timeout=None)
            message = str(r.recognize_google(audio))
            print('You said: ' + message)
            if google in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print('Google results for: '+str(st))
                url='https://google.com/search?q='+st
                webbrowser.open(url)
                v.Speak('Google Results for: '+str(st))
            elif mgntlink in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                query=str(st)
                url='https://pirateproxy.mx/search/'+query+'/0/99/0'
                print("Searching......")
                source=requests.get(url).text
                soup=bs(source,'lxml')
                results=soup.find_all('div',class_='detName')
                i=1
                for r in results:
                    print(i,r.text)
                    i=i+1
                print("Enter the Serial Number of the search item you like to download: ")
                v.Speak("Enter the Serial Number of the search item you like to download: ")
                choice=int(input())
                print ("Fetching data.....")
                v.Speak ("Fetching data.....")
                magnet_results=soup.find_all('a',title='Download this torrent using magnet')
                a=[]
                for m in magnet_results:
                    a.append(m['href'])
                magnet_link=(a[choice-1])
                print("Magnet Link of your selected choice has been fetched.")
                pyperclip.copy(magnet_link)
                v.Speak ("Your magnet link is now in your clipboard.")
            elif acad in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print('Academic results for: '+str(st))
                url='https://scholar.google.com/scholar?q='+st
                webbrowser.open(url)
                v.Speak('Academic Results for: '+str(st))
            elif wkp in message:
                try:
                    words=message.split()
                    del words[0:3]
                    st=' '.join(words)
                    wkpres=wikipedia.summary(st,sentences=2)
                    try:
                        print('\n'+str(wkpres)+'\n')
                        v.Speak(wkpres)
                    except UnicodeEncodeError:
                        v.Speak("Sorry! Please try searching again")
                except wikipedia.exceptions.DisambiguationError as e:
                    print(e.options)
                    v.Speak("Too many results for this keyword. Please be more specific and retry")
                    continue
                except wikipedia.exceptions.PageError as e:
                    print("This page doesn't exist")
                    v.Speak("No results found for: "+str(st))
                    continue
            elif rdds in message:
                print('Reading your text')
                words=message.split()
                del words[0:4]
                v.Speak(pyperclip.paste())

            elif say in message:
                words=message.split()
                del words[0:1]
                st=' '.join(words)
                print("Repeating the text: "+str(st))
                v.Speak('Alright, Saying..: '+ str(st))
            elif copy in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                pyperclip.copy(st)
                print("The text "+str(st)+" is now copied to clipboard!")
                v.Speak('The text ..'+str(st)+' is now in your clipboard... Happy Pasting!')
            elif youtube in message:
                words=message.split()
                del words[0:3]
                st=' '.join(words)
                print("Searching for "+str(st)+" on Youtube")
                url='https://www.youtube.com/results?search_query='+str(st)
                webbrowser.open(url)
                v.Speak('Youtube Search results for: '+str(st))
            elif vid in message:
                words=message.split()
                del words[0:2]
                st=' '.join(words)
                print("Searching for "+str(st)+" on Youtube")
                url='https://www.youtube.com/results?search_query='+str(st)
                webbrowser.open(url)
                v.Speak('Youtube Search results for: '+str(st))
            elif t in message:
                c=time.ctime()
                words=c.split()
                v.Speak("The time is: "+str(words[3]))
            elif d in message:
                c=time.ctime()
                words=c.split()
                v.Speak("Today, it is"+words[2]+words[1]+words[4])
            elif sc in message:
                try:
                    words=message.split()
                    del words[0:2]
                    st=' '.join(words)
                    scq=cl.query(st)
                    sca=next(scq.results).text
                    print('The answer is: '+str(sca))
                    url='https://www.wolframalpha.com/input/?i='+str(st)
                    v.Speak('The answeris: '+str(sca))
                except StopIteration:
                    print('Your question is ambiguous. Please try with another keyword!')
                    v.Speak('Your question is ambiguous. Please try with another keyword!')
                except Exception as e:
                    print(e)
                    v.Speak(e)
                else:
                    v.Speak("I'm always correct")

            elif calc in message:
                try:
                    words=message.split()
                    del words[0:1]
                    st=' '.join(words)
                    scq=cl.query(st)
                    sca=next(scq.results).text
                    print('The answer is: '+str(sca))
                    url='https://www.wolframalpha.com/input/?i='+str(st)
                    v.Speak('The answer is: '+str(sca))
                except StopIteration:
                    print('Your question is ambiguous. Please try with another keyword!')
                    v.Speak('Your question is ambiguous. Please try with another keyword!')
                except Exception as e:
                    print(e)
                    v.Speak(e)
                else:
                    v.Speak("I'm always correct")
            elif paint in message:
                os.system('mspaint')
            elif sav in message:
                print('Saving your text to file')
                with open('path to your text file','a') as f:
                    f.write(pyperclip.paste())
                    v.Speak('File is successfully saved')
            elif bkmk in message:
                shell.SendKeys("^d")
                v.Speak("Alright, Page Bookmarked!")
            elif calculat in message:
                os.system('calc')
            elif 'open' in message:
                words=message.split()
                del words [0:1]
                st=' '.join(words)
                if 'telegram' in str(st):
                    print('Opening Telegram')
                    v.Speak('Opening Telegram')
                    os.startfile(r'''C:\Users\asus1\AppData\Roaming\Telegram Desktop\Telegram.exe''')
                elif 'Chrome' in str(st):
                    print('Opening Chrome')
                    v.Speak('Opening Chrome')
                    os.startfile(r'''C:\Program Files (x86)\Google\Chrome\Application\chrome.exe''')
            elif wtis in message:
                try:
                    scq=cl.query(message)
                    sca=next(scq.results).text
                    print('The answer is: '+str(sca))
                    #url='https://www.wolframalpha.com/input/?i='+st
                    #webbrowser.open(url)
                    v.Speak('The answer is: '+str(sca))
                except UnicodeEncodeError:
                    v.Speak('The answer is: '+str(sca))
                except StopIteration:
                    words=message.split()
                    del words[0:2]
                    st=' '.join(words)
                    print('Google results for: '+str(st))
                    url='https://google.com/search?q='+st
                    webbrowser.open(url)
            elif whis in message:
                try:
                    scq=cl.query(message)
                    sca=next(scq.results).text
                    print('\nThe answer is: '+str(sca)+'\n')
                    v.Speak('The answer is: '+str(sca))
                except StopIteration:
                    try:
                        words=message.split()
                        del words[0:2]
                        st=' '.join(words)
                        wkpres=wikipedia.summary(st,sentences=2)
                        print('\n'+str(wkpres)+'\n')
                        v.Speak(wkpres)
                    except UnicodeEncodeError:
                        v.Speak(wkpres)
                    except:
                        words=message.split()
                        del words[0:2]
                        st=' '.join(words)
                        print('Google results for: '+str(st))
                        url='https://google.com/search?q='+st
                        webbrowser.open(url)
                        v.Speak('Google Results for: '+str(st))
            elif wtar in message:
                try:
                    scq=cl.query(message)
                    sca=next(scq.results).text
                    print('The answer is: '+str(sca))
                    v.Speak('The answer is: '+str(sca))
                except UnicodeEncodeError:
                    v.Speak('The answer is: '+str(sca))
                except StopIteration:
                    words=message.split()
                    del words[0:2]
                    st=' '.join(words)
                    print('Google results for: '+str(st))
                    url='https://google.com/search?q='+st
                    webbrowser.open(url)
            elif playmusic in message:
                words=message.split()
                del words[0:1]
                song_name=str(' '.join(words))
                base = "https://www.youtube.com/results?search_query="
                r = requests.get(base+song_name)
                page = r.text
                soup=bs(page,'html.parser')
                vids = soup.findAll('a',attrs={'class':'yt-uix-tile-link'})
                videolist=[]
                for v in vids:
                    tmp = 'https://www.youtube.com' + v['href']
                    videolist.append(tmp)
                v = pafy.new(videolist[0])
                song=str(v.title)
                print('Checking for best available quality: '+song)
                best=v.getbest(preftype="webm")
                best_url=best.url
                webbrowser.open(best_url)
            elif dlvideo in message:
                words=message.split()
                del words[0:2]
                song_name=str(' '.join(words))
                base = "https://www.youtube.com/results?search_query="
                r = requests.get(base+song_name)
                page = r.text
                soup=bs(page,'html.parser')
                vids = soup.findAll('a',attrs={'class':'yt-uix-tile-link'})
                videolist=[]
                for v in vids:
                    tmp = 'https://www.youtube.com' + v['href']
                    videolist.append(tmp)
                v = pafy.new(videolist[0])
                song=str(v.title)
                print('Checking for best available video quality for: '+song)
                best=v.getbest()
                best.download(quiet=False)

            elif dlmusic in message:
                words=message.split()
                del words[0:2]
                song_name=str(' '.join(words))
                base = "https://www.youtube.com/results?search_query="
                r = requests.get(base+song_name)
                page = r.text
                soup=bs(page,'html.parser')
                vids = soup.findAll('a',attrs={'class':'yt-uix-tile-link'})
                videolist=[]
                for v in vids:
                    tmp = 'https://www.youtube.com' + v['href']
                    videolist.append(tmp)
                v = pafy.new(videolist[0])
                song=str(v.title)
                print('Checking for best available audio quality for: '+song)
                best=v.getbestaudio()
                best.download()


            elif keywd in message:
                print('')
                print('Say " ' + google + ' " to return a Google search')
                print('Say " ' + acad + ' " to return a Google scholar search')
                print('Say " ' + sc + ' " to return a wolframalpha query')
                print('Say " ' + wkp + ' " to return a wikipedia page')
                print('Say " ' + rdds + ' " to read the text you have highlighted and Ctrl + C (Copied to clipboard)')
                print('Say " ' + sav + ' " to read the save you have highlighted and Ctrl + C-ed (Copied to clipboard) to a file')
                print('Say " ' + bkmk + ' " to bookmark the current page')
                print('Say " ' + vid + ' " to return the video results for the query')
                print("For more general questions, ask them normally and i will do my best to find a appropriate answer for you!")
                print('Say '+stoplst+' to shut down')
                print('')
#                print('Say " ' + book + ' " to return an Amazon Book Search')

        except:
            break
        finally:
            break
