import os
import pandas as pd
from pydub import AudioSegment
from gtts import gTTS
from playsound import playsound
import glob


def textToSpeech(text, filename):
    mytext = str(text)
    language = 'hi'
    myobj = gTTS(text=mytext, lang=language, slow=False)
    myobj.save(filename)
    

# This function returns pydubs audio segment
def mergeAudios(audios):
    combined = AudioSegment.empty()
    for audio in audios:
        combined += AudioSegment.from_mp3(audio)
    return combined


def generateSkeleton():
    audio = AudioSegment.from_mp3('railway.mp3')

    # 1 - Generate kripya dheyan dijiye
    start = 0000
    finish = 2080
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\1_hindi.mp3", format="mp3")

    # 3 - Generate se chalkar
    start = 2000
    finish = 3265
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\3_hindi.mp3", format="mp3")

    # 5 - Generate ke raaste
    start = 3265
    finish = 4250
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\5_hindi.mp3", format="mp3")

    # 7 - Generate ko jaane wali gaadi sankhya
    start = 4400
    finish = 7090
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\7_hindi.mp3", format="mp3")

    # 9 - Generate kuch hi samay mei platform sankhya
    start = 7100
    finish = 9790
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\9_hindi.mp3", format="mp3")

    # 11 - Generate par aa rahi hai
    start = 9800
    finish = 13000
    audioProcessed = audio[start:finish]
    audioProcessed.export(f"{dirName}\\11_hindi.mp3", format="mp3")


def generateAnnouncement(filename):
    df = pd.read_excel(filename)
    print(df)
    for index, item in df.iterrows():
        # 2 - Generate from-city
        textToSpeech(item['from'], f'{dirName}\\2_hindi.mp3')

        # 4 - Generate via-city
        textToSpeech(item['via'], f'{dirName}\\4_hindi.mp3')

        # 6 - Generate to-city
        textToSpeech(item['to'], f'{dirName}\\6_hindi.mp3')

        # 8 - Generate train no and name
        textToSpeech(item['train_no'] + " " + item['train_name'], f'{dirName}\\8_hindi.mp3')

        # 10 - Generate platform number
        textToSpeech(item['platform'], f'{dirName}\\10_hindi.mp3')

        audios = [f"{dirName}\\{i}_hindi.mp3" for i in range(1, 12)]
        announcement = mergeAudios(audios)
        announcement.export(f"{dirName2}\\announcement_{item['train_no']}_{index+1}.mp3", format="mp3")


def create_playlist(path):
    for aud in glob.iglob(path):
        print(f"Playing... {aud}")
        playsound(aud)


if __name__ == "__main__":
    dirName = 'tempAudio'
    if not os.path.exists(dirName):
        os.mkdir(dirName)

    dirName2 = 'announcements'
    if not os.path.exists(dirName2):
        os.mkdir(dirName2)

    generateSkeleton()

    print("Generating Announcements for following trains...")
    generateAnnouncement("announce_hindi.xlsx")

    print("\nNow Playing Announcements...")
    for filename in os.scandir(dirName2):
        if filename.is_file() and filename.name.endswith('.mp3'):
            print(f"Playing... {filename.path}")
            playsound(filename.path)
