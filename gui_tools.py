from os import listdir
from os.path import isfile, join

def get_songs_list():
    mypath = './resources/songs'
    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

    songs = []

    for filename in onlyfiles:
        songs.append(filename.rsplit( ".", 1 )[ 0 ])

    return songs

def validate_input_data(data):
    # validate songs
    songs = get_songs_list()

    for i,song in enumerate(data['ie_songs']):
        if song not in songs:
            return [False, f"Hiba: Az igeversek előtti {i+1}. számhoz nincs dia!"]

    for i,song in enumerate(data['iu_songs']):
        if song not in songs:
            return [False, f"Hiba: Az igeversek utáni {i+1}. számhoz nincs dia!"]

    for i,song in enumerate(data['tu_songs']):
        if song not in songs:
            return [False, f"Hiba: A tanítás utáni {i+1}. számhoz nincs dia!"]

    return [True,""]