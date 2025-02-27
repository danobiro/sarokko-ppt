from os import listdir
from os.path import isfile, join

def get_songs_list():
    mypath = './resources/songs'
    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

    songs = []

    for filename in onlyfiles:
        songs.append(filename.rsplit( ".", 1 )[ 0 ])

    return songs