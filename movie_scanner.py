import re
import pandas as pd
import os
from collections import OrderedDict
from pathlib import Path
from pandas import ExcelWriter
import excel_organize

def info_extract(tup):
    """ Extract a movie file's title, date, season/episode, and file type
    Args:
        title -- a tuple (movie title, movie path)
    Returns:
        an OrderedDict containing info for a single movie
    """
    # Initialize a dictionary for the final title
    movie = OrderedDict()
    movie['path'] = tup[1]
    title = tup[0]
    if 'sample' in title.lower():
        return None
    for symbol in ['-', r'(', r')']:
        title = title.replace(symbol, '')
    title = title.replace('..', '.')
    mo = re.search('(\w+\.)+\d\d\d(\d)?p|(\w+\.)+S\d\d|(\w+\.)+\d\d\d\d|(\w+\.)+part(\.)?(\d)?\d',
                   title)
    if mo:
        pieces = mo.group().split('.')
        info = ' '.join(pieces)
    else:
        info = title
    # Get Title
    pure_title = info
    regexes = ['S\d\d(\s)?E\d\d(\s\w+)*',
               'E\d\d(\s\w+)*',
               'S\d\d(\s\w+)*',
               's\d\de\d\d(\s\w+)*',
               '\d\d\d(\d)?p(\s\w+)*',
               'part\s(\d)?\d(\s\w+)*',
               '\d\d\d\d(\s\w+)*']
    for extra_text in regexes:
        pure_title = re.sub(extra_text, '', pure_title, re.IGNORECASE).strip()
    pure_title = re.sub('.mp4|.mkv|.avi', '', pure_title)
    movie['title'] = pure_title
    # Get Season
    season = re.search('S(\d\d)', title, re.IGNORECASE)
    if season:
        movie['season'] = season.group().lstrip('0')
    else:
        movie['season'] = None
    # Get episode
    episode = re.search('E(\d\d)|Part(\.)?(\d)?(\d)', title, re.IGNORECASE)
    if episode:
        movie['episode'] = episode.group().lstrip('0')
    else:
        movie['episode'] = None
    # Get Year
    year = re.findall('\d\d\d\d', title)
    year = [year for year in year if int(year) > 1500]
    if len(year) == 1:
        movie['year'] = int(year[0])
    else:
        movie['year'] = None
    # Get File Type
    file_type = re.findall('\.[a-z]{2}\w', title)[-1]
    movie['file_type'] = file_type
    return movie


def get_file_paths(directory='M:'):
    # https://stackoverflow.com/questions/9816816/get-absolute-paths-of-all-files-in-a-directory
    """ Retrieve all of the movie files in the directory
    Args:
        path -- the directory that contains the movies or movie folders
    Returns:
        a list of tuples containing each movie title and path
    """
    for dirpath, _, filenames in os.walk(directory):
        for f in filenames:
            path = Path(os.path.abspath(os.path.join(dirpath, f)))
            if path.suffix in ['.mp4', '.mkv', '.avi']:
                yield (f.replace(' ', '.'), path)


if __name__ == "__main__":
    directory = input(r"Enter the absolute path of the directory you wish to scan (for example, 'M:' or 'C:\Users\peter\Downloads'): ")
    print('Scanning...')
    movies = (info_extract(tup) for tup in get_file_paths(directory) if info_extract(tup))
    data = pd.DataFrame(movies)
    writer = ExcelWriter('Movies.xlsx')
    data.to_excel(writer, 'Sheet1', index=False)
    writer.save()
    print('Scan complete, spreadsheet size: {} rows X {} columns'.format(data.shape[0], data.shape[1]))
    print('Tidying up spreadsheet...')
    excel_organize.excel_organize(data.shape[0]+1)
    print('Tasks complete.')