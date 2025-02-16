import os
import re
import pandas as pd


def extract_movie_details(filename, file_path):
    filename = re.sub(r'www\.[\w.-]+ - ', '', filename)

    format_match = re.search(r'(\d{3,4}p|4K|2160p|HDRip|BluRay|HQ HDRip)', filename, re.IGNORECASE)
    video_format = format_match.group(1) if format_match else None

    encoding_match = re.search(r'(x264|x265|HEVC|AVC|SDR|HDR10\+?)', filename, re.IGNORECASE)
    encoding = encoding_match.group(1) if encoding_match else "x264"

    language_match = re.search(r'\[(.*?)\]', filename)
    bracketed_languages = language_match.group(1) if language_match else None

    inline_language_match = re.findall(
        r'\b(Tamil|Telugu|Hindi|Malayalam|Kannada|Bengali|Marathi|Punjabi|Gujarati|Odia|Urdu|English)\b', filename,
        re.IGNORECASE)

    language_set = set()
    if bracketed_languages:
        language_set.update(bracketed_languages.split(" + "))
    if inline_language_match:
        language_set.update(inline_language_match)

    languages = " + ".join(sorted(language_set)) if language_set else None

    audio_matches = re.findall(r'(DD\+5\.1|DD\+|DD plus|AAC|Dolby Atmos|2\.1|\d{3,4}Kbps)', filename, re.IGNORECASE)
    audio_format = ' & '.join(set(audio_matches)) if audio_matches else None

    season_match = re.search(r'[Ss](\d{1,2})[\sEeP]*(\d{1,2})', filename, re.IGNORECASE)
    season = season_match.group(1) if season_match else None
    episode = season_match.group(2) if season_match else None

    try:
        file_size_bytes = os.path.getsize(file_path)
        file_size = round(file_size_bytes / (1024 * 1024 * 1024), 2)
        file_size = f"{file_size} GB" if file_size >= 1 else f"{round(file_size_bytes / (1024 * 1024), 2)} MB"
    except FileNotFoundError:
        file_size = None

    filename = re.sub(r'\[.*?\]', '', filename).strip()

    match = re.search(r'(.+?) \((\d{4})\)', filename)
    if match:
        movie_name = match.group(1).strip()
        year = match.group(2)
    else:
        year_match = re.search(r'\b(19\d{2}|20\d{2})\b', filename)
        year = year_match.group(1) if year_match else None
        movie_name = filename.split(year)[0].strip() if year else filename.strip()

    return movie_name, year, video_format, encoding, languages, audio_format, season, episode, file_size


def process_movie_files(directory):
    movies_path = "Movies.xlsx"
    duplicates_path = "Duplicates.xlsx"

    if os.path.exists(movies_path):
        df_movies = pd.read_excel(movies_path)
    else:
        df_movies = pd.DataFrame(
            columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages', 'Audio Format', 'Season', 'Episode',
                     'Size'])

    if os.path.exists(duplicates_path):
        df_duplicates = pd.read_excel(duplicates_path)
    else:
        df_duplicates = pd.DataFrame(
            columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages', 'Audio Format', 'Season', 'Episode',
                     'Size'])

    new_movies = []
    new_duplicates = []

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".mkv"):
                file_path = os.path.join(root, file)
                movie_details = extract_movie_details(file, file_path)
                movie_name = movie_details[0]

                if movie_name in df_movies['Movie Name'].values:
                    existing_entries = df_movies[df_movies['Movie Name'] == movie_name]
                    df_duplicates = pd.concat(
                        [df_duplicates, existing_entries, pd.DataFrame([movie_details], columns=df_movies.columns)],
                        ignore_index=True)
                else:
                    new_movies.append(movie_details)

    return new_movies, df_movies, df_duplicates


def save_to_excel(new_movies, df_movies, df_duplicates):
    movies_path = "Movies.xlsx"
    duplicates_path = "Duplicates.xlsx"

    df_new_movies = pd.DataFrame(new_movies,
                                 columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages', 'Audio Format',
                                          'Season', 'Episode', 'Size'])
    df_movies = pd.concat([df_movies, df_new_movies], ignore_index=True)
    df_movies.to_excel(movies_path, index=False)
    print(f"Movies.xlsx updated at: {movies_path}")

    df_duplicates.to_excel(duplicates_path, index=False)
    print(f"Duplicates.xlsx updated at: {duplicates_path}")


if __name__ == "__main__":
    directory = r"F:\MOVIES\1"
    new_movies, df_movies, df_duplicates = process_movie_files(directory)
    save_to_excel(new_movies, df_movies, df_duplicates)
    print("Processing complete. Check Movies.xlsx and Duplicates.xlsx.")
