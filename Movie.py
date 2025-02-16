import os
import re
import pandas as pd
from collections import defaultdict


def extract_movie_details(filename, file_path):
    """Extracts movie details including name, year, format, encoding, language, audio, season, episode, and size."""
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
    """Processes MKV files from the given directory and categorizes them into unique and duplicate movies."""
    movie_dict = {}
    duplicates = []
    movie_data = []

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".mkv"):
                file_path = os.path.join(root, file)
                movie_details = extract_movie_details(file, file_path)

                if movie_details[0] and movie_details[1]:  # Ensure movie name and year exist
                    movie_name_key = (movie_details[0], movie_details[1])  # (Movie Name, Year) as key

                    if movie_name_key in movie_dict:
                        # If the movie is already in dict, add **both** original & duplicate
                        if movie_dict[movie_name_key] not in duplicates:
                            duplicates.append(movie_dict[movie_name_key])  # Add first occurrence to duplicates
                        duplicates.append(movie_details)  # Add new duplicate entry

                    movie_dict[movie_name_key] = movie_details  # Store movie details
                    movie_data.append(movie_details)  # Store all movie entries

    return movie_data, duplicates


def save_to_excel(movie_data, duplicates):
    """Saves movie data into Excel sheets while preserving existing records."""
    movies_path = os.path.abspath("Movies.xlsx")
    duplicates_path = os.path.abspath("Duplicates.xlsx")

    # Load existing Movies.xlsx if it exists
    df_existing = pd.read_excel(movies_path) if os.path.exists(movies_path) else pd.DataFrame()

    # Save movie data
    df_new = pd.DataFrame(movie_data,
                          columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages', 'Audio Format', 'Season',
                                   'Episode', 'Size'])
    df_movies = pd.concat([df_existing, df_new]).drop_duplicates().reset_index(drop=True)
    df_movies.to_excel(movies_path, index=False)
    print(f"Movies.xlsx updated at: {movies_path}")

    # Save duplicates if they exist
    if duplicates:
        df_existing_dup = pd.read_excel(duplicates_path) if os.path.exists(duplicates_path) else pd.DataFrame()

        df_duplicates_new = pd.DataFrame(duplicates, columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages',
                                                              'Audio Format', 'Season', 'Episode', 'Size'])
        df_duplicates = pd.concat([df_existing_dup, df_duplicates_new]).drop_duplicates().reset_index(drop=True)

        df_duplicates.to_excel(duplicates_path, index=False)
        print(f"Duplicates.xlsx updated at: {duplicates_path}")
    else:
        print("No duplicates found. Duplicates.xlsx not created.")


if __name__ == "__main__":
    directory = r"E:\Movies"
    movie_data, duplicates = process_movie_files(directory)
    save_to_excel(movie_data, duplicates)
    print("Processing complete. Check Movies.xlsx and Duplicates.xlsx.")
