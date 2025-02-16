import os
import re
import pandas as pd
from collections import defaultdict


def extract_movie_details(filename, file_path):
    """Extracts movie name, year, format (720p, 1080p, 4K, etc.), encoding (x264, x265, HEVC, AVC, SDR, HDR10, HDR10+), language, audio format, season, episode, and file size."""
    # Remove website prefixes like 'www.1TamilBlasters.cool - '
    filename = re.sub(r'www\.[\w.-]+ - ', '', filename)

    # Extract format (e.g., 720p, 1080p, 4K, HDRip, BluRay)
    format_match = re.search(r'(\d{3,4}p|4K|2160p|HDRip|BluRay|HQ HDRip)', filename, re.IGNORECASE)
    video_format = format_match.group(1) if format_match else None

    # Extract encoding type (x264, x265, HEVC, AVC, SDR, HDR10, HDR10+)
    encoding_match = re.search(r'(x264|x265|HEVC|AVC|SDR|HDR10\+?)', filename, re.IGNORECASE)
    encoding = encoding_match.group(1) if encoding_match else "x264"  # Default to x264

    # Extract language (Handles both bracketed and inline languages)
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

    # Extract audio formats
    audio_matches = re.findall(r'(DD\+5\.1|DD\+|DD plus|AAC|Dolby Atmos|2\.1|\d{3,4}Kbps)', filename, re.IGNORECASE)
    audio_format = ' & '.join(set(audio_matches)) if audio_matches else None

    # Extract season and episode numbers
    season_match = re.search(r'[Ss](\d{1,2})[\sEeP]*(\d{1,2})', filename, re.IGNORECASE)
    season = season_match.group(1) if season_match else None
    episode = season_match.group(2) if season_match else None

    # Extract file size in GB or MB
    try:
        file_size_bytes = os.path.getsize(file_path)
        file_size = round(file_size_bytes / (1024 * 1024 * 1024), 2)
        file_size = f"{file_size} GB" if file_size >= 1 else f"{round(file_size_bytes / (1024 * 1024), 2)} MB"
    except FileNotFoundError:
        file_size = None  # Handle missing files gracefully

    # Remove extra details inside square brackets
    filename = re.sub(r'\[.*?\]', '', filename).strip()

    # Extract year only if it's a 4-digit number (not mistaken for season/episode)
    match = re.search(r'(.+?) \((\d{4})\)', filename)
    if match:
        movie_name = match.group(1).strip()
        year = match.group(2)
    else:
        # Try to extract year from a common position if not inside parentheses
        year_match = re.search(r'\b(19\d{2}|20\d{2})\b', filename)
        year = year_match.group(1) if year_match else None
        movie_name = filename.split(year)[0].strip() if year else filename.strip()

    return movie_name, year, video_format, encoding, languages, audio_format, season, episode, file_size


def process_movie_files(directory):
    """Processes MKV files from the given directory and categorizes them into unique and duplicate movies based on movie name."""
    movie_dict = defaultdict(list)
    duplicates = []
    movie_data = []

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".mkv"):
                file_path = os.path.join(root, file)
                movie_name, year, video_format, encoding, languages, audio_format, season, episode, file_size = extract_movie_details(
                    file, file_path)
                if movie_name:
                    # Normalize movie name for comparison
                    normalized_name = movie_name.strip().lower()
                    if normalized_name in movie_dict:
                        # Add to duplicates if movie name already exists
                        duplicates.append((movie_name, year, video_format, encoding, languages, audio_format, season,
                                           episode, file_size))
                    else:
                        # Add to unique movies
                        movie_dict[normalized_name] = (
                        movie_name, year, video_format, encoding, languages, audio_format, season, episode, file_size)
                    movie_data.append(
                        (movie_name, year, video_format, encoding, languages, audio_format, season, episode, file_size))

    print(f"Number of duplicates found: {len(duplicates)}")
    return movie_data, duplicates


def save_to_excel(movie_data, duplicates):
    """Saves movie data into Excel sheets while preserving existing records."""
    movies_path = os.path.abspath("Movies.xlsx")
    duplicates_path = os.path.abspath("Duplicates.xlsx")

    # Load existing data if Excel files exist
    if os.path.exists(movies_path):
        df_existing = pd.read_excel(movies_path)
    else:
        df_existing = pd.DataFrame()

    df_new = pd.DataFrame(movie_data,
                          columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages', 'Audio Format', 'Season',
                                   'Episode', 'Size'])
    df_movies = pd.concat([df_existing, df_new]).drop_duplicates().reset_index(drop=True)
    df_movies.to_excel(movies_path, index=False)
    print(f"Movies.xlsx updated at: {movies_path}")

    if duplicates:
        if os.path.exists(duplicates_path):
            df_existing_dup = pd.read_excel(duplicates_path)
            # Ensure columns match
            if not set(df_existing_dup.columns) == set(df_new.columns):
                df_existing_dup = pd.DataFrame(columns=df_new.columns)
        else:
            df_existing_dup = pd.DataFrame()

        df_duplicates_new = pd.DataFrame(duplicates, columns=['Movie Name', 'Year', 'Format', 'Encoding', 'Languages',
                                                              'Audio Format', 'Season', 'Episode', 'Size'])
        df_duplicates = pd.concat([df_existing_dup, df_duplicates_new]).drop_duplicates().reset_index(drop=True)
        df_duplicates.to_excel(duplicates_path, index=False)
        print(f"Duplicates.xlsx updated at: {duplicates_path}")
    else:
        print("No duplicates found.")


if __name__ == "__main__":
    directory = r"E:\Movies"
    movie_data, duplicates = process_movie_files(directory)
    save_to_excel(movie_data, duplicates)
    print("Processing complete. Check Movies.xlsx and Duplicates.xlsx.")