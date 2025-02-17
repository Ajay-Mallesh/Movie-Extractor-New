**** Requirement & Functionality ****


Movie Organizer Script Documentation
Overview
This script extracts details from movie filenames and organizes them into an Excel database while identifying duplicates. It supports various file formats, encodings, audio types, and multiple languages.

- -> Features

- Extracts Movie Name, Year, Format,
- Encoding, Languages, Audio Format,
- Season, Episode, and File Size.
- Detects duplicates and stores them separately.
- Saves movie details in Movies.xlsx and duplicates in Duplicates.xlsx.
- Supports .mkv, .mp4, and .avi file types.

- -> File Parsing Rules
- Movie Name & Year Extracted from Movie Name (Year).
- If year is missing, it's inferred from numbers in the filename.
-> Video Format
- Extracted from resolutions like 1080p, 720p, 4K, HDRip, BluRay.
-> Encoding
- Identifies formats such as x264, x265, HEVC, HDR10+.
Languages
- Extracts languages from brackets "[Tamil + Telugu]" or inline mentions like "English Dubbed".
Audio Format
- Recognizes formats like DD+5.1, Dolby Atmos, AAC, 320Kbps.
-> TV Shows
- Detects Seasons (S01) and Episodes (E05).
  
- -> File Size
- Reads the actual file size (in GB/MB) using os.path.getsize().
** Workflow **
- Extract movie details from filenames.
- Check if movie exists in Movies.xlsx.
- If it exists → Add to Duplicates.xlsx.
- If new → Save in Movies.xlsx.
- Save updated data and print a summary.

- ** Usage Instructions **
- Place movie files inside a folder.
- Run the script and enter the folder path when prompted.
- Check Movies.xlsx for updated movie data.
- Check Duplicates.xlsx for duplicate entries.

- *** Example ***

Here’s how the script would extract details from the given filename:

Filename:
www.1TamilMV.bike - MAX (2025) Kannada HQ HDRip - 720p - x264 - (DD+5.1 - 192Kbps & AAC) - 1.2GB - ESub.mkv

Extracted Details:

Attribute	Extracted Value
Movie Name	MAX
Year	2025
Format	HQ HDRip - 720p
Encoding	x264
Languages	Kannada
Audio	DD+5.1 - 192Kbps & AAC
Size	1.2 GB
Subtitles	Yes (ESub)


- 
Filename:
www.1TamilMV.bike - Madurai Paiyanum Chennai Ponnum (2025) S01 EP (01-03) Tamil TRUE WEB-DL - 720p - AVC - AAC - 600MB - ESub.mkv

Extracted Details:
Attribute	Extracted Value
Show Name	Madurai Paiyanum Chennai Ponnum
Year	2025
Season	1
Episodes	01-03
Format	TRUE WEB-DL - 720p
Encoding	AVC
Languages	Tamil
Audio	AAC
Size	600 MB
Subtitles	Yes (ESub)

Notes:
The script correctly identifies this as a TV show (Season 1, Episodes 01-03).
It extracts video format (TRUE WEB-DL - 720p) and encoding (AVC).
Language (Tamil) and Audio (AAC) are also detected.
The file size (600MB) is extracted correctly.
Subtitles (ESub) are identified.
  
