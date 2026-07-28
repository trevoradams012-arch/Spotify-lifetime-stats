Reads every Streaming_History_Audio_*.json in a directory, drops plays shorter
than a threshold (20 seconds by default), and writes a workbook with one sheet
each for songs, artists, per-year totals, and an overall summary.

Usage

python3 spotify_stats.py [--min-seconds 20] [--indir .] [--out FILE.xlsx]
