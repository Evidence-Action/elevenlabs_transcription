# ============================================================
# master_run.py
# ============================================================
# Description: Master script for batch audio transcription.
#              Section 1: Configures all settings
#              Section 2: Loads the API key
#              Section 3: Runs the transcription pipeline
#
# Author:      Gariyasi Garg
# Created:     23 March 2026
# Last modified: 2 April 2026
#
# Usage:       Open this file, edit Section 1 if needed, then run:
#              python master_run.py
#
# Dependencies: requests, python-dotenv, python-docx
#
# Update log:
#
#
# ============================================================



# ============================================================
# master_run.py — Section 1: Configuration
# ============================================================
# This section defines ALL settings for the transcription job.
# You should only ever need to edit THIS section.
# Nothing below this section, and nothing in transcribe.py,
# should need to be changed for normal use.
# ============================================================

# --- Packages ---

# pip install python-dotenv
# pip install python-docx

import os  # Used for building file paths that work on any operating system
from dotenv import load_dotenv  # Used to read the API key from the .env file


# --- File paths ---

# Root path: the folder that holds Input/, Output/, Code/ and .env/
#
# By default this is worked out automatically: it is the folder that
# contains the Code/ folder this script lives in. If you cloned this
# repository and keep your audio inside it, you do NOT need to change
# this line.
#
# Change it only if your Input/ and Output/ folders live somewhere else,
# e.g. on Windows:
#   root_path = r"C:\Users\yourname\Documents\MLE Audio Audits_Pilot"
# or on Mac:
#   root_path = "/Users/yourname/Documents/MLE Audio Audits_Pilot"
root_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Input path
input_path = os.path.join(root_path, "Input")

# Output path
output_path = os.path.join(root_path, "Output")

# Code path
code_path = os.path.join(root_path, "Code")

# Path to the .env file that contains API key
env_file_path = os.path.join(root_path, ".env", ".env")



# --- Audio file settings ---

# The file extension to look for in the input folder
audio_file_extn = ".m4a"


# --- Folder handling settings ---

# Whether to search subfolders inside the input folder.
# True  = walk through every subfolder and transcribe all matching files
# False = only look at files sitting directly inside the input folder
# Set True when your input is organised into nested folders.
recursive_search = True

# Whether to recreate the input's subfolder structure inside the output folder.
# True  = an audio file at  Input/A/B/file.m4a  is saved to  Output/A/B/file.docx
# False = every transcript is saved flat, directly inside the output folder
# This only has an effect when recursive_search is True.
preserve_folder_structure = True


# --- ElevenLabs API settings ---

# Which transcription model to use (scribe_v2 is the latest)
elevenlabs_model_id = "scribe_v2"

# The language of your audio files (ISO 639-1 code)
# Full list: https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes
# "hi" = Hindi, "te" = Telugu, "en" = English, etc.
language_code = "hi"

# The API endpoint for speech-to-text transcription
elevenlabs_api_url = "https://api.elevenlabs.io/v1/speech-to-text"

# Whether to identify different speakers in the audio
# True = label who is speaking, False = just transcribe the text
enable_diarization = True

# How many speakers are in the audio (1 to 32)
# Set to None if you want the API to detect automatically
num_speakers = None

# The granularity of timestamps in the transcription
# "word" = timestamp per word, "character" = per character, "none" = no timestamps
timestamps_granularity = "word"

# --- Output settings ---

# The file extension for output files
output_file_extn = ".docx"

# What to do if an output file already exists in the Output folder:
# "replace" = overwrite it with a new transcription (uses credits again)
# "skip"    = leave the existing file and don't re-transcribe (saves credits)
output_if_exists = "skip"



# ============================================================
# master_run.py — Section 2: Load API Key
# ============================================================
# This section: loads the ElevenLabs API key from the .env file
# ============================================================

# --- Load the API key ---

# load_dotenv reads the .env file and makes its contents available
# as environment variables. IMPORTANT: Keep the key out of the code.
load_dotenv(env_file_path)

# os.getenv retrieves the value of elevenlabs_api_key from the .env file.
# If the key is missing or the .env file doesn't exist, this returns None.
api_key = os.getenv("elevenlabs_api_key")

# Check: did we get a key?
# If not, stop immediately and display error.
if not api_key:
    print("ERROR: No API key found.")
    print("Make sure your .env file exists at:", env_file_path)
    print("And that it contains: elevenlabs_api_key=your_key_here")
    exit()  # Stop the script, cannot continue without key

# If we get here, the key loaded successfully
print("API key loaded successfully.")



# ============================================================
# master_run.py — Section 3: Run the Transcription Pipeline
# ============================================================
# This section calls transcribe.py, passing it all the config
# it needs. transcribe.py handles: finding files, calling the
# API, and saving the results.
#
# All settings are passed as arguments so that transcribe.py
# never needs to be opened or edited.
# ============================================================

import sys  # Used to tell Python where to find our other scripts

# Add the code folder to Python's search path so it can find transcribe.py.
# Without this, Python only looks in the current working directory,
# which may not be the code folder if you run the script from elsewhere.
sys.path.insert(0, code_path)

from transcribe import run_transcription  # Import the main function

# Run the pipeline — pass all config as arguments
run_transcription(
    input_path=input_path,
    output_path=output_path,
    audio_file_extn=audio_file_extn,
    elevenlabs_model_id=elevenlabs_model_id,
    language_code=language_code,
    elevenlabs_api_url=elevenlabs_api_url,
    output_file_extn=output_file_extn,
    api_key=api_key,
    enable_diarization=enable_diarization,
    num_speakers=num_speakers,
    output_if_exists=output_if_exists,
    timestamps_granularity=timestamps_granularity,
    recursive_search=recursive_search,
    preserve_folder_structure=preserve_folder_structure
)