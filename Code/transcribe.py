# ============================================================
# transcribe.py
# ============================================================
# Description: Transcription pipeline for batch audio files.
#              This file is called by master_run.py — you should
#              never need to open or edit this file.
#
#              What it does:
#              1. Finds all audio files in the input folder
#              2. Sends each one to ElevenLabs API for transcription
#              3. Saves each transcription as a .docx file
#              4. Skips any file that fails and logs the error
#              5. Prints a summary at the end
#
# Author:      Gariyasi Garg
# Created:     24 March 2026
# Last modified: 2 April 2026
#
# Update log:
# - 2 April: added timestamps
#
# Usage:       Do NOT run this file directly.
#              Run master_run.py instead — it calls this file.
# ============================================================


# --- Packages ---

import os        # For file path operations (joining paths, listing files)
import requests  # For sending HTTP requests to the ElevenLabs API
from docx import Document  # For creating Word .docx files


def run_transcription(
    input_path,
    output_path,
    audio_file_extn,
    elevenlabs_model_id,
    language_code,
    elevenlabs_api_url,
    output_file_extn,
    api_key,
    enable_diarization,
    num_speakers,
    output_if_exists,
    timestamps_granularity,
    recursive_search=False,
    preserve_folder_structure=False,
    audio_file_extns=None,
    keyterms_dir=None,
    only_files=None
):
    """
    Main function that runs the full transcription pipeline.

    This function receives ALL its settings from master_run.py
    so that nothing is hardcoded here. The parameters are:

    input_path          - Folder containing the audio files
    output_path         - Folder where .docx files will be saved
    audio_file_extn     - File extension to look for (e.g. ".m4a")
    elevenlabs_model_id - Which ElevenLabs model to use (e.g. "scribe_v2")
    language_code       - ISO 639-1 language code (e.g. "te" for Telugu)
    elevenlabs_api_url  - The API endpoint URL
    output_file_extn    - Output file extension (e.g. ".docx")
    api_key             - Your ElevenLabs API key (loaded from .env)
    enable_diarization  - True to identify speakers, False for plain text
    num_speakers        - Max number of speakers (1-32), or None for auto
    output_if_exists    - "replace" to overwrite, "skip" to keep existing
    recursive_search    - True to also scan subfolders of input_path,
                          False to only scan files directly inside it
    preserve_folder_structure - True to mirror the input's subfolder layout
                          inside the output folder, False for flat output
    audio_file_extns    - Optional list of extensions to accept (e.g.
                          [".mp4", ".mpeg", ".ogg"]). Overrides audio_file_extn
                          when given. Folders whose name starts with "_" are
                          always skipped (handy for keyterm/glossary folders).
    keyterms_dir        - Optional folder of per-file keyterm lists. For an
                          audio file "X.mp4" the engine looks for "X.txt" here
                          (one keyterm per line) and biases the transcription
                          toward those terms (ElevenLabs keyterm prompting).
    only_files          - Optional list of file names / base names; when given,
                          only those files are processed (handy for re-runs).
    """


    # ============================================================
    # transcribe.py — Section 1: Find Audio Files
    # ============================================================
    # Scan the input folder and build a list of audio files to
    # process. If the folder is missing or empty, stop early.
    # ============================================================

    # Check: does the input folder actually exist?
    if not os.path.exists(input_path):
        print("ERROR: Input folder not found at:", input_path)
        print("Please create it and add your audio files.")
        return  # Exit this function — nothing to do

    # Build the set of extensions to accept. You can pass a single
    # audio_file_extn (e.g. ".m4a") OR a list via audio_file_extns
    # (e.g. [".mp4", ".mpeg", ".ogg"]). str.endswith accepts a tuple.
    extns = tuple(
        e.lower() for e in (audio_file_extns if audio_file_extns else [audio_file_extn])
    )

    # Get a list of every audio file to transcribe. Each entry is stored
    # as a path RELATIVE to input_path (e.g. "AWC/MP_.../file.m4a"), so we
    # can recreate the same folder structure in the output later if asked.
    # We use .lower() so that .M4A, .m4a, and .M4a all get picked up —
    # Windows doesn't care about case in file extensions, so neither should we.
    audio_file_list = []

    if recursive_search:
        # os.walk visits input_path and every subfolder beneath it,
        # so files nested any number of levels deep are found.
        for current_dir, sub_dirs, files_here in os.walk(input_path):
            # Don't descend into helper folders such as _keyterms / _glossaries.
            sub_dirs[:] = [d for d in sub_dirs if not d.startswith("_")]
            for file_name in files_here:
                if file_name.lower().endswith(extns):
                    full_path = os.path.join(current_dir, file_name)
                    # Store the path relative to input_path, not the full path.
                    audio_file_list.append(os.path.relpath(full_path, input_path))
    else:
        # Only look at files sitting directly inside input_path (no subfolders).
        for file_name in os.listdir(input_path):
            if file_name.lower().endswith(extns):
                audio_file_list.append(file_name)

    # If only_files was given, keep just those (match on base name, ignore extn).
    if only_files:
        wanted = set(os.path.splitext(os.path.basename(f))[0] for f in only_files)
        audio_file_list = [
            f for f in audio_file_list
            if os.path.splitext(os.path.basename(f))[0] in wanted
        ]

    # Check: did we find any matching files?
    if len(audio_file_list) == 0:
        print("ERROR: No", "/".join(extns), "files found in:", input_path)
        print("Please add your audio files and try again.")
        return  # Exit — nothing to transcribe

    # Tell the user how many files we found
    print("Found", len(audio_file_list), "audio file(s) to transcribe.")
    print("Output mode:", output_if_exists, "(existing files will be",
          "overwritten)" if output_if_exists == "replace" else "skipped)")
    print("---")  # Visual separator in the console output


    # ============================================================
    # transcribe.py — Section 2: Transcribe and Save Each File
    # ============================================================
    # For each audio file in the list:
    #   Step A:  Build the full file path
    #   Step A2: Check if output already exists (skip or replace)
    #   Step B:  Send it to the ElevenLabs API
    #   Step C:  Check the response and extract transcription
    #   Step D:  Save the text as a .docx file
    #
    # If any file fails at any step, we log the error, skip that
    # file, and continue to the next one. The script does NOT stop.
    # ============================================================

    # --- Prepare the API request headers ---
    # These headers are the same for every file, so we define
    # them once outside the loop. The xi-api-key header is how
    # ElevenLabs authenticates your request.
    request_headers = {
        "xi-api-key": api_key
    }

    # --- Build the API request data payload ---
    # These parameters are also the same for every file.
    # We build the dictionary once and reuse it in the loop.
    api_request_data = {
        "model_id": elevenlabs_model_id,       # e.g. "scribe_v2"
        "language_code": language_code,         # e.g. "te" for Telugu
        "diarize": str(enable_diarization).lower(),  # API expects "true"/"false"
        "timestamps_granularity": timestamps_granularity #timestamps
    }

    # Only include num_speakers if the user specified a value.
    # If it's None, we leave it out and let the API detect automatically.
    if num_speakers is not None:
        api_request_data["num_speakers"] = num_speakers

    # --- Create the output folder if it doesn't exist ---
    # We check here rather than failing later when we try to save.
    if not os.path.exists(output_path):
        os.makedirs(output_path)
        print("Created output folder at:", output_path)

    # --- Tracking variables ---
    # We keep count of successes, failures, and skips so we can
    # print a summary at the end. We also store the names of
    # failed files so the user knows exactly which ones to check.
    success_count = 0
    failure_count = 0
    skipped_count = 0           # Tracks files skipped because output exists
    failed_file_names = []      # List to store names of files that failed


    # --- Main loop: process each audio file ---
    for file_index, audio_file_name in enumerate(audio_file_list):

        # Print progress so the user knows what's happening.
        # enumerate gives us file_index (0, 1, 2...) alongside
        # the file name, so we can show "Processing file 1 of 5".
        print(
            "Processing file",
            file_index + 1,       # +1 because enumerate starts at 0
            "of",
            len(audio_file_list),
            ":",
            audio_file_name
        )

        # --- try/except block ---
        # Everything inside "try" will attempt to run. If ANY step
        # fails, Python jumps to the "except" block, logs the error,
        # and continues to the next file instead of crashing.
        try:

            # --- Step A: Build the full path to the audio file ---
            # We combine the input folder path with the file name
            # to get the complete path Python needs to open the file.
            audio_file_path = os.path.join(input_path, audio_file_name)


            # --- Step A2: Work out the output path & check if it exists ---
            # audio_file_name is a path relative to input_path. We turn it
            # into the matching output path, swapping the extension
            # (file.m4a -> file.docx).
            if preserve_folder_structure:
                # Keep the same subfolder path the audio file had.
                # e.g. "AWC/MP_.../file.m4a" -> "AWC/MP_.../file.docx"
                output_relative_path = os.path.splitext(audio_file_name)[0] + output_file_extn
            else:
                # Flat output: drop any subfolders, use the file name only.
                base_name = os.path.basename(audio_file_name)
                output_relative_path = os.path.splitext(base_name)[0] + output_file_extn

            # The short name (for printing) and the full path on disk.
            output_file_name = os.path.basename(output_relative_path)
            output_file_path = os.path.join(output_path, output_relative_path)

            # Make sure the subfolder for this file exists before we save.
            # exist_ok=True means it's fine if the folder is already there.
            os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

            # If the output file exists and mode is "skip", move on.
            # "continue" jumps to the next file in the loop — the API
            # is never called, so no credits are used.
            if os.path.exists(output_file_path) and output_if_exists == "skip":
                print("  SKIPPED: output already exists")
                skipped_count += 1
                print("---")
                continue  # Jump to the next file


            # --- Step B: Send the file to the ElevenLabs API ---
            # We open the audio file in binary read mode ("rb").
            # "rb" means "read binary" — audio files are binary data,
            # not text, so we must open them this way.
            #
            # The "with" statement automatically closes the file when
            # done, even if something fails — this prevents files
            # from getting locked on your system.
            #
            # requests.post sends the HTTP request and waits for
            # ElevenLabs to respond with the transcription.
            #
            # Build this file's form fields: start from the shared params,
            # then add any per-file keyterms. ElevenLabs expects "keyterms"
            # repeated once per term, so we send a list of (name, value)
            # tuples rather than a plain dict.
            request_fields = list(api_request_data.items())
            file_keyterms = []
            if keyterms_dir:
                keyterms_file_path = os.path.join(
                    keyterms_dir,
                    os.path.splitext(os.path.basename(audio_file_name))[0] + ".txt"
                )
                if os.path.exists(keyterms_file_path):
                    with open(keyterms_file_path, encoding="utf-8") as keyterms_file:
                        file_keyterms = [line.strip() for line in keyterms_file if line.strip()]
                    request_fields += [("keyterms", term) for term in file_keyterms]
            if file_keyterms:
                print("  Using", len(file_keyterms), "keyterms to bias transcription")

            with open(audio_file_path, "rb") as audio_file:
                api_response = requests.post(
                    elevenlabs_api_url,            # The API endpoint URL
                    headers=request_headers,        # Contains our API key
                    files={
                        "file": audio_file          # The audio file to transcribe
                    },
                    data=request_fields             # Model, language, diarization, keyterms
                )


            # --- Step C: Check the API response ---
            # raise_for_status checks if the API returned an error
            # (like 401 Unauthorized, 429 Rate Limit, 500 Server Error).
            # If so, it throws an exception which our "except" block catches.
            api_response.raise_for_status()

            # The API returns JSON data. We parse it into a Python dictionary.
            response_data = api_response.json()


            # --- Step D: Save the transcription as a .docx file ---
            # output_file_name and output_file_path were already
            # built in Step A2 above.

            # Create a new Word document using python-docx.
            # Document() creates a blank .docx file in memory.
            doc = Document()

            # How we write the text depends on whether diarization is on.
            if enable_diarization:
                # -------------------------------------------------
                # DIARIZATION ON: format text with speaker labels
                # -------------------------------------------------
                # With diarization, the response contains a "words"
                # list where each word has a "speaker_id" field
                # (e.g. "speaker_0", "speaker_1").
                #
                # We group consecutive words by the same speaker
                # and write each group as a paragraph:
                # "Speaker 0: their words here"
                # "Speaker 1: their words here"
                # -------------------------------------------------
                words_list = response_data.get("words", [])

                # Check: did we get any words back?
                if len(words_list) == 0:
                    print("  WARNING: API returned empty transcription for", audio_file_name)
                    # Still save an empty doc so the user knows it was processed
                    doc.add_paragraph("[Empty transcription]")
                else:
                    # Track which speaker is currently talking so we
                    # know when the speaker changes.
                    current_speaker = None
                    current_text = ""
                    current_start_time = 0  # NEW — track when this speaker started

                    for word_info in words_list:
                        word_speaker = word_info.get("speaker_id", "unknown")
                        word_text = word_info.get("text", "")
                        word_start = word_info.get("start", 0)  # NEW — timestamp in seconds

                        if word_speaker != current_speaker:
                            if current_speaker is not None:
                                # Format timestamp as MM:SS
                                minutes = int(current_start_time // 60)
                                seconds = int(current_start_time % 60)
                                timestamp = f"[{minutes:02d}:{seconds:02d}]"

                                speaker_label = current_speaker.replace(
                                    "speaker_", "Speaker "
                                )
                                doc.add_paragraph(
                                    timestamp + " " + speaker_label + ": " + current_text.strip()
                                )

                            current_speaker = word_speaker
                            current_text = word_text
                            current_start_time = word_start  # NEW — record start time
                        else:
                            current_text += word_text

                    # Last speaker
                    if current_speaker is not None:
                        minutes = int(current_start_time // 60)
                        seconds = int(current_start_time % 60)
                        timestamp = f"[{minutes:02d}:{seconds:02d}]"

                        speaker_label = current_speaker.replace(
                            "speaker_", "Speaker "
                        )
                        doc.add_paragraph(
                            timestamp + " " + speaker_label + ": " + current_text.strip()
                        )

            else:
                # -------------------------------------------------
                # DIARIZATION OFF: save plain transcription text
                # -------------------------------------------------
                transcription_text = response_data.get("text", "")

                # Check: did we get any text back?
                if not transcription_text:
                    print("  WARNING: API returned empty transcription for", audio_file_name)
                    doc.add_paragraph("[Empty transcription]")
                else:
                    doc.add_paragraph(transcription_text)

            # Save the Word document to the output folder
            doc.save(output_file_path)

            # If we got here, everything worked for this file
            print("  Saved:", output_file_name)
            success_count += 1


        except Exception as error:
            # -------------------------------------------------
            # ERROR HANDLER
            # -------------------------------------------------
            # Something went wrong with this file. Log it and
            # move on to the next one. "Exception" catches any
            # type of error — network issues, file permission
            # problems, API errors, etc.
            # -------------------------------------------------
            print("  FAILED:", audio_file_name)
            print("  Error:", str(error))
            failure_count += 1
            failed_file_names.append(audio_file_name)

        # After try/except, the loop continues to the next file
        print("---")  # Visual separator between files


    # ============================================================
    # transcribe.py — Section 3: Summary
    # ============================================================
    # Print a final report so the user knows what happened.
    # ============================================================

    print("")  # Blank line for readability
    print("========== TRANSCRIPTION COMPLETE ==========")
    print("Total files found:  ", len(audio_file_list))
    print("Successfully saved: ", success_count)
    print("Skipped (existing): ", skipped_count)
    print("Failed:             ", failure_count)

    # If any files failed, list them so the user knows which
    # ones to check or retry.
    if failure_count > 0:
        print("")
        print("The following files failed:")
        for failed_name in failed_file_names:
            print("  -", failed_name)

    print("=============================================")