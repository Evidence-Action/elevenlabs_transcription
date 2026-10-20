# Audio transcription pipeline (ElevenLabs Scribe)

Batch-transcribes a folder of audio or video recordings into Word documents,
using the [ElevenLabs Speech-to-Text API](https://elevenlabs.io/speech-to-text).

Built for the MLE audio-audit pilot (field-audit and household-survey
recordings in Hindi and Telugu), but nothing in it is specific to that project —
point it at a folder of recordings in any supported language and it will work.

**What you get:** one `.docx` per recording, with speaker labels and timestamps:

```
[00:00] Speaker 0: नमस्ते, मैं आपसे कुछ सवाल पूछना चाहती हूँ...
[00:14] Speaker 1: हाँ, बिल्कुल पूछिए।
[00:21] Speaker 0: सबसे पहले, आपका नाम क्या है?
```

Files that fail (bad audio, rate limits, network drops) are logged and skipped —
the run keeps going and tells you at the end exactly which files to retry.

---

## Contents

1. [How it works](#1-how-it-works)
2. [Setup — do this once](#2-setup--do-this-once)
3. [Running a transcription](#3-running-a-transcription)
4. [All the settings](#4-all-the-settings)
5. [What the console shows you](#5-what-the-console-shows-you)
6. [Troubleshooting](#6-troubleshooting)
7. [Cost and credits](#7-cost-and-credits)
8. [Optional extras](#8-optional-extras)

---

## 1. How it works

There are only two files you need to care about, and you only ever **edit one of them**.

| File | What it is | Do you edit it? |
| --- | --- | --- |
| `Code/master_run.py` | **The control panel.** Holds every setting (which folder, which language, speaker labels on/off...), loads your API key, and starts the job. | **Yes** — edit Section 1 before a run. |
| `Code/transcribe.py` | **The engine.** Finds the audio files, calls the API, writes the Word documents, counts successes and failures. Takes all its settings from `master_run.py`. | **No** — never needs editing. |

The flow is deliberately boring:

```
Input/  ─────▶  master_run.py  ─────▶  transcribe.py  ─────▶  ElevenLabs API
 (your          (your settings       (finds files,            (does the
  audio)         + API key)           saves .docx)             transcribing)
                                            │
                                            ▼
                                        Output/  (one .docx per recording)
```

### Folder layout

```
mle-audio-transcription/
├── Code/
│   ├── master_run.py            ← open and run this one
│   └── transcribe.py            ← the engine (don't edit)
├── Input/                       ← put your recordings here
├── Output/                      ← transcripts appear here
├── .env/
│   └── .env                     ← your API key (you create this; never shared)
├── .env.example                 ← template for the file above
└── requirements.txt             ← the three Python packages needed
```

`Input/` and `Output/` are deliberately empty on GitHub. Recordings and transcripts
are participant data and must never be committed — `.gitignore` blocks them, along
with the `.env` folder that holds the API key.

---

## 2. Setup — do this once

### 2a. Get the code

Either download the ZIP from GitHub (green **Code** button → *Download ZIP*) or, if you
have git installed:

```bash
git clone https://github.com/<owner>/<repo>.git
cd <repo>
```

### 2b. Install Python

Download from [python.org/downloads](https://www.python.org/downloads/) and run the installer.

> **On Windows, tick "Add python.exe to PATH" during installation.** This is the single
> most common thing people miss, and nothing works without it.

Check it worked — open Command Prompt (Windows) or Terminal (Mac) and run:

```bash
python --version
pip --version
```

(On a Mac you may need `python3` and `pip3` everywhere this README says `python` and `pip`.)

### 2c. Install the three packages

From the project folder:

```bash
pip install -r requirements.txt
```

That installs `requests` (talks to the API), `python-dotenv` (reads the API key) and
`python-docx` (writes the Word files).

### 2d. Get an ElevenLabs API key

An API key is a long password that lets this script use your ElevenLabs account. Steps:

1. Log in at [elevenlabs.io](https://elevenlabs.io).
2. Bottom-left, click **Developers** → **API Keys** → **Create**.
3. Give it a descriptive name, e.g. `transcription-script`.
4. Set every permission to **No Access except Speech-to-Text**. The key then can't
   do anything else on the account, even if it leaks.
5. **Set a credit limit** on the key. This caps what a runaway script can spend.
6. **Copy the key immediately** — ElevenLabs shows it only once. It starts with `sk_`.

### 2e. Store the key (never paste it into the code)

The script reads the key from a file that is never committed to GitHub.

1. In the project root, create a folder called `.env`
2. Inside it, create a file called `.env` — so the path is `.env/.env`
3. Put one line in it:

```
elevenlabs_api_key=sk_paste_your_actual_key_here
```

No quotes, no spaces around the `=`, nothing else in the file. `.env.example` in the
project root is a copy-paste template.

> **Why a folder called `.env` containing a file called `.env`?** It's just the layout
> this project settled on, and the scripts look for it there. If you'd rather use a
> plain `.env` file in the project root, change this line in `master_run.py`:
> `env_file_path = os.path.join(root_path, ".env", ".env")` → `os.path.join(root_path, ".env")`

Your key is now set up. You never have to touch it again.

---

## 3. Running a transcription

1. **Put your recordings in `Input/`.** By default the script looks for `.m4a` files.
2. **Open `Code/master_run.py`** and check Section 1 — above all `language_code`
   (`"hi"` Hindi, `"te"` Telugu, `"en"` English).
3. **Run it:**

```bash
python Code/master_run.py
```

   Or open the file in VS Code and press the ▶ Run button.

4. **Wait.** Progress prints as it goes — roughly real time or faster per recording.
5. **Collect your transcripts from `Output/`,** one `.docx` per recording.

Safe to re-run: with `output_if_exists = "skip"` (the default), files already
transcribed are left alone and cost nothing. If a run dies halfway, just run it
again — it picks up where it stopped.

---

## 4. All the settings

Everything lives in **Section 1 of `master_run.py`**. The ones you'll actually change
between runs are marked ★.

| Setting | What it does | Default |
| --- | --- | --- |
| ★ `language_code` | Language of the audio, as an [ISO 639-1 code](https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes). `"hi"` Hindi, `"te"` Telugu, `"en"` English. Getting this wrong gives garbled output. | `"hi"` |
| ★ `output_if_exists` | `"skip"` keeps transcripts you already have (saves credits). `"replace"` re-transcribes and overwrites them. | `"skip"` |
| ★ `enable_diarization` | `True` labels who is speaking (`Speaker 0`, `Speaker 1`...). `False` gives one block of plain text. | `True` |
| ★ `num_speakers` | How many voices are on the recording (1–32). `None` lets the API work it out. Set it when you know — it improves the labels. | `None` |
| `audio_file_extn` | Which file type to look for in `Input/`. | `".m4a"` |
| `recursive_search` | `True` also searches subfolders of `Input/`. `False` only looks at files sitting directly in it. | `True` |
| `preserve_folder_structure` | `True` mirrors the input's subfolders in `Output/` (`Input/A/B/x.m4a` → `Output/A/B/x.docx`). `False` saves everything flat. Only matters when `recursive_search` is on. | `True` |
| `timestamps_granularity` | `"word"`, `"character"` or `"none"`. The `[MM:SS]` stamps in the output come from this. | `"word"` |
| `elevenlabs_model_id` | Which ElevenLabs model to use. | `"scribe_v2"` |
| `output_file_extn` | Output file type. | `".docx"` |
| `root_path` | The folder holding `Input/`, `Output/`, `Code/` and `.env/`. Works itself out from where the script sits, so normally leave it alone. | auto |

---

## 5. What the console shows you

```
API key loaded successfully.
Found 3 audio file(s) to transcribe.
Output mode: skip (existing files will be skipped)
---
Processing file 1 of 3 : interview_01.m4a
  Saved: interview_01.docx
---
Processing file 2 of 3 : interview_02.m4a
  SKIPPED: output already exists
---
Processing file 3 of 3 : interview_03.m4a
  FAILED: interview_03.m4a
  Error: 429 Too Many Requests
---

========== TRANSCRIPTION COMPLETE ==========
Total files found:   3
Successfully saved:  1
Skipped (existing):  1
Failed:              1

The following files failed:
  - interview_03.m4a
=============================================
```

The summary is the bit that matters: anything under "The following files failed"
needs another run.

---

## 6. Troubleshooting

| What you see | What it means and what to do |
| --- | --- |
| `ERROR: No API key found.` | The script can't find or read `.env/.env`. Check the path, check the line reads `elevenlabs_api_key=sk_...` with no quotes or spaces, and check the spelling — the name is case-sensitive. |
| `ERROR: Input folder not found` | There's no `Input/` folder where the script expects it. Create it, or fix `root_path`. |
| `ERROR: No .m4a files found` | `Input/` is empty, or your recordings have a different extension — set `audio_file_extn` to match (e.g. `".mp3"`). |
| `FAILED ... 401 Unauthorized` | The API key is wrong, expired, or lacks Speech-to-Text permission. Make a new one. |
| `FAILED ... 429 Too Many Requests` | You hit the rate limit or ran out of credits. Wait a few minutes and re-run — `"skip"` mode resumes where you stopped. |
| `FAILED ... 422 Unprocessable Entity` | That file is corrupt or in a format the API won't take. Try playing it locally to check. |
| `ModuleNotFoundError: No module named 'dotenv'` | Packages aren't installed, or VS Code is using a different Python. Run `pip install -r requirements.txt`; in VS Code, `Ctrl+Shift+P` → *Python: Select Interpreter*. |
| Backslash errors in a path | On Windows, put `r` before the quote: `r"C:\Users\..."`. |
| `WARNING: API returned empty transcription` | The API heard nothing — usually silence, a near-empty file, or the wrong `language_code`. |

---

## 7. Cost and credits

- ElevenLabs bills speech-to-text **per minute of audio**, so cost scales with
  recording length, not file count.
- The Pro plan includes roughly 300 hours of API transcription per month; past that,
  overage rates apply.
- **Set a credit limit on the API key** (step 2d) so a mistake can't drain the account.
- Keep `output_if_exists = "skip"` — it's the main thing standing between you and
  paying twice for the same recording.
- Watch usage at elevenlabs.io → **Developers** → **Analytics**.

---

## 8. Optional extras

`transcribe.py` can do three things `master_run.py` doesn't switch on. If you need any
of them, copy `master_run.py` to a new file (say `run_my_project.py`), change the
settings there, and pass the extra arguments to `run_transcription()` at the bottom.
The standard runner then stays simple for everyone else.

**Several file types at once** — pass a list instead of a single extension:

```python
audio_file_extns = [".mp4", ".mpeg", ".mpg", ".ogg", ".m4a", ".mp3", ".wav"]
```

...then add `audio_file_extns=audio_file_extns` to the `run_transcription(...)` call.
Video files are fine — the API reads the audio track and ignores the picture.

**Keyterm prompting** — for names, places and jargon the model keeps mishearing.
Make a folder of plain-text lists, one term per line, each named after the recording it
belongs to (`interview_04.mp4` → `_keyterms/interview_04.txt`), then pass
`keyterms_dir=os.path.join(input_path, "_keyterms")`. The transcription is biased toward
those spellings. Folders whose name starts with `_` are never scanned for audio, so the
lists can safely sit inside `Input/`.

**Re-run only certain files** — pass `only_files=["interview_04"]` to re-do a few
recordings without touching the rest. To take them from the command line instead:

```python
only_files = sys.argv[1:] if len(sys.argv) > 1 else None
```

```bash
python Code/run_my_project.py "interview_04" "interview_07"
```

---

## A note on what's *not* in this repository

Recordings, transcripts and the API key stay on your machine. The `.gitignore`
blocks audio, video, `.docx` files and the `.env` folder, so a routine
`git add .` can't publish participant data or your key by accident. Keep it that way.
