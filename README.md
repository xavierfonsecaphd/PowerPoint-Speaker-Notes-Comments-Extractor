# PowerPoint Speaker Notes Comments Extractor
This script will extract all speaker notes from your PowerPoint presentation. Speaker notes are the text that appears in the notes section below each slide in PowerPoint - this might be what usually people refer to as "comments."


It will create .txt and .json files with all the speaker notes found. You can then for e.g. print it and have it with you if you don't have two screens during the presentation.


# How TO:

    1. Create .venv python environment

        # Use current directory name
        py -m venv .venv

        # Activate environment (Windows)
        .\.venv\Scripts\activate 

        or

        # Activate environment (Linux)
        source .venv/bin/activate

    2. Install requirements

        py -m pip install .\requirements.txt 

        or 

        py -m pip install pywin32 

    3. Run

        py .\Speaker_Notes_Extractor.py '.\YOUR PRESENTATION FILE.pptx'   