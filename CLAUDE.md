# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the scripts

```bash
# Generate the standard-styled PPTX
python create_ppt.py

# Generate the ENTAAR-branded PPTX
python create_ppt_entaar.py
```

Both scripts write output files directly to `/home/user/personal/` and print the saved path on success. Required packages: `python-pptx` and `Pillow`.

## Architecture

This repository contains two standalone Python scripts that programmatically generate PowerPoint presentations (`.pptx`) reporting on the IT/DX organizational structures of Japan's five largest general contractors (ゼネコン5社: 鹿島建設、大林組、清水建設、大成建設、竹中工務店).

**`create_ppt.py`** — Standard report style  
Uses a dark-blue/orange corporate color palette. Builds slides sequentially as a flat script: helper functions at the top (`add_rect`, `add_textbox`, `add_text_in_shape`, `slide_header`, `footer`), then each slide's drawing code inline. Slide content is hardcoded directly in the drawing calls.

**`create_ppt_entaar.py`** — ENTAAR company branding  
Adds `Pillow` to generate a dot-grid background image and a chevron watermark programmatically, then embeds them as pictures behind the shapes. Company data is extracted into a `companies` list of dicts and iterated to draw slides 3–7, instead of duplicating code per company. Helper functions follow the same pattern but use slightly different names (`tb` instead of `add_textbox`, `shape_text` for multi-paragraph shapes).

**`ゼネコン5社_IT_DX組織構成.md`** — Source reference document  
Markdown file containing the research data (org structures, notes, source URLs) that both scripts draw their content from. When updating slide content, start here.
