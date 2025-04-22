# Description

This script allows to convert .txt to .docx, primarily targeting GOST 7.32 (ГОСТ, государственный стандарт, national standard).
The main goal is to use traditional text editors, such as vim, to produce .docx files. Thus, allowing more control over the end result without weird shenanigans of WYSIWYG editors, often inconsistent and error-prone programs.

With style definition per each tag/macro, you can change document appearance in CSS-like manner. Each style is defined in .json.
You can also easily add your own macros that can check, fix or stylize .docx paragraphs - for an example look into GdocxHandler.py.

You can use the script as... a script.

Also, use it as a library! Just use `init_gostdocx(**kwargs)` to pass arguments
to the module (as if through command line) and then `process_txt(inpath, outpath)`
# Usage

Install dependencies:
```
pip install python-docx
pip install docxcompose
```

Create example .txt and .docx files:
```
python3 example.py
```

Process your .txt file:
```
python3 main.py -i YOUR_FILE.txt -o YOUR_OUTPUT.docx -s -se
```
