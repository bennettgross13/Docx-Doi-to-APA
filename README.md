# Docx-Doi-to-APA
Searches and replaces DOI's in a docx with APA-Style Citations

This Python code searches a docx file for a DOI in curly braces ( e.g. "{10.3389/fncel.2015.0047}"). The program sends a cURL-request to the Crossref Content Negotiation, replaces the DOI in the docx with a reference and attaches an Bibliography in APA-Style at the end of the document.
The program uses the python-docx module.

Currently the program only supports APA format, but this could be added by changing the cURL-request and a few of the functions.

**NOTE**: Currently the program only works correctly when the formatting of the curly braces and the DOI are the same. 





## Prerequisites

- Python 3.10
- Docx module ( `pip install python-docx` ) 


## Note


This is my first public Github project and the program is poorly written, so I would be grateful for any suggestions to improve it :)


