# gicr_docx
A super-simple script to convert gicr snippets in xlsx to docx.

Just what it says. [GICR](http://www.webcorpora.ru/) is a Russian Internet megacorpus (you can get free access by applying by mail), and the notebook allows you to convert its snippets to docx documents.

Dependencies:

    pip install python-docx
    pip install pandas
    pip install openpyxl
    
Supposedly should work fine both with version 1.0 and 2.0 (as long as the table has columns *left*, *result* and *right*). 

Place your snippets in the data folder and get them back in .docx format from docs folder.

Updates

Version 1.5: you can convert your tsvs now too! Just put them in the folder.
