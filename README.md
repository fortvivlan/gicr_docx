# gicr_docx
A super-simple script to convert gicr snippets in xlsx to docx.

Just what it says. [GICR](http://www.webcorpora.ru/) is a Russian Internet megacorpus (you can get free access by applying by mail), and the notebook allows you to convert its snippets to docx documents.

Dependencies:

    pip install python-docx
    pip install pandas
    pip install openpyxl
    
Supposedly should work fine both with version 1.0 and 2.0 (as long as the table has columns *left*, *result* and *right*). 
