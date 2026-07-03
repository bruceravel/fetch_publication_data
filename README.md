# fetch_publication_data
Scrape data about NSLS-II beamline publications from the publications website

This repo has two tools that do the same thing:

1. A python script that scrapes data from the NSLS-II publications
   webpages and writes the results to a spreadsheet
   
2. A jupyter notebook that does the same scraping and stores to a
   pandas dataframe.  The notebook then sorts the results in
   interesting ways.

This works by noticing that the publication web page has (at the time
of this writing, 3 July, 2026) one and only one table.  It is the one
on the upper right labeled "Publication Summary".

Using the Pandas `read_html` function, the contents of the summary
table is imported as a list of strings, then parsed.

This is then repackaged as an excel spreadsheet or a pandas dataframe.

Super simple, but helpful for basic research about beamline
productivity.

BMM does OK for a beamline with only 50% GU time!


## Dependencies

To run these notebooks, you will need to install some things.  venv is
an easy way to do so.

In this folder, do

    python3 -m venv ./my-venv
        
Then install the dependencies locally by doing:

    ./my-venv/bin/pip install unidecode jupyter pandas openpyxl matplotlib lxml

After installing the dependencies into the virtual environment, try
something like this:

    ./my-venv/bin/jupyter notebook
