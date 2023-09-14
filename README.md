# fetch_publication_data
Scrape data about NSLS-II beamline publications from the publications website

This simple script grabs the publication web page for each beamline at
NSLS-II.  The publication web page has (at the time of this writing,
14 September, 2023) one and only one table.  It is the one on the
upper right labeled "Publication Summary".

Using the Pandas `read_html` function, the contents of the summary
table is imported as a list of strings, then parsed.

This is then repackaged as an excel spreadsheet.

Super simple, but helpful for basic research about beamline
productivity.

BMM does OK for a beamline with only 50% GU time!
