This is a program that extracts the securities table from the https://www.luxse.com/issuer/DeutscheBank/23571 site using selenium and chromedriver. It can easily be adapted to scrape other tables from luxse. One thing to note is since there isn't a way for the program to figure out when the last page has already been reached,  one has to monitor when the program gets to the last page of the table, at which point one has to manually stop the program with Control(^) + C. Then, one can use the parse.ipynb file to check for the duplicates in the dataFrame and remove any redundant rows that were added more than once at the end of the table.
