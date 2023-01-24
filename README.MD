# Plano Basico's Web Scraping with Data Treatment 

## General Description
- This automation go to the desired page
- Get all data from it
- Saves as pandas dataframe
- Manipulate and process data
- Saves as an excel file

## Code detailed description (steps)
1. The code go to desired page (url)
2. It show just 250 itens per page, so is needed to set this on the table's page view settings
3. Filter the 'Serviço' field with 'TV' name
4. As there are 250 itens per page, one needs to know how many pages will need to be viewed to get all the data
5. Every page is a new dataframe, so we concat the old dataframe with the new until we see all the pages
6. After getting all data, we manipulate and process data with some custom rules
7. By the end, we save all as an excel file

## Requirements
* python
* excel
* all libraries (see libraries topic)

## How to run on your environment 
* install all requirements
* run with "python automation.py"

## Libraries used
    * selenium
    * webdriver_manager
    * pandas
    * pandas
    * lxml