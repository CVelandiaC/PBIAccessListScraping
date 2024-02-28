# PBIAccessListScraping
## Web Scrapping the List of People With Access to PBI Reports

**Author:** Cristian C. Velandia C.

**Description:** This notebook shows how to configure Selenium Chrome Driver and perform a simple web scarpping task to create an inventory in excel about people with access to a Power BI report, This was done this way given the fact that Power BI does not offer a simple way to extract those lists, not even for admins. This offers a simple and effective solution to thtat problematic. 

**Input:** csv file with the following columns with no headers
* company: The name of the company that belongs to the report
* Working Area: Power BI work area name, This could be extracted during the process but for simplicity while extracting the URLs can be added to the list 
* PBI Report Name: Report name, This could be extracted during the process but for simplicity while extracting the URLs can be added to the list 
* URL: The URL of the Direct Access section of the Power BI Report

**Output:** An Excel File that contains the list of people with access and the permissions with a name composed by ["id"_"working area" "PBI Name" "current date".xlsx]

**Main Libraries to be used:** 
- selenium                  4.16.0
- pandas                    2.1.1
- openpyxl                  3.1.2

**Python Version:** Python 3.9.18