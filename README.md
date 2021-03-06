Bidscraper

This is a script I wrote while working as an Estimator/Salesman at a commercial construction subcontractor.

The subcontractor used isqft, cmdgroup, and dodgepipline to search for projects (leads) requiring the products they sold in order to esimate the cost and bid on the projects.

Bidscraper uses Selenium to log into the 3 FTP websites and navigate to the 6 different individual product filters and returns an excel spreadsheet of all the projects and their corresponding products (skylight, operable partition, louver, etc.).

The script then organizes all the spreadsheets into one, deletes duplicate listings, and marks the projects with flags that identify each potential product that could be needed for each project, because it is likely for a project to require multiply products that they sell. 

I setup a computer in the office to run this script every morning just before 8:00am and then have the company's filemaker database import the organized spreadsheet into a leads table with links to every project posting on all the ftp sites.  


