# What is this?

This is a script/application that automates the creation of a document that the Head Of Staff at Stavox can use to take notes and have the links for the applicant ready at hand. The document is a markdown file.

This started as a node project, but I ran into some problems which ment I quickly had to convert to Python to get the script up and running in time.

## How?

This is currently a simple Python script that uses a chromedriver.
The script opens the page with the applications, and then one by one opens the applications and copies the applicant's name, link to application and steamid.

This is then inserted into the markdown file in a table format. To use the table in eg. Google Docs, open the markdown document in a markdown viewer and select the generated table into the document.

Example of the document opened in a markdown viewer:
<br>
![Document example](https://i.imgur.com/EqWZ8zF.png)

## Requirements

-   Python
-   Selenium
-   webdriver-manager
-   python-docx
