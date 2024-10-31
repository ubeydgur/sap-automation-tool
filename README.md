# SAP Automation Tool

**A task I received through Upwork**

A bot that retrieves user information from an `Excel` file on the SAP website using `Selenium`, logs in with this information, and then scrapes the Software Components and Support Package sections on the website and processes the data into an `Excel` file.

Note: I cannot share Excel files as they contain the client's user information such as login information.

- SAP Website: [sap.com](https://www.sap.com/index.html)

## Features
- The **login** function allows you to log in to the website and bypass cookies.
  
- With the **scraping** function, specified parts of the website are scraped and the scraped data is kept in a list.
  
- With the **read_excel** and **writing_excel** functions, the **input** `Excel` file is read and the engraved data is written to the **output** `Excel` file.

## Dependencies

- [Selenium](https://selenium-python.readthedocs.io/)

- [Pandas](https://pandas.pydata.org/)

## Usage

1. Install the required dependencies:

    ```bash
    pip install -r requirements.txt
    ```

2. Run the script:

    ```bash
    python main.py
    ```

3. Enter your search query when prompted.

4. View the results.
