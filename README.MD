# Overview

This project involves the development of a Python-based tool that reads event data from an Excel sheet and automatically adds those events to a Google Calendar. The software demonstrates how to bridge local data with a cloud-based service to streamline event scheduling. By creating this automated workflow, I sought to enhance my skills as a software engineer, particularly in handling data extraction, processing automation, and integration with external APIs.

The purpose of writing this software was to learn how to automate the transfer of scheduling information from a structured data source (like Excel) into a widely used calendar platform. Through this, I aimed to streamline repetitive tasks, reduce manual data entry, and deepen my understanding of authentication and authorization flows when interacting with Google’s APIs.

[Software Demo Video](https://youtu.be/Jil-Y9l29QY)

# Development Environment

I developed the software using Python 3 on macOS. The primary tools included:

- Python 3.12
- Visual Studio Code as the code editor
- `openpyxl` library for reading Excel files
- `google-api-python-client` and `google-auth-httplib2` libraries for interacting with the Google Calendar API
- `python-dateutil` for parsing and manipulating dates and times

# Useful Websites

- [Google Calendar API Documentation](https://developers.google.com/calendar/api)
- [Python 'openpyxl' Documentation](https://openpyxl.readthedocs.io/)
- [Python 'dateutil' Documentation](https://dateutil.readthedocs.io/en/stable/)
- [Google Calendar API overview for Developers](https://developers.google.com/calendar/api/guides/overview#:~:text=The%20Google%20Calendar%20API%20is,the%20Google%20Calendar%20Web%20interface./)

# Future Work

- Add improved error handling to gracefully handle unexpected data formats in Excel.
- Implement user-defined settings for calendar selection and custom time zones.
- Enhance the UI to allow for manual overrides or confirmation before creating events.
- Add better NLP to better enhance its abilities to extract data 
- Possibly integrate Machine Learning and word tokenization 