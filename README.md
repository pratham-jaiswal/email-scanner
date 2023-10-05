# Email Scanner
This is a Python application for scanning your email inbox and searching for specific terms in the email body and attachments, of emails received on that day. The application uses a graphical user interface built with the tkinter library for user interaction. It is converted into a [standalone executable file](https://github.com/pratham-jaiswal/email-scanner/releases/tag/Latest) for ease of use and distribution.

> ***Note: This is a windows only application***

> ***Note: This only supports gmail and gmail related work accounts***

## Use Case
- To let user look up whether they are mentioned in any email or its attachments, for example in placement related emails for students.
- To let students look up important emails, for example keeping track of placement related emails for students.

## Features
- Look up for emails, received on that day, with custom terms in body or attachments.
- Delete any entry of a tracked email.
- Mark the entry of email as "Checked" which can mean noted or completed.
- The entries get saved and need not be scanned again.

## Installing and Using the App
- Download the *EmailScannerV1.exe* from [here](https://github.com/pratham-jaiswal/email-scanner/releases/tag/Latest).
- Double click on it to install *Email Scanner*.
- Open the *Email Scanner* app isntalled.
- Click on *Configure* -> *Edit Search Terms*, then *search_terms.json* will open in a notepad, and then add your search terms. For example,
    ```json
    {
        "search_terms": ["Pratham", "name@example.com", "Placement"]
    }
    ```
- Now save the *json* file and you can close the notepad.
- Now get back to the *Email Scanner* app, enter your email password, and then click on *Start Scan*.
- Once the scan is complete you will see popup window with success message.

## Get Started with Code
- Make sure you have Python installed
- Clone the repository
    ```sh
    git clone https://github.com/pratham-jaiswal/email-scanner.git
    ```
- Install the dependencies
    ```sh
    pip install tkinter imaplib pytz pandas PyPDF2 docx pptx
    ```
- Run the *main.py*
    ```sh
    python main.py
    ```
- The *Email Scanner* app will open.
- Go to *Configure* -> *Edit Search Terms*, then *search_terms.json* will open in a notepad, and then add your search terms. For example,
    ```json
    {
        "search_terms": ["Pratham", "name@example.com", "Placement"]
    }
    ```
- Now save the *json* file and you can close the notepad.
- Now get back to the *Email Scanner* app, enter your email password, and then click on *Start Scan*.
- Once the scan is complete you will see popup window with success message.

## License
This project is licensed under the MIT License. Feel free to use and modify the code for your own purposes.