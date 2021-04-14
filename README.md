# JobCard

Use python code to print Job Card
---


# Set up
Install python and the [library docx](https://python-docx.readthedocs.io/en/latest/)


# Documentation

### fitst.py
When you first time uses this to create the Doc file, you need to run it first. It will create the doc file and store the job number in var.txt.

### regular.py
After you got a JobCard.docx file, you able to run this, it will recreate the Doc file and the jobnumber will increase 4.

### JobCard.docx
This is a word file also able to open by the LibreOffice in Linux.

### var.txt
This file stored the variable that the job number, which will increase by 4 each time you run regular.py.

### print.sh
A shell file that converts Docx to PDF by command 'unoconv' then sends the task to the printer.

### runthis.sh
Run print.sh 10 times
