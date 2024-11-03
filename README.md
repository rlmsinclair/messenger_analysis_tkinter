Export Facebook Messenger chats and analyze them with AI.

Instructions for Mac OS (Probably the same for Linux, might be slightly different for windows):

Make sure you are on chrome version 130 by going to chrome://version

Download main.py, chromedriver and requirements.txt.

You can click on the green "Code" in the top right to download a zip file containing these.

Extract the zip file.

Make sure you have Python 3.11 installed.

Run these commands in order.

cd your_directory_here

python3 -m venv venv

source venv/bin/activate

pip install -r requirements.txt

**If you don't have tkinter, you can install it for mac with homebrew (you will have to install homebrew first) using this:

brew install python-tk

python3 main.py
