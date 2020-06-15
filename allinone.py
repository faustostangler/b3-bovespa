Skip to content
Search or jump to…

Pull requests
Issues
Marketplace
Explore
 
@faustostangler 
Learn Git and GitHub without any code!
Using the Hello World guide, you’ll start a branch, write comments, and open a pull request.


faustostangler
/
b3-bovespa
1
00
 Code
 Issues 0
 Pull requests 0 Actions
 Projects 0
 Wiki
 Security 0
 Insights
 Settings
b3-bovespa
/
allinone.py
 

Spaces

4

No wrap
1
# -*- encoding: utf-8 -*-
2
# pip install --upgrade virtualenv google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread oauth2client
3
# https://github.com/faustostangler/b3-bovespa/edit/master/allinone.py
4
​
5
# selenium
6
from selenium import webdriver
7
from selenium.webdriver.support.ui import WebDriverWait
8
from selenium.webdriver.common.by import By
9
from selenium.webdriver.support import expected_conditions as EC
10
from selenium.webdriver.support.ui import Select
11
​
12
# datetime
13
from datetime import datetime
14
from datetime import timedelta
15
import time
16
​
17
# google
18
import pickle
19
import os.path
20
from googleapiclient.discovery import build
21
from google_auth_oauthlib.flow import InstalledAppFlow
22
from google.auth.transport.requests import Request
23
import gspread
24
from oauth2client.service_account import ServiceAccountCredentials
25
​
26
import re
27
​
28
​
29
# back-to-basics
30
def user_defined_variables():
31
    try:
32
        # browser path
33
        global base_azevedo
34
        global base_note
35
        global base_hbo
36
        global chrome
37
        global firefox
@faustostangler
Commit changes
Commit summary
Update allinone.py
Optional extended description
Add an optional extended description…
 Commit directly to the master branch.
 Create a new branch for this commit and start a pull request. Learn more about pull requests.
 
© 2020 GitHub, Inc.
Terms
Privacy
Security
Status
Help
Contact GitHub
Pricing
API
Training
Blog
About
