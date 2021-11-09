README for Project: csInvControl

(python 3.9.7)

The csInvControl-library is designed for daily use in my inventory control position.

There are three sub-modules:
- csInvControl.csInvControl
- csInvControl.quarterlyReports
- csInvControl.masterLists

csInvControl:
- Generate cycle count spreadsheets.
- Generate total inventory

quarterlyReports:
- Generate a current inventory "snapshot"
- Compare all previously generated snapshots and compile into a spreadsheet

masterLists:
- Generate a reference spreadsheet to distrubute based on a master list.


To Install and Setup:
- mkdir dir
- cd path/to/dir
- git clone https://github.com/mattrrobins/csInvControl.git
- cd csInvControl

pipenv:
- pipenv install

(Ana)conda:
- conda env create -f environment.yml
- pip install -e .    # To fix the imports not included in env.yml 
