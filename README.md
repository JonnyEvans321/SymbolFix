# SymbolFix
An Excel Add-in that uses VBA to auto-format engineering symbols.

## Backstory
Back in 2015, my friend, colleague and flatmate James Kilcran came to me with a problem. As a Mechanical Engineer, he would spend a lot of time creating detailed, well presented engineering models in Excel for clients and colleagues. A major hassle in this job was formatting all the engineering symbols. Something that Excel does not do very well.

So we both set about creating an Excel Add-in that would auto-format engineering symbols, SymbolFix. Subscripting, capitalising, correcting typos, creating greek letters, math symbols... SymbolFix does them all automatically! 

After testing on a handful of Engineers at work, SymbolFix has built up a diverse library of symbols which it will auto-format. What's more, the add-in contains a user-form that allows users to add extra engineering symbols to the library. When a user adds a symbol, it automatically gets added to a list of 'pending symbols' (if consenting, via email), meaning that there is an ever growing library of useful symbols that SymbolFix can auto-format.

## Installation
1.	Download the file ‘SymbolFix.xlam’
1.	Open Excel
1.	Enable SymbolFix add-in:
  1.	Click the windows button
  1.	Click ‘Excel options’
  1.	Go to the ‘add-ins’ tab
  1.	Click ‘Go’ next to ‘manage Excel add-ins’
  1.	Click ‘browse’ and then select the downloaded .xlam file
  1.	Tick the check-box next to ‘SymbolFix’
1. At this point you will need to close all open workbooks to enact the changes. 
1. You're done. Now every time you open/create an Excel workbook, you should see the SymbolFix collection of macros on the 'home' tab:
<img width="250" height="100" src="https://user-images.githubusercontent.com/22935783/42175493-7e2e495c-7e1d-11e8-9e5d-aa60395f2087.PNG">

## Usage
* To auto-format symbols, simply highlight the cells required, and press the 'Apply to selection' macro.

* To insert Greek letter or math symbols, use the following syntax before using the 'Apply to selection' macro:

<img width="600" height="400" src="https://user-images.githubusercontent.com/22935783/42175087-473c18ee-7e1c-11e8-8376-50cca78ebaf1.PNG">

* To clear any actions made by SymbolFix on a selection of cells, click 'Clear SYmbolFix'.

* If you have a symbol that's not in the SymbolFix library, click 'Custom Symbol' and follow the instructions to add it in.

## Uninstall
1. Click the windows button
1. Click ‘Excel options’
  1.	Go to the ‘add-ins’ tab
  1.	Click ‘Go’ next to ‘manage Excel add-ins’
  1.	Click ‘browse’ and then select the downloaded .xlam file
  1.  Untick the check-box next to ‘SymbolFix’

