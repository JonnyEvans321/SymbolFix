# SymbolFix
An Excel Add-in that uses VBA to auto-format engineering symbols. Requires Excel 2010 or later.

![ezgif com-optimize](https://user-images.githubusercontent.com/22935783/42184668-3bed0238-7e3e-11e8-950b-f9c9095c8f35.gif)

## Backstory
Back in 2015, my flatmate James Kilcran came to me with a problem. As a Mechanical Engineer, he would spend a lot of time creating detailed, well presented engineering models in Excel for clients and colleagues. A major hassle in this job was formatting all the engineering symbols. Something that Excel does not do very well.

So we set about creating an Excel Add-in that would auto-format engineering symbols, SymbolFix. Subscripting, capitalising, correcting typos, creating greek letters, math symbols... SymbolFix does them all automatically! 

After testing on a handful of Engineers, SymbolFix has built up a diverse library of symbols which it will auto-format. What's more, SymbolFix gives users the option to add extra symbols to the library.

## Installation
1.	Download the file ‘SymbolFixVBA.xlam’
2.	Open Excel
3.	Enable the SymbolFix add-in:
    1.	Click the windows button
    2.	Click ‘Excel options’
    3.	Go to the ‘add-ins’ tab
    4.	Click ‘Go’ next to ‘manage Excel add-ins’
    5.	Click ‘browse’ and then select the downloaded .xlam file
    6.	Tick the check-box next to ‘SymbolFix’
4. At this point you will need to close all open workbooks to enact the changes. 
5. You're done. Now every time you open/create an Excel workbook, you should see the SymbolFix group in the 'home' tab:
<img width="250" height="100" src="https://user-images.githubusercontent.com/22935783/42175493-7e2e495c-7e1d-11e8-9e5d-aa60395f2087.PNG">

## Usage
* To auto-format symbols, simply highlight the required cells, and press the 'Apply to selection' macro.

* To check out which symbols are in the SymbolFix library, check out the file 'SymbolLibrary.xlsx'.

* To insert Greek letter or math symbols, use the syntax below, then click 'Apply to selection' (note this can be found in 'SymbolFix Settings' which is in the bottom right hand side of the group).

<img width="600" height="400" src="https://user-images.githubusercontent.com/22935783/42175087-473c18ee-7e1c-11e8-8376-50cca78ebaf1.PNG">

* To clear any actions made by SymbolFix, highlight the required cells, and click 'Clear SymbolFix'.

* If you have a symbol that's not in the SymbolFix library, click 'Custom Symbol' and follow the instructions to add it in.

## Uninstall
1. Click the windows button
1. Click ‘Excel options’
  1.	Go to the ‘add-ins’ tab
  1.	Click ‘Go’ next to ‘manage Excel add-ins’
  1.	Click ‘browse’ and then select the downloaded .xlam file
  1.  Untick the check-box next to ‘SymbolFix’

