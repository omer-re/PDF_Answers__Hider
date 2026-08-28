# PDF Answers Hider
When your PDF past exams has the answers marked and highlighted - this tool removes it

What does this app do?
Take a look
![Usage example](https://github.com/omer-re/PDF_Answers__Hider/blob/master/what%20I%20do.jpg)


## PDF_Answers_Hider 2.0 - main updates:
- New and improved scraping logic that keeps tables, equations and sketches [thank you, Boaz Shebo Berger for your great contribution]
- Batch action- choose multiple PDF files at once.
- Minor changes in GUI and buttons' descriptions.


## Instructions
1. download the .py file and run it on your python interperter OR just download the EXE file which made of the same script and can be found here:
   [PDF_Answers_Hider_2.0.exe](https://github.com/Tombombadil7/PDF_Answers__Hider/releases/latest)
if your windows warn you about it, click "more info" and then allow it to run.
3. choose a pdf file with answers marked on it, then choose where to locate the "clean" pdf and let it run.

## MacOs / Linux users:
1. The exe file won't work for you.
2. Make sure you have all the dependancies installed: `pip3 install -r requirements.txt`
3. Download the whole project, then run the the .py file wtih `python3 remover.py `.
4. That should work.

## NOTES
1. Adobe Reader doesn't like some files, I am not sure why. If it happens to you - try to open the clean file using Chrome, FireFox, Foxit etc.
2. It isn't an AI, ML, CV fancy script, it basically removes a layer of the pdf of shapes, leaving the text as it is.
If there are other images, sketches or some other objects on this layer- they might be removed as well, take it into considerations.

## ⭐️ Special Thanks

All the fundamental credit goes to **[omer-re](https://github.com/omer-re)**. This codebase was built as an enhancement and continuation of their fantastic original work. 

You can find the amazing source project right here: 
[Original Repository by omer-re](https://github.com/omer-re/PDF_Answers__Hider).
