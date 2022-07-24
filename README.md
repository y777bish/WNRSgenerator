# WNRSgenerator

This is a project which focuses on making text-cards pattern for different games, in .docx document.

## Details

In this project I focused on mimicking the pattern of cards used in original card game "We Are Not Really Strangers" in order to create simple generator in which you can generate your own ones.
Original site, where you can buy yourself [WNRS](https://www.werenotreallystrangers.com/) game (it's worth it).

## How it works?

The project uses docx library which allows you to format .docx documents. The feature of the program is that you can pass it a .txt file, with lines, wirtten line by line and it will format them as you want.
Because I wanted to mimick WNRS game cards style, the .docx document with each line in different cell, formatted in certain pattern (which looks cool in my opinion) will appear after the compilation. 

## How to use it?

To run the application firstly you need to download the project. Using Python console, type `“pip install python-docx”`, to install docx library.
After that, in code, you should change the source of your .txt file with lines you want to format.
Then you only need to compile the project, go grab some coffee, call your friends, maybe watch some Neon Genesis Evangelion in meantime... you have time for yourself right? ... At the end, in your project folder, .docx document file with formatted tables with cells will appear. Your only job now is to print it and crop them out.