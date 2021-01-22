# xlxs_Traverser
A python script that I wrote to traverse a bunch of xlxs files (a type of file used by Excel).

---

### Background story
For your information, This program was coded during the time I was attending a mathematical modelling contest, in which i needed to process hundreds of .xlxs files. I nearly freaked out when I dicovered that my teammate had been reading and documenting these files manually. So i wrote this script in about half an hour and processed the files for him.


## About the program
Here's how the program works:
1. It checks all the xlxs files in the given folder(You should change the path of the folder into the one of your own)
2. It repeatedly does these tasks： 
- opens one of the files
- checks the title ( check if it matches the given key works)
- if it matches the words, the program will copy the data in certain blocks into a new xlxs file
- closed the readed file, save the writen file

## Before you run this program
+ Make sure you've installed the openpyxl module.
