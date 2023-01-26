# Smart-Cropping

This script will convert PDF to image files to a given location, 1 image for each page of the PDF.
Then it will iterate each image of that location and crop each side of the images based on reference background color and then
insert images into Excel.  Note, this was created specifically for my needs.  
I removed all the file locations i use from the script, so you will have to enter those in yourself.
Use what you want from it.


Note:  The way the Smart cropping works is by reading the background color, then iterates through the TOP/SIDES/BOTTOM pixels
each separately until it finds an RGB value significately different from the background color and sets that location
for the crop.
