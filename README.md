# import-apple-files
Bulk download all files from iPhone/iPad DCIM folder to local disk under Widows

Work is based on [this script](https://github.com/dblume/list-photos-on-phone).

### Why was this needed?
To manually backup iPhone files.

## Improvements
- searching for My Computer / This Computer / etc for any localization;
- downloading files, skip existing files
- check file size according to metadata
- set files creation according to metadata - it helps sorting files that do not have EXIF date inside
 
### This is active project now :)
iPhone 6 that I use provide DCIM folder. 

## Getting Started

1) Connect iPhone to Windows computer and unlock it
2) Open iPhone DCIM folder in Explorer: Win+E -> This Computer -> iPhone's Name -> Internal Storage -> DCIM
3) For each of 100APPLE-like subfolder open it and wait for Explorer to load list of all files
4) Run the script from IDLE or an IDE for easier clipboard access to its output, otherwise, from the DOS command line you can redirect to a file like so:

    C:\Python27\python.exe import-apple-files.py my_folder

5) Explore all files from iPhone in my_folder  
