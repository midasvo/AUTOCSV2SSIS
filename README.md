# CSV2SSIS
Simple script that monitors a directory for new CSV files and calls the SSIS package.

## Steps

1. Registers an ObjectEvent (FileSystemWatcher) which monitors a specific directory
2. User drops one or more CSV files into the directory
3. The ObjectEvent of the FileSystemWatcher uses a timer to establish whether or not we have a record of all the added files. It does this by turning off the Timer at the start of the function and enabling it at the end. By this logic once the timer runs out we have a record of all the added files. Each ObjectEvent adds the name of the file to an Array.
4. When the Timer runs out, copy each file to the SSIS working directory and rename it to the name given in the SSIS package for the Flat File Source. This name is taken from the CSV file. 
5. When all files are copied and renamed we call the SSIS package using DTEXEC
6. Check if SSIS package executed succesfully, relay this information to the user
7. Remove all files from the SSIS working directory
8. Clear the Array
9. Done. Waits for new files to be dropped into the monitoring folder

