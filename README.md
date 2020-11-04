# <div align="center">ExcelBulk Processing</div>

This class is currently an on going process. I wrote this while working with realtime data collection using multiple inertial measurment units. 
## Data
The ```raw_data``` contains ```.xlsx``` files of 7 Subjects, with 11 columns. This is to test the class.

## How to run
Navigate the directory and run ```python3 main.py -i FOLDER_TO_READ -o OUTPUT_FOLDER```
or
Use the notebook provided, it's more convenient to work with.

## Future Work
- [x] Add an ```argument parser``` command line instructions
- [ ]This process is slower than expected, there instead of saving the files as new can be replaced with memory such as the ```keras.preprocessing``` methods.
- [ ]This is a part of a bigger Deep Learning pipeline, therefore in the future I will be converting this to a complete Deep Learning program, deployable on flask.
- [ ]Add more classes -> convert to a package -> publish on PyPi

