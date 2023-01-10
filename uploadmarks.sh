#!/bin/bash
if [ $# -eq 0 ]
  then
    echo "Run with name of excel file and subject to upload"
elif [ $# -eq 1 ]
  then
    python3 ./upload-marks.py $1
elif [ $# -eq 2 ]
  then
    python3 ./upload-marks.py $1 $2 credentials.txt config.txt
else 
    echo "Run with name of excel file and subject to upload"
fi
