#!/bin/bash
if [ $# -eq 0 ]
  then
    echo "Run with name of excel file and subject to verify"
elif [ $# -eq 1 ]
  then
    python3 ./verify-marks.py $1
elif [ $# -eq 2 ]
  then
    python3 ./verify-marks.py $1 $2 credentials.txt config.txt
else 
    echo "Run with name of excel file and subject to verify"
fi
