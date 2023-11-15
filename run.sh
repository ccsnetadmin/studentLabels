#!/bin/sh

python3 "`dirname $0`"/labelGenerator.py $1
open ~/Desktop/Labels.pdf