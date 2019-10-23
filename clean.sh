#!/bin/bash

cat .gitignore | while read line
do
    path_pattern=`echo "$line" | sed -e "s/\/$//g"`

    for path in `find . -name "$path_pattern"`
    do
        echo $path
        rm -rf $path
    done
done
