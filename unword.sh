#!/bin/sh

#====================================
# unword.sh
# by Jake Kara
# jake@jakekara.com
#
# usage: unword.sh wordfile[s]
#
# Microsoft  word files can be turned
# into .zip files and unzipped. If 
# there are image files, they'll be
# in the word/media/ subfolder
#
# It's easy to do manually, but it's
# even easier to use this script.
#
#==================================== 


# check if this is a .doc or .docx
# return 1 if it is. 
function isWordFile()
{
	for file in "$@"
	do
		extension="${file##*.}"
		#echo extension is $extension
		if [[ "$extension" == "doc" ]]
		then
			#echo Converting $file to docx
			#textutil -convert docx "$file" 
			# textutil not a solution.
			# Doesn't preserve images
			return 2
			#return -1 #No support for .doc files yet :(
		elif [[ "$extension" == "docx" ]]
		then 
			return 1
		fi
		return -1
	done
	return -1
}

function usage()
{
	echo "usage: $0 word_doc(s) [options]
	
		OPTIONS:
	-d	Delete word doc afterwards!!
	-o 	Open word doc in textedit 
	-h	Show this help message"
}

# Defaults. Don't delete file. Don't open it.
delete="NO";
open="NO";

# Parse arguments
while getopts ":hdo" flag; do
  	case $flag in
    	h)
		usage
		exit 0
      		;;
	d)
		delete="YES"
		;;
	o)
		delete="NO"
		open="YES"
		;;
    	\?)
      		echo "Invalid option: -$OPTARG" >&2
		usage
		exit 1
      		;;
  	esac
done

for file in "$@"
do

	isWordFile "$file"
	word=$?
	if [[ "$word" != 1 ]]
	then
		#echo $file is not a word file.
		if [[ "$word" == 2 ]]
		then
			# do nothing with .docs
			# for now

			continue;
		else
			continue;
		fi
	fi	
        echo ------------------------

	echo "Extracting images from " $file

	mv "$file" "${file%.*}.zip";
	unzip "${file%.*}.zip" -d "${file%.*}";
	#echo cd to "${file%.*}";
	if [ -d "${file%.*}/word/media" ]
	then
		echo "Found images."
		mv "${file%.*}/word/media" "${file%.*} photos";
	else
		echo "No images found."
	fi

	#=============================
	# cleanup taks
	
        rm -rf "${file%.*}/"
	
	if [[ "$delete" == "YES" ]]
	then
		rm "${file%.*}.zip"   #deletes word doc

	else
		mv "${file%.*}.zip" "$file" #keeps word doc

	fi
	
	if [[ $open == "YES" ]]
	then
		open "$file" -a TextEdit 
	fi
done