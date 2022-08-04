#########################
#
# This scritpt uses pandoc (https://pandoc.org/) 
# for converting HTML to DOCX, See https://pandoc.org/demos.html
# for examples.
# Make sure you download and install pandoc 
# for the appropriate platform from: https://pandoc.org/installing.html
# 
#
#########################

import glob
import os
import subprocess

# cd to the directory where you have HTML files

os.chdir("/Users/gkns/Desktop/dir_listing_download/")

for root, dirs, files in os.walk(".", topdown=True):
    for file in files:
    	if ".html" in file:
    		input_html_file_path = os.path.join(root, file)
    		output_doc_file_path = os.path.splitext(input_html_file_path)[0] + ".docx"
    		print ("Input: " + input_html_file_path)
    		print ("Output: " + output_doc_file_path)

    		args = ['pandoc', '-s', input_html_file_path, '-o', output_doc_file_path]

    		# check_call makes sure that if any error occurs,
    		# the script will stop there.
    		subprocess.check_call(args)
    		