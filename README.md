

# **diff\_excel.py**


	usage: diff_excel.py [-h] -1 OLD_FILE -2 NEW_FILE [-s1 {0,1,2,3,4,5,6}]
	                     [-s2 {0,1,2,3,4,5,6}] -k1 KEY_FROM_OLD -k2 KEY_FROM_NEW
	                     [-s] [-u] [-v]
	
	    Compare 2 Excel in line level by set columns as unique KEY.
	    Output0: Mark a label (old, both, new) to a new last column in new copied excel files named mark-*.
	    Output1: Different lines to single files from them.
	    Output2: Different and Unique KEY to single files from them.
	
	
	optional arguments:
	  -h, --help           show this help message and exit
	  -1 OLD_FILE          older excel to compare
	  -2 NEW_FILE          newer excel to compare
	  -s1 {0,1,2,3,4,5,6}  sheet index need to compare from older excel
	  -s2 {0,1,2,3,4,5,6}  sheet index need to compare from newer excel
	  -k1 KEY_FROM_OLD     unique Key Column from older excel,Simple: A or A+C or A+B+D
	  -k2 KEY_FROM_NEW     unique Key Column from newer excel,Simple: A or A+C or B+D+F
	  -s                   output different lines to 3 single excel file
	  -u                   only output unique different key column data to 2 single excel file, use only with -s option
	  -v                   show program's version number and exit