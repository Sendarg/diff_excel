

# **diff\_excel.py**


	usage: diff_excel.py [-h] -1 OLD_FILES [OLD_FILES ...] -2 NEW_FILES
                     [NEW_FILES ...] [-s1 {0,1,2,3,4,5,6}]
                     [-s2 {0,1,2,3,4,5,6}] -k1 KEY_FROM_OLD -k2 KEY_FROM_NEW
                     [-l1 MARK_OLD_LABEL] [-l2 MARK_NEW_LABEL] [-s] [-u] [-v]

    Compare 2 Excel groups in line level by set columns as unique KEY.
    Output0: Mark a user defined label to a new last column in new copied excel files named mark-*.
    Output1: Different lines to single files from them.
    Output2: Different and Unique KEY to single files from them.
    

	optional arguments:
	  -h, --help            show this help message and exit
	  -1 OLD_FILES [OLD_FILES ...]
	                        older excels to compare, many excels must special same sheet and Keys
	  -2 NEW_FILES [NEW_FILES ...]
	                        newer excels to compare, many excels must special same sheet and Keys
	  -s1 {0,1,2,3,4,5,6}   sheet index need to compare from older excel, default is 0
	  -s2 {0,1,2,3,4,5,6}   sheet index need to compare from newer excel, default is 0
	  -k1 KEY_FROM_OLD      unique Key Column from older excel,Simple: A or A+C or B+D+G
	  -k2 KEY_FROM_NEW      unique Key Column from newer excel,Simple: A or B+C or D+A+F
	  -l1 MARK_OLD_LABEL    Special a word to mark different line in older file as Label, default is "old"
	  -l2 MARK_NEW_LABEL    Special a word to mark different line in newer file as Label, default is "新增"
	  -s                    output different lines to 2 single excel file
	  -u                    only output unique different key column data to 2 single excel file, use only with -s option
	  -v                    show program's version number and exit

