#!/usr/bin/env python
# -*- coding: utf-8 -*-
# update by le @ 2017.6.10 for qxq

import os
from openpyxl import Workbook, load_workbook
import argparse
from argparse import RawTextHelpFormatter

old_label = u""
new_label = u""


def get_unique_key_seqs(Sheetname, Keys):
	# get all column value as row's unique key
	K1 = []
	for l in Sheetname.rows:
		keys = [Keys]
		if "+" in Keys:
			keys = Keys.split("+")
		k1 = ""
		for k in keys:
			i = ord(k) - 65  # change to number index
			if k1:
				k1 += "^^"
			k1 += unicode(l[i].value)
		K1.append(k1)
	return K1


def get_num_key(Keys):
	# turn A+B+C to [0,1,2]
	keys = [Keys]
	if "+" in Keys:
		keys = Keys.split("+")
	Key = [ord(k) - 65 for k in keys]  # change to number index
	return Key


def mark_diff(Workbook, Sheet, Indexs, MarkLabel, Basefilename):
	c = Sheet.max_column
	
	Sheet.cell(row=1, column=c + 1).value = u"Changes"
	for i in Indexs:
		Sheet.cell(row=i + 1, column=c + 1).value = MarkLabel
	Workbook.save("mark-%s" % Basefilename.decode('utf-8'))


def single_workbook(Sheet, Indexs, Key, MarkLabel, Basefilename):
	# store data 2 single workbook not all other data
	wb = Workbook()
	ws = wb.active
	ws.title = Sheet.title
	for i in Indexs:
		for cell in Sheet.iter_rows(min_row=i + 1, max_row=i + 1):
			line = []
			for c in cell:
				# cell 2 cell copy data, maybe too slow
				line.append(c.value)
			ws.append(line)
	wb.save("single-%s-%s" % (MarkLabel, Basefilename.decode('utf-8')))


def single_unique_workbook(SheetName, OnlyData, MarkLabel, Basefilename):
	wb = Workbook()
	ws = wb.active
	ws.title = SheetName
	
	for d in OnlyData:
		line = d.split("^^")
		ws.append(line)
	wb.save("single-unique-%s-%s" % (MarkLabel, Basefilename.decode("utf-8")))


def diff(Old_File, New_File, Sheet_Index_From_Old, Sheet_Index_From_New, Key_From_Old, Key_From_New, Single_File,
         Only_Unique_Key):
	w1 = load_workbook(filename=Old_File)
	s1 = w1.worksheets[Sheet_Index_From_Old]
	K1 = get_unique_key_seqs(s1, Key_From_Old)
	
	w2 = load_workbook(filename=New_File)  # data_only=True
	s2 = w2.worksheets[Sheet_Index_From_New]
	K2 = get_unique_key_seqs(s2, Key_From_New)
	
	only1 = set(K1) - set(K2)
	only2 = set(K2) - set(K1)
	# both = (set(K1) & set(K2)) - set(K1[0:1]) - set(K2[0:1])  # 去掉标题
	
	# get all index form K1 seq not just first one
	index1 = [[i for i, x in enumerate(K1) if x == o] for o in only1]
	index1 = [j for i in index1 for j in i]
	
	index2 = [[i for i, x in enumerate(K2) if x == o] for o in only2]
	index2 = [j for i in index2 for j in i]
	
	# ## temp remove both for perf
	# both_index1 = [[i for i, x in enumerate(K1) if x == b] for b in both]
	# both_index1 = [j for i in both_index1 for j in i]
	# # same data have different index in 2 files
	# both_index2 = [[i for i, x in enumerate(K2) if x == b] for b in both]
	# both_index2 = [j for i in both_index2 for j in i]
	
	if not Single_File:
		mark_diff(w1, s1, index1, old_label, Old_File.name)
		mark_diff(w2, s2, index2, new_label, New_File.name)
	else:
		if not Only_Unique_Key:
			single_workbook(s1, index1, Key_From_Old, old_label, Old_File.name)
			# single_workbook(s1, both_index1, Key_From_Old, "both", Old_File)
			single_workbook(s2, index2, Key_From_New, new_label, New_File.name)
		# single_workbook(s2, both_index2, Key_From_New, u"both", New_File)
		else:
			single_unique_workbook(s1.title, only1, old_label, Old_File.name)
			single_unique_workbook(s2.title, only2, new_label, New_File.name)
	print "== Done"


if __name__ == '__main__':
	parser = argparse.ArgumentParser(description='''
    Compare 2 Excel in line level by set columns as unique KEY.
    Output0: Mark a user defined label to a new last column in new copied excel files named mark-*.
    Output1: Different lines to single files from them.
    Output2: Different and Unique KEY to single files from them.
    ''', formatter_class=RawTextHelpFormatter)
	
	parser.add_argument("-1", dest="Old_File", required=True, type=file, help="older excel to compare")
	parser.add_argument("-2", dest="New_File", required=True, type=file, help="newer excel to compare")
	parser.add_argument("-s1", dest="Sheet_Index_From_Old", type=int, default=0, choices=[0, 1, 2, 3, 4, 5, 6],
	                    help="sheet index need to compare from older excel, default is 0")
	parser.add_argument("-s2", dest="Sheet_Index_From_New", type=int, default=0, choices=[0, 1, 2, 3, 4, 5, 6],
	                    help="sheet index need to compare from newer excel, default is 0")
	parser.add_argument("-k1", dest="Key_From_Old", required=True,
	                    help="unique Key Column from older excel,Simple: A or A+C or B+D+G")
	parser.add_argument("-k2", dest="Key_From_New", required=True,
	                    help="unique Key Column from newer excel,Simple: A or B+C or D+A+F")
	parser.add_argument("-l1", dest="Mark_Old_Label", default=u"old",
	                    help="Special a word to mark different line in older file as Label, default is \"old\"")
	parser.add_argument("-l2", dest="Mark_New_Label", default=u"新增",
	                    help="Special a word to mark different line in newer file as Label, default is \"新增\"")
	parser.add_argument("-s", dest="Single_File", action="store_true", default=False,
	                    help="output different lines to 2 single excel file")
	parser.add_argument("-u", dest="Only_Unique_Key", action="store_true", default=False,
	                    help="only output unique different key column data to 2 single excel file, use only with -s option")
	parser.add_argument('-v', action='version', version='%(prog)s 1.0')
	args = parser.parse_args()
	
	old_label = args.Mark_Old_Label
	new_label = args.Mark_New_Label
	
	diff(args.Old_File, args.New_File, args.Sheet_Index_From_Old, args.Sheet_Index_From_New, args.Key_From_Old,
	     args.Key_From_New, args.Single_File, args.Only_Unique_Key)
