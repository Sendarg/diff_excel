#!/usr/bin/env python
# -*- coding: utf-8 -*-
# version 1.0 update by le @ 2017.6.10 for qxq
# version 1.0 update by le @ 2017.6.18 for qxq, support group compare

import os, glob
from openpyxl import Workbook, load_workbook
import argparse
from argparse import RawTextHelpFormatter

old_label = u""
new_label = u""


def get_unique_key_seqs(Sheet, Keys):
	# get all column value as row's unique key
	K1 = []
	for l in Sheet.rows:
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
	print "== Use Key:\t[ %s ]\tof Sheet:\t[%s]" % (K1[0], Sheet.title)
	return K1


def get_num_key(Keys):
	# turn A+B+C to [0,1,2]
	keys = [Keys]
	if "+" in Keys:
		keys = Keys.split("+")
	Key = [ord(k) - 65 for k in keys]  # change to number index
	return Key


def mark_diff1(Workbook, Sheet, Indexs, MarkLabel, Basefilename):
	c = Sheet.max_column
	
	Sheet.cell(row=1, column=c + 1).value = u"Changes"
	for i in Indexs:
		Sheet.cell(row=i + 1, column=c + 1).value = MarkLabel
	
	name = Basefilename.decode('utf-8').split("/")[-1]
	print "== OK:\t[%s]\t[%s]" % (name, Sheet.title)
	Workbook.save("mark-%s" % name)


def single_workbook(Sheet, Indexs, Key, MarkLabel, Basefilename):
	# store data 2 single workbook not all other data
	wb = Workbook()
	ws = wb.active
	ws.title = Sheet.title
	for i in Indexs:
		for line in Sheet.iter_rows(min_row=i + 1, max_row=i + 1):
			ws.append([c.value for c in line])
	name = Basefilename.decode('utf-8').split("/")[-1]
	print "== OK:\t[%s]\t[%s]" % (name, Sheet.title)
	wb.save("single-%s-%s" % (MarkLabel, name))


def single_unique_workbook(SheetName, OnlyData, MarkLabel, Basefilename):
	wb = Workbook()
	ws = wb.active
	ws.title = SheetName
	
	for d in OnlyData:
		line = d.split("^^")
		ws.append(line)
	name = Basefilename.decode('utf-8').split("/")[-1]
	print "== OK:\t[%s]\t[%s]" % (name, SheetName)
	wb.save("single-unique-%s-%s" % (MarkLabel, name))


def get_sub_index_index(Big_2LevelList, Sub_1LevelList, GroupIndex):
	# get all index form K1s seq not just first one
	index_a = [[[i, j] for i, x in enumerate(Big_2LevelList) for j, y in enumerate(Big_2LevelList[i]) if y == o] for o
	           in Sub_1LevelList]
	index_b = [j for i in index_a for j in i]
	# group 2 single list
	## method1,lambda ?? what mean?
	index_c = [[i for i in index_b if i[GroupIndex] == u] for u in set(map(lambda x: x[GroupIndex], index_b))]
	# ## method2 , sort+group
	# index1b.sort(key=itemgetter(0))
	# index1c=[[i for i in g] for k,g in groupby(index1b,key=itemgetter(0))]
	return index_c


def get_sub_index_value(Big_2LevelList, Sub_1LevelList, GroupIndex):
	# get all index form K1s seq not just first one
	index_a = [[[i, o] for i, x in enumerate(Big_2LevelList) for j, y in enumerate(Big_2LevelList[i]) if y == o] for o
	           in Sub_1LevelList]
	index_b = [j for i in index_a for j in i]
	# group 2 single list
	## method1,lambda ?? what mean?
	index_c = [[i for i in index_b if i[GroupIndex] == u] for u in set(map(lambda x: x[GroupIndex], index_b))]
	# ## method2 , sort+group
	# from itertools import groupby
	# from operator import itemgetter
	# index1b.sort(key=itemgetter(0))
	# index1c=[[i for i in g] for k,g in groupby(index1b,key=itemgetter(0))]
	return index_c


def diffPlus(Old_Files, New_Files, Sheet_Index_From_Old, Sheet_Index_From_New, Key_From_Old, Key_From_New, Single_File,
             Only_Unique_Key):
	W1s = [load_workbook(filename=old) for old in Old_Files]
	S1s = [w1.worksheets[Sheet_Index_From_Old] for w1 in W1s]
	K1s = [get_unique_key_seqs(s1, Key_From_Old) for s1 in S1s]
	K1 = [j for i in K1s for j in i]
	
	W2s = [load_workbook(filename=new) for new in New_Files]
	S2s = [w2.worksheets[Sheet_Index_From_New] for w2 in W2s]
	K2s = [get_unique_key_seqs(s2, Key_From_New) for s2 in S2s]
	K2 = [j for i in K2s for j in i]
	
	only1 = set(K1) - set(K2)
	only2 = set(K2) - set(K1)
	# both = (set(K1s) & set(K2s)) - set(K1s[0:1]) - set(K2s[0:1])  # 去掉标题
	
	## only old
	index1 = get_sub_index_index(K1s, only1, 0)
	index1o = get_sub_index_value(K1s, only1, 0)
	
	## only new
	index2 = get_sub_index_index(K2s, only2, 0)
	index2o = get_sub_index_value(K2s, only2, 0)
	
	if not Single_File:
		## fucking here
		# clean
		[os.remove(f) for f in glob.glob("mark-*")]
		[mark_diff1(W1s[i[0][0]], S1s[i[0][0]], [j[1] for j in i], old_label, Old_Files[i[0][0]].name)
		 for i in index1]
		[mark_diff1(W2s[i[0][0]], S2s[i[0][0]], [j[1] for j in i], new_label, New_Files[i[0][0]].name)
		 for i in index2]
	
	else:
		if not Only_Unique_Key:
			[os.remove(f) for f in glob.glob("single-*")]
			# disable only old, not needed and much write operate
			# [single_workbook(S1s[i[0][0]], [j[1] for j in i], Key_From_Old, old_label, Old_Files[i[0][0]].name) for i in index1c]
			[single_workbook(S2s[i[0][0]], [j[1] for j in i], Key_From_New, new_label, New_Files[i[0][0]].name) for i in
			 index2]
		# single_workbook(s2, both_index2, Key_From_New, u"both", New_File)
		else:
			[os.remove(f) for f in glob.glob("single-unique-*")]
			[single_unique_workbook(S1s[i[0][0]].title, [j[1] for j in i], old_label, Old_Files[i[0][0]].name) for i in
			 index1o]
			[single_unique_workbook(S2s[i[0][0]].title, [j[1] for j in i], new_label, New_Files[i[0][0]].name) for i in
			 index2o]
	print "== Done"


def diff1(Old_File, New_File, Sheet_Index_From_Old, Sheet_Index_From_New, Key_From_Old, Key_From_New, Single_File,
          Only_Unique_Key):
	w1 = load_workbook(filename=Old_File)  # data_only=True, copied excel need format&style or not ?
	s1 = w1.worksheets[Sheet_Index_From_Old]
	K1 = get_unique_key_seqs(s1, Key_From_Old)
	
	w2 = load_workbook(filename=New_File)
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
		mark_diff1(w1, s1, index1, old_label, Old_File.name)
		mark_diff1(w2, s2, index2, new_label, New_File.name)
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
    Compare 2 Excel groups in line level by set columns as unique KEY.
    Output0: Mark a user defined label to a new last column in new copied excel files named mark-*.
    Output1: Different lines to single files from them.
    Output2: Different and Unique KEY to single files from them.
    ''', formatter_class=RawTextHelpFormatter)
	
	parser.add_argument("-1", dest="Old_Files", nargs="+", required=True, type=file,
	                    help="older excels to compare, many excels must special same sheet and Keys")
	parser.add_argument("-2", dest="New_Files", nargs="+", required=True, type=file,
	                    help="newer excels to compare, many excels must special same sheet and Keys")
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
	parser.add_argument('-v', action='version', version='%(prog)s 2.0')
	args = parser.parse_args()
	
	old_label = args.Mark_Old_Label
	new_label = args.Mark_New_Label
	
	diffPlus(args.Old_Files, args.New_Files, args.Sheet_Index_From_Old, args.Sheet_Index_From_New, args.Key_From_Old,
	         args.Key_From_New, args.Single_File, args.Only_Unique_Key)
