
import sys
import json
sys.path.insert(0, 'C:\\Users\\dell\\Desktop\\trafficcount')

from force_exact_totals_12hour import force_exact_totals

totals = json.loads('''{"Olkalau Bound":[677,918,446,311,413,37,17,23,279,176,1,1],"Njambini Bound":[644,938,433,258,388,36,10,18,283,138,6,1]}''')
input_file = '''C:\\Users\\dell\\Desktop\\trafficcount\\temp_input_1776745663243.xlsx'''
output_file = '''C:\\Users\\dell\\Desktop\\trafficcount\\temp_output_1776745663243.xlsx'''

force_exact_totals(input_file, output_file, totals)
