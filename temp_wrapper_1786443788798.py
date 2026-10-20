
import sys
import json
sys.path.insert(0, 'C:\\Users\\dell\\Desktop\\modifier_tool')

from force_exact_totals_12hour import force_exact_totals

totals = json.loads('''{"Kanyonyo Bound":[662,1473,673,119,139,38,7,0,11,9,4,4],"Embu Bound":[206,806,479,93,61,26,2,5,11,6,0,3]}''')
input_file = '''C:\\Users\\dell\\Desktop\\modifier_tool\\temp_input_1786443788798.xlsx'''
output_file = '''C:\\Users\\dell\\Desktop\\modifier_tool\\temp_output_1786443788798.xlsx'''

force_exact_totals(input_file, output_file, totals)
