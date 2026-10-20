
import sys
import json
sys.path.insert(0, 'C:\\Users\\dell\\Desktop\\modifier_tool')

from force_exact_totals_12hour import force_exact_totals

totals = json.loads('''{"Nakuru Bound":[805,864,311,333,196,14,10,63,94,43,9,13],"Marigat Bound":[1014,903,314,251,263,30,10,7,115,33,2,9]}''')
input_file = '''C:\\Users\\dell\\Desktop\\modifier_tool\\temp_input_1792488670604.xlsx'''
output_file = '''C:\\Users\\dell\\Desktop\\modifier_tool\\temp_output_1792488670604.xlsx'''

force_exact_totals(input_file, output_file, totals)
