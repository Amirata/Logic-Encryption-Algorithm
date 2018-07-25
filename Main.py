import xlsxwriter
import sympy as sy
import fractions as fr
import re


#-----------------------------------------------------------------------------------------------------------------------------------
#  Reading benchmark file and create circuit dictionary with probabilities and vulnerability factors
#-----------------------------------------------------------------------------------------------------------------------------------

file_name = 'c17'

# Open benchmark file for read - OK
with open(f'Inputs/{file_name}.bench') as file:
    benchmark_file = file.read()

# Search gates and primary input dictionary and return value
def get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, key):
    if node in gates_dictionary:
        return gates_dictionary[node][key]
    else:
        return primary_inputs_dictionary[node][key]

# Return one probability of gate - OK
def compute_one_probability(type, *p1_inputs):
    probability = '1'
    probability_temp = '1'
    if type == 'AND':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*{i}')
    elif type == 'NAND':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*{i}')
        probability = sy.S(f'1-{probability}')
    elif type == 'OR':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*(1-{i})')
        probability = sy.S(f'1-{probability}')
    elif type == 'NOR':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*(1-{i})')
    elif type == 'XOR':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*{i}')
        for i in p1_inputs:
            probability_temp = sy.S(f'{probability_temp}*(1-{i})')
        probability = sy.S(f'1-({probability}+{probability_temp})')
    elif type == 'XNOR':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*{i}')
        for i in p1_inputs:
            probability_temp = sy.S(f'{probability_temp}*(1-{i})')
        probability = sy.S(f'{probability}+{probability_temp}')
    elif type == 'NOT':
        for i in p1_inputs:
            probability = sy.S(f'1-({probability}*{i})')
    elif type == 'BUFF':
        for i in p1_inputs:
            probability = sy.S(f'{probability}*{i}')
    return probability

# Return inputs of each gates one probabilitiy (ex. 12 = AND(2 ,3) for input 2,3) - OK
def extract_inputs_one_probability(gates_dictionary, primary_inputs_dictionary, nodes, symbolic_flag):
    p1_list = []
    if symbolic_flag:
        for node in nodes:
            if node in primary_inputs_dictionary:
                p1_list.append(primary_inputs_dictionary[node]['P1'])
            else:
                p1_list.append(gates_dictionary[node]['P1'])
    else:
        for node in nodes:
            if node in primary_inputs_dictionary:
                p1_list.append(primary_inputs_dictionary[node]['P1_F'])
            else:
                p1_list.append(gates_dictionary[node]['P1_F'])
    return p1_list

# Returns all instances of primary input pattern in a list - OK
def extract_all_primary_inputs(benchmark_file):
    benchmark_primary_input_pattern = re.compile(r'INPUT\(\d+\)')
    return benchmark_primary_input_pattern.findall(benchmark_file)

# Returns all instances of primary output pattern in a list - OK
def extract_all_primary_outputs(benchmark_file):
    benchmark_primary_output_pattern = re.compile(r'OUTPUT\(\d+\)')
    return benchmark_primary_output_pattern.findall(benchmark_file)

# Returns all instances of gate pattern in a list - OK
def extract_all_gates(benchmark_file):
    benchmark_gate_pattern = re.compile(r'\d+ = [A-Z]+\([\d\,? ?]+\)')
    return benchmark_gate_pattern.findall(benchmark_file)

# Create primary inputs dictionary - OK
def create_primary_inputs_dictionary(primary_inputs_list, symbolic_flag=True):
    number_pattern = re.compile(r'\d+')
    nodes_list = [number_pattern.search(index).group()
                  for index in primary_inputs_list]
    p1 = sy.S('1/2')
    p0 = sy.S('1/2')
    vf = sy.S('0')
    if symbolic_flag:
        primary_inputs_dictionary = {node:
                                     {'Type': 'PI',
                                      'P1': p1,
                                      'P0': p0,
                                      'VF': vf,
                                      'P1_F': sy.Float(p1),
                                      'P0_F': sy.Float(p0),
                                      'VF_F': sy.Float(vf),
                                      'ExPI_ID': 'None'
                                      } for node in nodes_list
                                     }
    else:
        primary_inputs_dictionary = {node:
                                     {'Type': 'PI',
                                      'P1_F': sy.Float(p1),
                                      'P0_F': sy.Float(p0),
                                      'VF_F': sy.Float(vf),
                                      'ExPI_ID': 'None'
                                      } for node in nodes_list
                                     }
    return primary_inputs_dictionary

# Create primary outpus dictionary - OK
def create_primary_outputs_dictionary(primary_outputs_list, gates_dictionary, primary_inputs_dictionary, symbolic_flag=True):
    number_pattern = re.compile(r'\d+')
    nodes_list = [number_pattern.search(index).group()
                  for index in primary_outputs_list]
    if symbolic_flag:
        primary_inputs_dictionary = {node:
                                     {'Type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Type')+'(PO)',
                                      'P1': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1'),
                                      'P0': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0'),
                                      'VF': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF'),
                                      'P1_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1_F')),
                                      'P0_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0_F')),
                                      'VF_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF_F')),
                                      'ExPI_ID': 'None'
                                      } for node in nodes_list
                                     }
    else:
        primary_inputs_dictionary = {node:
                                     {'Type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Type')+'(PO)',
                                      'P1_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1_F')),
                                      'P0_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0_F')),
                                      'VF_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF_F')),
                                      'ExPI_ID': 'None'
                                      } for node in nodes_list
                                     }
    return primary_inputs_dictionary

# Create gates dictionary - OK
def create_gates_dictionary(gates_list, primary_inputs_dictionary, symbolic_flag = True):
    number_pattern = re.compile(r'(\d+)')
    type_pattern = re.compile(r'[A-Z]+')
    temporary_dictionary = {}
    gates_dictionary = {}
    for item in gates_list:
        node = number_pattern.search(item).group()
        inputs_list = number_pattern.findall(item)
        del inputs_list[0]
        gate_type = type_pattern.search(item).group()
        p1_inputs = extract_inputs_one_probability(
            gates_dictionary, primary_inputs_dictionary, inputs_list, symbolic_flag)
        p1 = compute_one_probability(gate_type, *p1_inputs)
        p0 = sy.S(f'1-{p1}')
        vf = sy.S(f'(1-{p1})-{p1}')
        if symbolic_flag:
            temporary_dictionary = {node: {
                'Type': gate_type,
                'P1': p1,
                'P0': p0,
                'VF': vf,
                'P1_F': sy.Float(p1),
                'P0_F': sy.Float(p0),
                'VF_F': sy.Float(vf),
                'ExPI_ID': 'None'
            }}
        else:
            temporary_dictionary = {node: {
                'Type': gate_type,
                'P1_F': sy.Float(p1),
                'P0_F': sy.Float(p0),
                'VF_F': sy.Float(vf),
                'ExPI_ID': 'None'
            }}
        gates_dictionary.update(temporary_dictionary)
    return gates_dictionary

# Create circuit dictionary - OK
def create_circuit_dictionary(symbolic_flag = False):
    circuit_dictionary = {}
    gates_list = extract_all_gates(benchmark_file)
    primary_inputs_list = extract_all_primary_inputs(benchmark_file)
    primary_output_list = extract_all_primary_outputs(benchmark_file)
    primary_inputs_dictionary = create_primary_inputs_dictionary(
             primary_inputs_list,symbolic_flag)
    circuit_dictionary.update(primary_inputs_dictionary)
    gates_dictionary = create_gates_dictionary(
        gates_list, primary_inputs_dictionary,symbolic_flag)
    circuit_dictionary.update(gates_dictionary)
    primary_outputs_dictionary = create_primary_outputs_dictionary(
        primary_output_list, gates_dictionary,primary_inputs_dictionary,symbolic_flag)
    circuit_dictionary.update(primary_outputs_dictionary)
    return circuit_dictionary

# Make circuit dictionary for algorithm - OK
circuit_dictionary = create_circuit_dictionary()


#-----------------------------------------------------------------------------------------------------------------------------------
#  Writing first results of benchmark vulnerabilities to xlsx file
#-----------------------------------------------------------------------------------------------------------------------------------


# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(f'Outputs/{file_name}.xlsx')
worksheet = workbook.add_worksheet()


# Add a bold format to use to highlight cells.
header_cells_format = workbook.add_format()
cells_format = workbook.add_format()
vulnerability_cells_format = workbook.add_format()
sheet_format = workbook.add_format()

# Color variables
border_color = '#fefefe'
cell_bg_color = '#f2f2f2'
header_bg_color = '#ffab91'
font_color = '#616161'
vulnerability_bg_color = '#ffab91'


cells_format.set_border()
cells_format.set_border_color(border_color)
cells_format.set_bg_color(cell_bg_color)
cells_format.set_font_color(font_color)
cells_format.set_align('left')
cells_format.set_font_size(12)

header_cells_format.set_align('center')
header_cells_format.set_align('vcenter')
header_cells_format.set_bg_color(header_bg_color)
header_cells_format.set_font_size(14)
header_cells_format.set_font_color(font_color)
header_cells_format.set_bold()

sheet_format.set_bg_color(cell_bg_color)

vulnerability_cells_format.set_bg_color(vulnerability_bg_color)
# Write some data headers.
row = 1
col = 0
headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)')
worksheet.set_column('A1:XFD1048576',30,sheet_format)
worksheet.set_row(0,30)
worksheet.merge_range('A1:E1', f'Benchmark {file_name} Vulnerability Factor', header_cells_format)
for header_data in headers_data:
    worksheet.set_row(row,20)
    worksheet.write(row,col,header_data,header_cells_format)
    col += 1

# Start from the second cell below the headers
row += 1
col = 0
# Iterate over the data and write it out row by row.
for node in circuit_dictionary:
    worksheet.write(row, col, circuit_dictionary[node]['Type'],cells_format)
    worksheet.write(row, col + 1, int(node), cells_format)
    worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],cells_format)
    worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],cells_format)
    worksheet.write(row, col + 4, circuit_dictionary[node]['VF_F'],cells_format)    
    row += 1

worksheet.conditional_format(f'E2:E{row}', {'type': 'data_bar','bar_color': 
    f'{vulnerability_bg_color}','bar_negative_color_same': True,'bar_no_border': True})

worksheet.conditional_format(f'A2:A{row}', {'type':'formula',
    'criteria':f'=ABS(E2:E{row})>0.95','format': vulnerability_cells_format})

workbook.close()
    
