import xlsxwriter
import sympy as sy
import fractions as fr
import re


#-----------------------------------------------------------------------------------------------------------------------------------
#  Reading benchmark file and create circuit dictionary with probabilities and vulnerability factors
#-----------------------------------------------------------------------------------------------------------------------------------

benchmark_file_name = 'c17'
delay_file_name = 'delays'

# Open benchmark file for read - OK
with open(f'Inputs/{benchmark_file_name}.bench') as file:
    benchmark_file = file.read()

# Open delay file for read - OK
with open(f'Inputs/Gates Delay/{delay_file_name}.d') as file:
    delay_file = file.read()

# Search gates and primary input dictionary and return value
def get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, key):
    if node in gates_dictionary:
        return gates_dictionary[node][key]
    else:
        return primary_inputs_dictionary[node][key]

# Return one probability of gate - OK
def compute_one_probability(type, p_inputs_dictionary):
    probability = '1'
    if type == 'AND':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = sy.S(f'{probability}*{p1}')
    elif type == 'NAND':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = sy.S(f'{probability}*{p1}')
        probability = sy.S(f'1-{probability}')
    elif type == 'OR':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = sy.S(f'{probability}*{p0}')
        probability = sy.S(f'1-{probability}')
    elif type == 'NOR':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = sy.S(f'{probability}*{p0}')
    elif type == 'XOR':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(2**power):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            if i.count('1') % 2 == 0:
                temp_list_two.append(i)
        for i in temp_list_two:
            temp_list_one.remove(i)
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
    elif type == 'XNOR':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(2**power):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            if i.count('1') % 2 != 0:
                temp_list_two.append(i)
        for i in temp_list_two:
            temp_list_one.remove(i)
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
    elif type == 'NOT':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = sy.S(f'{probability}*{p0}')
    elif type == 'BUFF':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = sy.S(f'{probability}*{p1}')
    elif type == 'KAND':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(int((2**power)/2),(2**power)-1):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
        probability = 1 - probability
    elif type == 'KNAND':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(0,int((2**power)/2)-1):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
    elif type == 'KOR':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(1,int((2**power)/2)):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
    elif type == 'KNOR':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(int((2**power)/2)+1,2**power):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
        probability = 1 - probability
    elif type == 'KXOR':
        probability = 0
        probability_temp = '1'
        temp_list_one = []
        temp_list_two = []
        k = 0
        power = len(p_inputs_dictionary)
        for i in range(2**power):
            binary_string = '{0:0{1}b}'.format(i,power)
            temp_list_one.append(list(binary_string))
        for i in temp_list_one:
            if i.count('1') % 2 == 0:
                temp_list_two.append(i)
        for i in temp_list_two:
            temp_list_one.remove(i)
        for i in temp_list_one:
            for j in p_inputs_dictionary:
                p = p_inputs_dictionary[j][f'P{i[k]}']
                probability_temp = sy.S(f'{probability_temp}*{p}')
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1
    elif type == 'KNOTO':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = sy.S(f'{probability}*{p1}')
        probability = 1 - probability
    elif type == 'KNOTZ':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P0']
            probability = sy.S(f'{probability}*{p1}')
    return probability

def compute_critical_path(circuit_dictionary,primary_output_node):
    stack = []
    previous_node = None
    type = circuit_dictionary[primary_output_node]['Type']
    node_type_delay = circuit_dictionary[primary_output_node]['Node_type_delay']
    critical_path_dictionary = {primary_output_node:{'Type': type, 'Node_type_delay': node_type_delay}}
    stack.extend(circuit_dictionary[primary_output_node]['Input_node_delay'])
    if len(stack) != 0:
        previous_node = stack.pop()
    else:
        return critical_path_dictionary
    while True:
        type = circuit_dictionary[previous_node]['Type']
        node_type_delay = circuit_dictionary[previous_node]['Node_type_delay']
        temp_dic = {previous_node:{'Type': type, 'Node_type_delay': node_type_delay}}
        critical_path_dictionary.update(temp_dic)
        stack.extend(circuit_dictionary[previous_node]['Input_node_delay'])
        if len(stack) != 0:
            previous_node = stack.pop()
        else:
            type = circuit_dictionary[previous_node]['Type']
            node_type_delay = circuit_dictionary[previous_node]['Node_type_delay']
            temp_dic = {previous_node:{'Type': type, 'Node_type_delay': node_type_delay}}
            critical_path_dictionary.update(temp_dic)
            break
    return critical_path_dictionary


# Return inputs of each gates one probabilitiy (ex. 12 = AND(2 ,3) for input 2,3) - OK
def extract_inputs_probability(gates_dictionary, primary_inputs_dictionary, nodes, symbolic_flag):
    p_dic = {}
    p_inputs_dictionary = {}
    if symbolic_flag:
        for node in nodes:
            if node in primary_inputs_dictionary:
                p_dic = {node:
                            {
                             'P1': primary_inputs_dictionary[node]['P1'],
                             'P0': primary_inputs_dictionary[node]['P0']
                            }
                        }
                p_inputs_dictionary.update(p_dic)
            else:
                p_dic = {node:
                            {
                             'P1': gates_dictionary[node]['P1'],
                             'P0': gates_dictionary[node]['P0']
                            }
                        }
                p_inputs_dictionary.update(p_dic)
    else:
        for node in nodes:
            if node in primary_inputs_dictionary:
                p_dic = {node:
                            {
                             'P1': primary_inputs_dictionary[node]['P1_F'],
                             'P0': primary_inputs_dictionary[node]['P0_F']
                            }
                        }
                p_inputs_dictionary.update(p_dic)
            else:
                p_dic = {node:
                            {
                             'P1': gates_dictionary[node]['P1_F'],
                             'P0': gates_dictionary[node]['P0_F']
                            }
                        }
                p_inputs_dictionary.update(p_dic)
    return p_inputs_dictionary

def extract_inputs_delay(gates_dictionary, primary_inputs_dictionary, nodes):
    gates_delay_dictionary = {}
    for node in nodes:
        if node in primary_inputs_dictionary:
            d_dic = {node: primary_inputs_dictionary[node]['Node_delay']}
            gates_delay_dictionary.update(d_dic)
        else:
            d_dic = {node: gates_dictionary[node]['Node_delay']}
            gates_delay_dictionary.update(d_dic)
    return gates_delay_dictionary

# Returns all instances of gate delay pattern in a list - OK
def extract_all_gates_delay(delay_file):
    gate_delay_pattern = re.compile(r'[A-Z0-9]+ ?= ?\d+.?\d?\d*')
    return gate_delay_pattern.findall(delay_file)

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

def create_gates_delay_dictionary(gates_delay_list):
    number_pattern = re.compile(r'\b ?\d+.?(\d+)*\b')
    gate_name_pattern = re.compile(r'([A-Z]+)([0-9]+)?')
    delays_list = [[gate_name_pattern.search(gate_delay).group(), sy.Float(number_pattern.search(gate_delay).group())] for gate_delay in gates_delay_list]
    gates_delay_dictionary = {node[0]:node[1] for node in delays_list}
    return gates_delay_dictionary

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
                                      'Inputs_delay_dictionary':'None',
                                      'Node_delay': 0,
                                      'Node_type_delay': 0,
                                      'Input_node_delay': [],
                                      'Input_node_delay_type': 'None',
                                      'P1': p1,
                                      'P0': p0,
                                      'VF': vf,
                                      'P1_F': sy.Float(p1),
                                      'P0_F': sy.Float(p0),
                                      'VF_F': sy.Float(vf),
                                      'ExPI_node': 'None'
                                      } for node in nodes_list
                                     }
    else:
        primary_inputs_dictionary = {node:
                                     {'Type': 'PI',
                                      'Inputs_delay_dictionary':'None',
                                      'Node_delay': 0,
                                      'Node_type_delay': 0,
                                      'Input_node_delay': [],
                                      'Input_node_delay_type': 'None',
                                      'P1_F': sy.Float(p1),
                                      'P0_F': sy.Float(p0),
                                      'VF_F': sy.Float(vf),
                                      'ExPI_node': 'None'
                                      } for node in nodes_list
                                     }
    return primary_inputs_dictionary

# Create primary outpus dictionary - OK
def create_primary_outputs_dictionary(primary_outputs_list, gates_dictionary, primary_inputs_dictionary, symbolic_flag=True):
    number_pattern = re.compile(r'\d+')
    nodes_list = [number_pattern.search(index).group()
                  for index in primary_outputs_list]
    if symbolic_flag:
        primary_outputs_dictionary = {node:
                                     {'Type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Type')+'(PO)',
                                      'Inputs_delay_dictionary': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Inputs_delay_dictionary'),
                                      'Node_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Node_delay'),
                                      'Node_type_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Node_type_delay'),
                                      'Input_node_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Input_node_delay'),
                                      'Input_node_delay_type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Input_node_delay_type'),
                                      'P1': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1'),
                                      'P0': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0'),
                                      'VF': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF'),
                                      'P1_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1_F')),
                                      'P0_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0_F')),
                                      'VF_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF_F')),
                                      'ExPI_node': 'None'
                                      } for node in nodes_list
                                     }
    else:
        primary_outputs_dictionary = {node:
                                     {'Type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Type')+'(PO)',
                                      'Inputs_delay_dictionary': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Inputs_delay_dictionary'),
                                      'Node_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Node_delay'),
                                      'Node_type_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Node_type_delay'),
                                      'Input_node_delay': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Input_node_delay'),
                                      'Input_node_delay_type': get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'Input_node_delay_type'),
                                      'P1_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P1_F')),
                                      'P0_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'P0_F')),
                                      'VF_F': sy.Float(get_dictionary_value(gates_dictionary,primary_inputs_dictionary, node, 'VF_F')),
                                      'ExPI_node': 'None'
                                      } for node in nodes_list
                                     }
    return primary_outputs_dictionary

# Create gates dictionary - OK
def create_gates_dictionary(gates_list, primary_inputs_dictionary, gates_delay_dictionary, symbolic_flag = True):
    number_pattern = re.compile(r'(\d+)')
    type_pattern = re.compile(r'[A-Z]+')
    temporary_dictionary = {}
    gates_dictionary = {}
    for item in gates_list:
        node = number_pattern.search(item).group()
        inputs_list = number_pattern.findall(item)
        del inputs_list[0]
        gate_type = type_pattern.search(item).group()
        gate_type_inputs_number = str(len(inputs_list)) if len(inputs_list) > 2 else ''
        p_inputs_dictionary = extract_inputs_probability(
            gates_dictionary, primary_inputs_dictionary, inputs_list, symbolic_flag)
        inputs_delay_dictionary = extract_inputs_delay(gates_dictionary,
            primary_inputs_dictionary, inputs_list)
        max_inputs = max(inputs_delay_dictionary.values())
        node_type_delay = gates_delay_dictionary[gate_type+gate_type_inputs_number]
        node_delay = max_inputs + node_type_delay
        inputs_node_delay = [k for k, v in inputs_delay_dictionary.items() if v == max_inputs]
        nodes_delay_type = [get_dictionary_value(gates_dictionary,primary_inputs_dictionary,input_node_delay,'Type') for input_node_delay in inputs_node_delay]
        p1 = compute_one_probability(gate_type, p_inputs_dictionary)
        p0 = sy.S(f'1-{p1}')
        vf = sy.S(f'(1-{p1})-{p1}')
        if symbolic_flag:
            temporary_dictionary = {node: {
                'Type': gate_type,
                'Inputs_delay_dictionary': inputs_delay_dictionary,
                'Node_delay': node_delay,
                'Node_type_delay': node_type_delay,
                'Input_node_delay': inputs_node_delay,
                'Input_node_delay_type': nodes_delay_type,
                'P1': p1,
                'P0': p0,
                'VF': vf,
                'P1_F': sy.Float(p1),
                'P0_F': sy.Float(p0),
                'VF_F': sy.Float(vf),
                'ExPI_node': 'None'
            }}
        else:
            temporary_dictionary = {node: {
                'Type': gate_type,
                'Inputs_delay_dictionary': inputs_delay_dictionary,
                'Node_delay': node_delay,
                'Node_type_delay': node_type_delay,
                'Input_node_delay': inputs_node_delay,
                'Input_node_delay_type': nodes_delay_type,
                'P1_F': sy.Float(p1),
                'P0_F': sy.Float(p0),
                'VF_F': sy.Float(vf),
                'ExPI_node': 'None'
            }}
        gates_dictionary.update(temporary_dictionary)
    return gates_dictionary

# Create circuit dictionary - OK
def create_circuit_dictionary(symbolic_flag = False):
    circuit_dictionary = {}

    # Read delay file and extract all gates delay.
    gates_delay_list = extract_all_gates_delay(delay_file)

    # Create gates delay dictionary from delay file to compute circute gates delay.
    gates_delay_dictionary = create_gates_delay_dictionary(gates_delay_list)

    # Read benchmark file and extract all gates.
    gates_list = extract_all_gates(benchmark_file)

    # Read benchmark file and extract all primary inputs.
    primary_inputs_list = extract_all_primary_inputs(benchmark_file)

    # Read benchmark file and extract all primary outputs.
    primary_output_list = extract_all_primary_outputs(benchmark_file)

    primary_inputs_dictionary = create_primary_inputs_dictionary(
             primary_inputs_list,symbolic_flag)

    circuit_dictionary.update(primary_inputs_dictionary)


    gates_dictionary = create_gates_dictionary(
        gates_list, primary_inputs_dictionary,gates_delay_dictionary,symbolic_flag)


    circuit_dictionary.update(gates_dictionary)


    primary_outputs_dictionary = create_primary_outputs_dictionary(
        primary_output_list, gates_dictionary,primary_inputs_dictionary,symbolic_flag)


    circuit_dictionary.update(primary_outputs_dictionary)


    return circuit_dictionary

def create_circuit_ciritical_path_dictionary(circuit_dictionary):
    circuit_ciritical_path_dictionary = {}
    temp_dic = {}
    primary_output_list = [node for node in circuit_dictionary if 'PO' in circuit_dictionary[node]['Type']]
    for node in primary_output_list:
        temp_dic = {node:{
                'Total_delay': circuit_dictionary[node]['Node_delay'],
                'Path': compute_critical_path(circuit_dictionary,node)
            }}
        circuit_ciritical_path_dictionary.update(temp_dic)
    return circuit_ciritical_path_dictionary

# Make circuit dictionary for algorithm - OK
circuit_dictionary = create_circuit_dictionary()

circuit_ciritical_path_dictionary = create_circuit_ciritical_path_dictionary(circuit_dictionary)


#-----------------------------------------------------------------------------------------------------------------------------------
#  Writing first results of benchmark vulnerabilities to xlsx file
#-----------------------------------------------------------------------------------------------------------------------------------

def create_vulnerability_factor_xlsx_file(benchmark_file_name,circuit_dictionary):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}.xlsx')
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
    headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)','Gate Delay')
    worksheet.set_column('A1:XFD1048576',30,sheet_format)
    worksheet.set_column('F:F',40,sheet_format)
    worksheet.set_row(0,30)
    worksheet.merge_range('A1:F1', f'Benchmark {benchmark_file_name} Vulnerability Factor', header_cells_format)
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
        worksheet.write(row, col + 5, circuit_dictionary[node]['Node_delay'],cells_format)
        row += 1

    worksheet.conditional_format(f'E2:E{row}', {'type': 'data_bar','bar_color':
        f'{vulnerability_bg_color}','bar_negative_color_same': True,'bar_no_border': True})

    worksheet.conditional_format(f'A2:A{row}', {'type':'formula',
        'criteria':f'=ABS(E2:E{row})>0.95','format': vulnerability_cells_format})

    workbook.close()

def create_critical_path_xlsx_file(benchmark_file_name, circuit_ciritical_path_dictionary):

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}_critical_path.xlsx')
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    title_cells_format = workbook.add_format()
    header_cell_format = workbook.add_format()
    cells_format = workbook.add_format()
    sheet_format = workbook.add_format()

    # Color variables
    border_color = '#FFFFFF'
    cell_bg_color = '#ffeedd'
    title_bg_color = '#eaeaea'
    font_color = '#616161'
    sheet_color = 'f2f2f2'
    


    cells_format.set_border()
    cells_format.set_border_color(border_color)
    cells_format.set_bg_color(cell_bg_color)
    cells_format.set_font_color(font_color)
    cells_format.set_align('left')
    cells_format.set_font_size(12)


    title_cells_format.set_border()
    title_cells_format.set_border_color(border_color)
    title_cells_format.set_align('vcenter')
    title_cells_format.set_bg_color(title_bg_color)
    title_cells_format.set_font_size(12)
    title_cells_format.set_font_color(font_color)

    header_cell_format.set_align('vcenter')
    header_cell_format.set_bg_color(sheet_color)
    header_cell_format.set_font_size(18)
    header_cell_format.set_font_color(font_color)
    header_cell_format.set_bold()


    sheet_format.set_bg_color(sheet_color)

    # Write some data headers.
    row = 1
    col = 0
    worksheet.set_column('A1:XFD1048576',15,sheet_format)
    # worksheet.set_column('F:F',40,sheet_format)
    worksheet.set_row(0,50)
    worksheet.merge_range('A1:F1', f'Benchmark {benchmark_file_name} Critical Pathes', header_cell_format)
    
   
    # Iterate over the data and write it out row by row.
    for node in circuit_ciritical_path_dictionary:
        path = circuit_ciritical_path_dictionary[node]['Path']
        for path_node in path:
             if col == 0:
                worksheet.write(row, col, 'Type',title_cells_format)
                worksheet.write(row + 1, col, 'Path node',title_cells_format)
                col += 1
             worksheet.write(row, col, path[path_node]['Type'],cells_format)
             worksheet.write(row + 1, col, int(path_node),cells_format)
             col += 1
        worksheet.write(row + 2,0 , 'Path delay',title_cells_format)
        worksheet.write(row + 2,1 , circuit_ciritical_path_dictionary[node]['Total_delay'],title_cells_format)
        col = 0
        row += 4 

    workbook.close()

create_vulnerability_factor_xlsx_file(benchmark_file_name,circuit_dictionary)
create_critical_path_xlsx_file(benchmark_file_name, circuit_ciritical_path_dictionary)