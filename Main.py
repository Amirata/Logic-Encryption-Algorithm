from copy import deepcopy
import numpy as np
import xlsxwriter
import os
import re


def main():
    priority(4)
    benchmark_file_name = 'test_bench1'
    VFth = 0.9
    NoR = 30
    delay_file_name = 'delays'

    # Read delay file and extract all gates delay.
    gates_delay_list = extract_all_gates_delay(delay_file_name)

    # Create gates delay dictionary from delay file to compute circute gates delay.
    gates_delay_dictionary = create_gates_delay_dictionary(gates_delay_list)

    # Make circuit dictionary for algorithm - OK
    circuit_dictionary = create_circuit_dictionary(benchmark_file_name,deepcopy(gates_delay_dictionary))

    # circuit_ciritical_path_dictionary = create_circuit_ciritical_path_dictionary(circuit_dictionary)

    # eDN = real_algorithm(deepcopy(circuit_dictionary),VFth,NoR,benchmark_file_name,deepcopy(gates_delay_dictionary),True,True)

    real_algorithm_for_VFth_range(deepcopy(circuit_dictionary),NoR,benchmark_file_name,deepcopy(gates_delay_dictionary))
   
    # eDN = my_eal_algorithm(deepcopy(circuit_dictionary),VFth,NoR,benchmark_file_name,deepcopy(gates_delay_dictionary))

    # my_real_algorithm_for_VFth_range(deepcopy(circuit_dictionary),NoR,benchmark_file_name,deepcopy(gates_delay_dictionary))


def priority(priority = 2):
    """ Set the priority of the process to below-normal."""
    
    import sys
    try:
        sys.getwindowsversion()
    except AttributeError:
        isWindows = False
    else:
        isWindows = True

    if isWindows:
        # Based on:
        #   "Recipe 496767: Set Process Priority In Windows" on ActiveState
        #   http://code.activestate.com/recipes/496767/
        import win32api,win32process,win32con
        priorityclasses = [win32process.IDLE_PRIORITY_CLASS,
                       win32process.BELOW_NORMAL_PRIORITY_CLASS,
                       win32process.NORMAL_PRIORITY_CLASS,
                       win32process.ABOVE_NORMAL_PRIORITY_CLASS,
                       win32process.HIGH_PRIORITY_CLASS,
                       win32process.REALTIME_PRIORITY_CLASS]
        pid = win32api.GetCurrentProcessId()
        handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid)
        win32process.SetPriorityClass(handle, priorityclasses[priority])
    else:
        import os

        os.nice(1)

# Search gates and primary input dictionary and return value
def get_dictionary_value(node, key, gates_dictionary,primary_inputs_dictionary = []):
    if node in gates_dictionary:
        return gates_dictionary[node][key]
    else:
        return primary_inputs_dictionary[node][key]

# Return one probability of gate
def compute_one_probability(type, p_inputs_dictionary):
    probability = 1.0
    if type == 'AND':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = probability * p1
    elif type == 'NAND':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = probability * p1
        probability = 1.0 - probability
    elif type == 'OR':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = probability * p0
        probability = 1.0 - probability
    elif type == 'NOR':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = probability * p0
    elif type == 'XOR':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k += 1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
    elif type == 'XNOR':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
    elif type == 'NOT':
        for i in p_inputs_dictionary:
            p0 = p_inputs_dictionary[i]['P0']
            probability = probability * p0
    elif type == 'BUFF':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = probability * p1
    elif type == 'KAND':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
        probability = 1.0 - probability
    elif type == 'KNAND':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
    elif type == 'KOR':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
    elif type == 'KNOR':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
        probability = 1.0 - probability
    elif type == 'KXOR':
        probability = 0.0
        probability_temp = 1.0
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
                probability_temp = probability_temp * p
                k+=1
            k = 0
            probability += probability_temp
            probability_temp = 1.0
    elif type == 'KNOTO':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P1']
            probability = probability * p1
        probability = 1.0 - probability
    elif type == 'KNOTZ':
        for i in p_inputs_dictionary:
            p1 = p_inputs_dictionary[i]['P0']
            probability = probability * p1
    elif type == 'KBUFFO':
        k = 1
        for i in p_inputs_dictionary:
            p = p_inputs_dictionary[i][f'P{k}']
            probability = probability * p
            k = 0
        probability = 1.0 - probability
    elif type == 'KBUFFZ':
        k = 0
        for i in p_inputs_dictionary:
            p = p_inputs_dictionary[i][f'P{k}']
            probability = probability * p
            k = 1
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

# Return inputs of each gates one probabilitiy (ex. 12 = AND(2 ,3) for input 2,3)
def extract_inputs_probability(nodes, gates_dictionary, primary_inputs_dictionary = []):
    p_dic = {}
    p_inputs_dictionary = {}
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


def extract_inputs_delay(nodes, gates_dictionary, primary_inputs_dictionary = []):
    gates_delay_dictionary = {}
    for node in nodes:
        if node in primary_inputs_dictionary:
            d_dic = {node: primary_inputs_dictionary[node]['Node_delay']}
            gates_delay_dictionary.update(d_dic)
        else:
            d_dic = {node: gates_dictionary[node]['Node_delay']}
            gates_delay_dictionary.update(d_dic)
    return gates_delay_dictionary

# Returns all instances of gate delay pattern in a list
def extract_all_gates_delay(delay_file_name):
    # Open delay file for read - OK++++++++++++++++++
    with open(f'Inputs/Gates Delay/{delay_file_name}.d') as file:
        delay_file = file.read()
    gate_delay_pattern = re.compile(r'[A-Z0-9]+ ?= ?\d+.?\d?\d*')
    return gate_delay_pattern.findall(delay_file)


def create_gates_delay_dictionary(gates_delay_list):
    number_pattern = re.compile(r'\b ?\d+.?(\d+)*\b')
    gate_name_pattern = re.compile(r'([A-Z]+)([0-9]+)?')
    delays_list = [[gate_name_pattern.search(gate_delay).group(), float(number_pattern.search(gate_delay).group())] for gate_delay in gates_delay_list]
    gates_delay_dictionary = {node[0]:node[1] for node in delays_list}
    return gates_delay_dictionary

# Create circuit dictionary
def create_circuit_dictionary(benchmark_file_name,gates_delay_dictionary):
    
    # Open benchmark file for read - OK++++++++++++++++
    with open(f'Inputs/{benchmark_file_name}.bench') as file:
        benchmark_file = file.read()
    
    # Returns all instances of gate pattern in a list - OK++++++++++
    def extract_all_gates(benchmark_file):
        benchmark_gate_pattern = re.compile(r'\d+ = [A-Z]+\([\d\,? ?]+\)')
        return benchmark_gate_pattern.findall(benchmark_file)
    # Returns all instances of primary input pattern in a list - OK+++++++++
    def extract_all_primary_inputs(benchmark_file):
        benchmark_primary_input_pattern = re.compile(r'INPUT\(\d+\)')
        return benchmark_primary_input_pattern.findall(benchmark_file)
    # Returns all instances of primary output pattern in a list - OK+++++++++++
    def extract_all_primary_outputs(benchmark_file):
        benchmark_primary_output_pattern = re.compile(r'OUTPUT\(\d+\)')
        return benchmark_primary_output_pattern.findall(benchmark_file)
    # Create primary inputs dictionary - OK++++++++++++
    def create_primary_inputs_dictionary(primary_inputs_list):
        number_pattern = re.compile(r'\d+')
        nodes_list = [int(number_pattern.search(index).group())
                      for index in primary_inputs_list]
        p1 = 0.5
        p0 = 0.5
        vf = 0.0
        primary_inputs_dictionary = {node:
                                     {'Type': 'PI',
                                      'Inputs_list': [],
                                      'Inputs_delay_dictionary':None,
                                      'Node_delay': 0.0,
                                      'Node_type_delay': 0.0,
                                      'Input_node_delay': [],
                                      'Input_node_delay_type': None,
                                      'P1_F': p1,
                                      'P0_F': p0,
                                      'VF_F': vf,
                                      'ExPI_node': None
                                      } for node in nodes_list
                                     }
        return primary_inputs_dictionary

    # Create gates dictionary - OK+++++++++++
    def create_gates_dictionary(gates_list, primary_inputs_dictionary, gates_delay_dictionary):
        def list_to_int(l):
            return [int(i) for i in l]
        number_pattern = re.compile(r'(\d+)')
        type_pattern = re.compile(r'[A-Z]+')
        temporary_dictionary = {}
        gates_dictionary = {}
        for item in gates_list:
            #gate inputs with own node number at first.
            inputs_list = list_to_int(number_pattern.findall(item))
            #remove gate number from inputs list.
            node = inputs_list.pop(0)
            gate_type = type_pattern.search(item).group()
            # add inputs number to gate to compute delay.
            gate_type_inputs_number = str(len(inputs_list)) if len(inputs_list) > 2 else ''
            p_inputs_dictionary = extract_inputs_probability(inputs_list,
                gates_dictionary, primary_inputs_dictionary)
            inputs_delay_dictionary = extract_inputs_delay(inputs_list, gates_dictionary,
                primary_inputs_dictionary)
            max_inputs = max(inputs_delay_dictionary.values())
            #return gate delay Ex. AND3 = 23.3ps
            node_type_delay = gates_delay_dictionary[gate_type + gate_type_inputs_number]
            node_delay = max_inputs + node_type_delay
            inputs_node_delay = [k for k, v in inputs_delay_dictionary.items() if v == max_inputs]
            nodes_delay_type = [get_dictionary_value(input_node_delay,'Type', gates_dictionary,primary_inputs_dictionary) for input_node_delay in inputs_node_delay]
            p1 = compute_one_probability(gate_type, p_inputs_dictionary)
            p0 = 1.0 - p1
            vf = p0 - p1
            temporary_dictionary = {node: {
                'Type': gate_type,
                'Inputs_list':inputs_list,
                'Inputs_delay_dictionary': inputs_delay_dictionary,
                'Node_delay': node_delay,
                'Node_type_delay': node_type_delay,
                'Input_node_delay': inputs_node_delay,
                'Input_node_delay_type': nodes_delay_type,
                'P1_F': p1,
                'P0_F': p0,
                'VF_F': vf,
                'ExPI_node': None
            }}
            gates_dictionary.update(temporary_dictionary)
        return gates_dictionary
    
    # Create primary outpus dictionary - OK+++++++++++++
    def create_primary_outputs_dictionary(primary_outputs_list, gates_dictionary, primary_inputs_dictionary):
        number_pattern = re.compile(r'\d+')
        nodes_list = [int(number_pattern.search(index).group())
                      for index in primary_outputs_list]
        primary_outputs_dictionary = {node:
                                     {'Type': get_dictionary_value(node, 'Type', gates_dictionary,primary_inputs_dictionary)+'(PO)',
                                      'Inputs_list':get_dictionary_value(node, 'Inputs_list', gates_dictionary,primary_inputs_dictionary),
                                      'Inputs_delay_dictionary': get_dictionary_value(node, 'Inputs_delay_dictionary', gates_dictionary,primary_inputs_dictionary),
                                      'Node_delay': get_dictionary_value(node, 'Node_delay', gates_dictionary,primary_inputs_dictionary),
                                      'Node_type_delay': get_dictionary_value(node, 'Node_type_delay', gates_dictionary,primary_inputs_dictionary),
                                      'Input_node_delay': get_dictionary_value(node, 'Input_node_delay', gates_dictionary,primary_inputs_dictionary),
                                      'Input_node_delay_type': get_dictionary_value(node, 'Input_node_delay_type', gates_dictionary,primary_inputs_dictionary),
                                      'P1_F': get_dictionary_value(node, 'P1_F', gates_dictionary,primary_inputs_dictionary),
                                      'P0_F': get_dictionary_value(node, 'P0_F', gates_dictionary,primary_inputs_dictionary),
                                      'VF_F': get_dictionary_value(node, 'VF_F', gates_dictionary,primary_inputs_dictionary),
                                      'ExPI_node': None
                                      } for node in nodes_list
                                     }
        return primary_outputs_dictionary
    
    circuit_dictionary = {}

    # Read benchmark file and extract all gates.
    gates_list = extract_all_gates(benchmark_file)

    # Read benchmark file and extract all primary inputs.
    primary_inputs_list = extract_all_primary_inputs(benchmark_file)

    # Read benchmark file and extract all primary outputs.
    primary_output_list = extract_all_primary_outputs(benchmark_file)

    primary_inputs_dictionary = create_primary_inputs_dictionary(
             primary_inputs_list)

    circuit_dictionary.update(primary_inputs_dictionary)


    gates_dictionary = create_gates_dictionary(
        gates_list, primary_inputs_dictionary,gates_delay_dictionary)


    circuit_dictionary.update(gates_dictionary)


    primary_outputs_dictionary = create_primary_outputs_dictionary(
        primary_output_list, gates_dictionary,primary_inputs_dictionary)


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

#-----------------------------------------------------------------------------------------------------------------------------------
#  Algorithm
#-----------------------------------------------------------------------------------------------------------------------------------

def get_replace_gate_type(node_type):
    if node_type == 'AND':
        return 'KAND'
    elif node_type == 'OR':
        return 'KOR'
    elif node_type == 'NAND':
        return 'KNAND'
    elif node_type == 'NOR':
        return 'KNOR'
    elif node_type == 'XOR':
        return 'KXOR'
    elif node_type == 'XNOR':
        return 'KXNOR'
    elif node_type == 'NOT':
        return 'KNOT'
    elif node_type == 'BUFF':
        return 'KBUFF'
    elif node_type == 'AND(PO)':
        return 'KAND(PO)'
    elif node_type == 'OR(PO)':
        return 'KOR(PO)'
    elif node_type == 'NAND(PO)':
        return 'KNAND(PO)'
    elif node_type == 'NOR(PO)':
        return 'KNOR(PO)'
    elif node_type == 'XOR(PO)':
        return 'KXOR(PO)'
    elif node_type == 'XNOR(PO)':
        return 'KXNOR(PO)'
    elif node_type == 'NOT(PO)':
        return 'KNOT(PO)'
    elif node_type == 'BUFF(PO)':
        return 'KBUFF(PO)'
    else:
        return None

def Replace_Gate(eDN_t, node ,gates_delay_dictionary):
    
    new_node_type = get_replace_gate_type(eDN_t[node]['Type'])
    if new_node_type == None:
        return eDN_t
    if new_node_type == 'KNOT':
        if eDN_t[node]['VF_F'] < 0.0:
            new_node_type = 'KNOTZ'
        else:
            new_node_type = 'KNOTO'
    elif new_node_type == 'KNOT(PO)':
        if eDN_t[node]['VF_F'] < 0.0:
            new_node_type = 'KNOTZ(PO)'
        else:
            new_node_type = 'KNOTO(PO)'
    elif new_node_type == 'KBUFF':
        if eDN_t[node]['VF_F'] < 0.0:
            new_node_type = 'KBUFFZ'
        else:
            new_node_type = 'KBUFFO'
    elif new_node_type == 'KBUFF(PO)':
        if eDN_t[node]['VF_F'] < 0.0:
            new_node_type = 'KBUFFZ(PO)'
        else:
            new_node_type = 'KBUFFO(PO)'
    node_list = [node for node in eDN_t.keys()]
    maximum_node_number = max(node_list)
    new_node_number = maximum_node_number + 1
    primary_input_dictionary = {new_node_number:
                                     {'Type': 'PI',
                                      'Inputs_list': [],
                                      'Inputs_delay_dictionary':None,
                                      'Node_delay': 0.0,
                                      'Node_type_delay': 0.0,
                                      'Input_node_delay': [],
                                      'Input_node_delay_type': None,
                                      'P1_F': 0.5,
                                      'P0_F': 0.5,
                                      'VF_F': 0,
                                      'ExPI_node': None
                                      }
                                     }
    eDN_t.update(primary_input_dictionary)
    eDN_t[node]['Type'] = new_node_type
    eDN_t[node]['Inputs_list'].insert(0,new_node_number)
    eDN_t[node]['ExPI_node'] = new_node_number
    for node in eDN_t:
        if eDN_t[node]['Type'] != 'PI' and eDN_t[node]['Type'] != 'PI(PO)':
            inputs_list = eDN_t[node]['Inputs_list']
            gate_type = eDN_t[node]['Type']
            if 'PO' in gate_type:
                gate_type = gate_type[:-4]
            if 'K' in eDN_t[node]['Type']:
                gate_type_inputs_number = str(len(inputs_list)) if len(inputs_list) > 3 else ''
            else:
                gate_type_inputs_number = str(len(inputs_list)) if len(inputs_list) > 2 else ''
            p_inputs_dictionary = extract_inputs_probability(inputs_list, eDN_t)
            inputs_delay_dictionary = extract_inputs_delay(inputs_list, eDN_t)
            max_inputs = max(inputs_delay_dictionary.values())
            node_type_delay = gates_delay_dictionary[gate_type + gate_type_inputs_number]
            node_delay = max_inputs + node_type_delay
            inputs_node_delay = [k for k, v in inputs_delay_dictionary.items() if v == max_inputs]
            nodes_delay_type = [get_dictionary_value(input_node_delay,'Type',eDN_t) for input_node_delay in inputs_node_delay]
            p1 = compute_one_probability(gate_type, p_inputs_dictionary)
            p0 = 1.0 - p1
            vf = p0 - p1
            eDN_t[node]['Inputs_delay_dictionary'] = inputs_delay_dictionary
            eDN_t[node]['Node_delay'] = node_delay
            eDN_t[node]['Node_type_delay'] = node_type_delay
            eDN_t[node]['Input_node_delay'] = inputs_node_delay
            eDN_t[node]['Input_node_delay_type'] = nodes_delay_type
            eDN_t[node]['P1_F'] = p1
            eDN_t[node]['P0_F'] = p0
            eDN_t[node]['VF_F'] = vf
    
    return eDN_t

def get_CN(eDN,NoVN,VFth,benchmark_file_name,iteration,gates_delay_dictionary ,get_cn_write_xlsx_flag):
    # for excell file
    row = 0
    workbook = None
    worksheet = None
    if get_cn_write_xlsx_flag:
        newpath = f'Outputs/{benchmark_file_name}' 
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_get_CN_iteration_{iteration}.xlsx')
        worksheet = workbook.add_worksheet()
    def create_vulnerability_factor_xlsx_file(circuit_dictionary,iteration_node,VFth):
        # Create a workbook and add a worksheet.
        nonlocal worksheet
        nonlocal row

        # Add a bold format to use to highlight cells.
        header_cells_format = workbook.add_format()
        cells_format = workbook.add_format()
        key_gate_cells_format = workbook.add_format()
        vulnerability_cells_format = workbook.add_format()
        footer_cells_format = workbook.add_format()
        sheet_format = workbook.add_format()

        # Color variables
        border_color = '#ffffff'
        cell_bg_color = '#f2f2f2'
        header_bg_color = '#E6E6E6'
        font_color = '#616161'
        vulnerability_bg_color = '#FFCCCC'
        key_gate_color = '#CCECFF'
        data_bar_color = '#FFADAE'

        cells_format.set_border()
        cells_format.set_border_color(border_color)
        cells_format.set_bg_color(cell_bg_color)
        cells_format.set_font_color(font_color)
        cells_format.set_align('left')
        cells_format.set_font_size(12)
        
        footer_cells_format.set_border()
        footer_cells_format.set_border_color(border_color)
        footer_cells_format.set_bg_color(header_bg_color)
        footer_cells_format.set_font_color(font_color)
        footer_cells_format.set_align('left')
        footer_cells_format.set_font_size(12)
        footer_cells_format.set_bold()
        
        key_gate_cells_format.set_border()
        key_gate_cells_format.set_border_color(border_color)
        key_gate_cells_format.set_bg_color(key_gate_color)
        key_gate_cells_format.set_font_color(font_color)
        key_gate_cells_format.set_align('left')
        key_gate_cells_format.set_font_size(12)
        
        vulnerability_cells_format.set_border()
        vulnerability_cells_format.set_border_color(border_color)
        vulnerability_cells_format.set_bg_color(vulnerability_bg_color)
        vulnerability_cells_format.set_font_color(font_color)
        vulnerability_cells_format.set_align('left')
        vulnerability_cells_format.set_font_size(12)

        header_cells_format.set_align('center')
        header_cells_format.set_align('vcenter')
        header_cells_format.set_bg_color(header_bg_color)
        header_cells_format.set_font_size(14)
        header_cells_format.set_font_color(font_color)
        header_cells_format.set_bold()

        sheet_format.set_bg_color(cell_bg_color)

        
        
        headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)')
        worksheet.set_column('A1:XFD1048576',30,sheet_format)
        
        # {benchmark_file_name} Real Algorithm Vulnerability Factor
        if iteration_node == None:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Original Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        else:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Node = {iteration_node} Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        # Write some data headers.
        row += 1
        col = 0
        
        for header_data in headers_data:
            worksheet.set_row(row,20)
            worksheet.write(row,col,header_data,header_cells_format)
            col += 1

        # Start from the second cell below the headers
        row += 1
        col = 0
        first_row = row
        # Iterate over the data and write it out row by row.
        for node in circuit_dictionary:
            if 'K' in circuit_dictionary[node]['Type']:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],key_gate_cells_format)
                worksheet.write(row, col + 1, int(node), key_gate_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],key_gate_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],key_gate_cells_format)
            elif abs(circuit_dictionary[node]['VF_F']) >= VFth:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],vulnerability_cells_format)
                worksheet.write(row, col + 1, int(node), vulnerability_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],vulnerability_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],vulnerability_cells_format)
            else:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],cells_format)
                worksheet.write(row, col + 1, int(node), cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],cells_format)
            worksheet.write(row, col + 4, circuit_dictionary[node]['VF_F'],cells_format)
            row += 1
        worksheet.write(row, col, 'Sum Of VF :',footer_cells_format)
        worksheet.write(row, col + 1, f'=SUMIF(E{first_row + 1}:E{row},">0")+ABS(SUMIF(E{first_row + 1}:E{row},"<0"))',footer_cells_format)
        worksheet.write(row + 1, col, f'Number Of VF Above {VFth}',footer_cells_format)
        worksheet.write(row + 1, col + 1, f'=COUNTIF(E{first_row + 1}:E{row},">={VFth}")+COUNTIF(E{first_row + 1}:E{row},"<=-{VFth}")',footer_cells_format)
            
        worksheet.conditional_format(f'E{first_row + 1}:E{row}', {'type': 'data_bar','bar_color':
            f'{data_bar_color}','bar_negative_color_same': True,'bar_no_border': True})
        row += 3
    if get_cn_write_xlsx_flag:
        create_vulnerability_factor_xlsx_file(eDN,None,VFth)
    NoVNR_max = 0
    CN_temp = None
    for node in eDN:
        if eDN[node]['Type'] != 'PI' and eDN[node]['Type'] != 'PI(PO)':
            eDN_temp = Replace_Gate(deepcopy(eDN),node,gates_delay_dictionary)
            if get_cn_write_xlsx_flag:
                create_vulnerability_factor_xlsx_file(eDN_temp,node,VFth)
            VF_List_temp = [node for node in eDN_temp if abs(eDN_temp[node]['VF_F']) >= VFth]
            NoVN_temp = len(VF_List_temp)
            if ((NoVN - NoVN_temp) > NoVNR_max):
                NoVNR_max = (NoVN - NoVN_temp)
                CN_temp = node
    if get_cn_write_xlsx_flag:
        workbook.close()
    return CN_temp

def real_algorithm(circuit_dictionary,VFth,NoR,benchmark_file_name,gates_delay_dictionary,real_algorithm_write_xlsx_flag = False,get_cn_write_xlsx_flag = False):
    # for excell file
    row = 0
    workbook = None
    worksheet = None
    if real_algorithm_write_xlsx_flag:
        newpath = f'Outputs/{benchmark_file_name}' 
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_real_algorithm.xlsx')
        worksheet = workbook.add_worksheet()
    def create_vulnerability_factor_xlsx_file(circuit_dictionary,iteration,VFth):
        # Create a workbook and add a worksheet.
        nonlocal worksheet
        nonlocal row

        # Add a bold format to use to highlight cells.
        header_cells_format = workbook.add_format()
        cells_format = workbook.add_format()
        key_gate_cells_format = workbook.add_format()
        vulnerability_cells_format = workbook.add_format()
        footer_cells_format = workbook.add_format()
        sheet_format = workbook.add_format()

        # Color variables
        border_color = '#ffffff'
        cell_bg_color = '#f2f2f2'
        header_bg_color = '#E6E6E6'
        font_color = '#616161'
        vulnerability_bg_color = '#FFCCCC'
        key_gate_color = '#CCECFF'
        data_bar_color = '#FFADAE'

        cells_format.set_border()
        cells_format.set_border_color(border_color)
        cells_format.set_bg_color(cell_bg_color)
        cells_format.set_font_color(font_color)
        cells_format.set_align('left')
        cells_format.set_font_size(12)
        
        footer_cells_format.set_border()
        footer_cells_format.set_border_color(border_color)
        footer_cells_format.set_bg_color(header_bg_color)
        footer_cells_format.set_font_color(font_color)
        footer_cells_format.set_align('left')
        footer_cells_format.set_font_size(12)
        footer_cells_format.set_bold()
        
        key_gate_cells_format.set_border()
        key_gate_cells_format.set_border_color(border_color)
        key_gate_cells_format.set_bg_color(key_gate_color)
        key_gate_cells_format.set_font_color(font_color)
        key_gate_cells_format.set_align('left')
        key_gate_cells_format.set_font_size(12)
        
        vulnerability_cells_format.set_border()
        vulnerability_cells_format.set_border_color(border_color)
        vulnerability_cells_format.set_bg_color(vulnerability_bg_color)
        vulnerability_cells_format.set_font_color(font_color)
        vulnerability_cells_format.set_align('left')
        vulnerability_cells_format.set_font_size(12)

        header_cells_format.set_align('center')
        header_cells_format.set_align('vcenter')
        header_cells_format.set_bg_color(header_bg_color)
        header_cells_format.set_font_size(14)
        header_cells_format.set_font_color(font_color)
        header_cells_format.set_bold()

        sheet_format.set_bg_color(cell_bg_color)

        
        
        headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)')
        worksheet.set_column('A1:XFD1048576',30,sheet_format)
        
        # {benchmark_file_name} Real Algorithm Vulnerability Factor
        if iteration == None:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Original Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        else:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Iteration = {iteration} Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        # Write some data headers.
        row += 1
        col = 0
        
        for header_data in headers_data:
            worksheet.set_row(row,20)
            worksheet.write(row,col,header_data,header_cells_format)
            col += 1

        # Start from the second cell below the headers
        row += 1
        col = 0
        first_row = row
        # Iterate over the data and write it out row by row.
        for node in circuit_dictionary:
            if 'K' in circuit_dictionary[node]['Type']:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],key_gate_cells_format)
                worksheet.write(row, col + 1, int(node), key_gate_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],key_gate_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],key_gate_cells_format)
            elif abs(circuit_dictionary[node]['VF_F']) >= VFth:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],vulnerability_cells_format)
                worksheet.write(row, col + 1, int(node), vulnerability_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],vulnerability_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],vulnerability_cells_format)
            else:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],cells_format)
                worksheet.write(row, col + 1, int(node), cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],cells_format)
            worksheet.write(row, col + 4, circuit_dictionary[node]['VF_F'],cells_format)
            row += 1
        worksheet.write(row, col, 'Sum Of VF :',footer_cells_format)
        worksheet.write(row, col + 1, f'=SUMIF(E{first_row + 1}:E{row},">0")+ABS(SUMIF(E{first_row + 1}:E{row},"<0"))',footer_cells_format)
        worksheet.write(row + 1, col, f'Number Of VF Above {VFth}',footer_cells_format)
        worksheet.write(row + 1, col + 1, f'=COUNTIF(E{first_row + 1}:E{row},">={VFth}")+COUNTIF(E{first_row + 1}:E{row},"<=-{VFth}")',footer_cells_format)
            
        worksheet.conditional_format(f'E{first_row + 1}:E{row}', {'type': 'data_bar','bar_color':
            f'{data_bar_color}','bar_negative_color_same': True,'bar_no_border': True})
        row += 3    
    eDN = circuit_dictionary
    if real_algorithm_write_xlsx_flag:
        create_vulnerability_factor_xlsx_file(eDN,None,VFth)
    for j in range(NoR):
        VF_list = [node for node in eDN if abs(eDN[node]['VF_F']) >= VFth]
        NoVN = len(VF_list)
        CN = get_CN(deepcopy(eDN),NoVN,VFth,benchmark_file_name,j + 1,gates_delay_dictionary,get_cn_write_xlsx_flag)
        if (NoVN != 0 and CN != None):
            eDN = Replace_Gate(deepcopy(eDN),CN,gates_delay_dictionary)
            if real_algorithm_write_xlsx_flag:
                create_vulnerability_factor_xlsx_file(eDN,j + 1,VFth)
        else:
            break
    if real_algorithm_write_xlsx_flag:
        workbook.close()
    return eDN

def real_algorithm_for_VFth_range(circuit_dictionary,NoR,benchmark_file_name,gates_delay_dictionary):

    newpath = f'Outputs/{benchmark_file_name}' 
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_real_algorithm_for_VFth_range.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Add a bold format to use to highlight cells.
    header_cells_format = workbook.add_format()
    cells_format = workbook.add_format()
    title_cells_format = workbook.add_format()
    sheet_format = workbook.add_format()

    row = 0

    border_color = '#ffffff'
    cell_bg_color = '#f2f2f2'
    header_bg_color = '#E6E6E6'
    font_color = '#616161'
    color1 = '#EE5C85'
    color2 = '#FFEB63'
    color3 = '#57ADE0'
    color4 = '#9CEA5B'
    color5 = '#FFAD63'
    color6 = '#8857E0'


    
    header_cells_format.set_align('vcenter')
    header_cells_format.set_bg_color(header_bg_color)
    header_cells_format.set_font_size(14)
    header_cells_format.set_font_color(font_color)
    header_cells_format.set_bold()

    cells_format.set_border()
    cells_format.set_border_color(border_color)
    cells_format.set_bg_color(cell_bg_color)
    cells_format.set_font_color(font_color)
    cells_format.set_align('left')
    cells_format.set_font_size(12)
        
    title_cells_format.set_border()
    title_cells_format.set_align('center')
    title_cells_format.set_border_color(border_color)
    title_cells_format.set_bg_color(header_bg_color)
    title_cells_format.set_font_color(font_color)
    title_cells_format.set_align('left')
    title_cells_format.set_font_size(12)
    title_cells_format.set_bold()

    sheet_format.set_bg_color(cell_bg_color)

    # Add the worksheet data that the charts will refer to.
    worksheet.set_column('A1:XFD1048576',30,sheet_format)
    worksheet.set_row(row,30)
    worksheet.merge_range('A1:R1', f'Compare {benchmark_file_name} Number And Sum Of Vulnerability Factor Between VFth_min = 0.5 And VFth_max = 0.9', header_cells_format)
       
    title_data = ('VFth','Circuit with VFth','Sum all absolut VFs in each VFth','Sum all Negative VFs in each VFth','Original circuit','Circuit with VFth = 0.5','Circuit with VFth = 0.6','Circuit with VFth = 0.7','Circuit with VFth = 0.8','Circuit with VFth = 0.9')
    VFth_list = np.arange(0.5,1,0.1)
    circuit_with_vfth = ['Original','VFth = 0.5','VFth = 0.6','VFth = 0.7','VFth = 0.8','VFth = 0.9']
    max_row = len(VFth_list) + 1
    data = []
    VF_sum_list = []
    VF_sum_list_n = []
    num_of_VFO_list = []
    num_of_VF5_list = []
    num_of_VF6_list = []
    num_of_VF7_list = []
    num_of_VF8_list = []
    num_of_VF9_list = []
    flag5 = True
    flag6 = True
    flag7 = True
    flag8 = True
    flag9 = True
    eDN = circuit_dictionary
    vf_list = [abs(eDN[node]['VF_F']) for node in eDN]
    vf_list_n = [abs(eDN[node]['VF_F']) for node in eDN if (eDN[node]['VF_F']) < 0]
    vf_list5 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.5]
    vf_list6 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.6]
    vf_list7 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.7]
    vf_list8 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.8]
    vf_list9 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.9]
    vf5_gt_c = len(vf_list5)
    vf6_gt_c = len(vf_list6)
    vf7_gt_c = len(vf_list7)
    vf8_gt_c = len(vf_list8)
    vf9_gt_c = len(vf_list9)
    num_of_VF9 = vf9_gt_c
    num_of_VF8 = vf8_gt_c - vf9_gt_c
    num_of_VF7 = vf7_gt_c - vf8_gt_c
    num_of_VF6 = vf6_gt_c - vf7_gt_c
    num_of_VF5 = vf5_gt_c - vf6_gt_c
    num_of_VFO_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
    vf_sum = sum(vf_list)
    vf_sum_n = sum(vf_list_n)
    VF_sum_list.append(vf_sum)
    VF_sum_list_n.append(vf_sum_n)

    for VFth in VFth_list:
        eDN = real_algorithm(deepcopy(circuit_dictionary),VFth,NoR,benchmark_file_name,deepcopy(gates_delay_dictionary))
        vf_list = [abs(eDN[node]['VF_F']) for node in eDN]
        vf_list_n = [abs(eDN[node]['VF_F']) for node in eDN if (eDN[node]['VF_F']) < 0]
        vf_list5 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.5]
        vf_list6 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.6]
        vf_list7 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.7]
        vf_list8 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.8]
        vf_list9 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.9]
        vf5_gt_c = len(vf_list5)
        vf6_gt_c = len(vf_list6)
        vf7_gt_c = len(vf_list7)
        vf8_gt_c = len(vf_list8)
        vf9_gt_c = len(vf_list9)
        num_of_VF9 = vf9_gt_c
        num_of_VF8 = vf8_gt_c - vf9_gt_c
        num_of_VF7 = vf7_gt_c - vf8_gt_c
        num_of_VF6 = vf6_gt_c - vf7_gt_c
        num_of_VF5 = vf5_gt_c - vf6_gt_c
        if flag5:
            num_of_VF5_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag5 = False
        elif flag6:
            num_of_VF6_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag6 = False
        elif flag7:
            num_of_VF7_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag7 = False
        elif flag8:
            num_of_VF8_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag8 = False
        elif flag9:
            num_of_VF9_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag9 = False
        vf_sum = sum(vf_list)
        vf_sum_n = sum(vf_list_n)
        VF_sum_list.append(vf_sum)
        VF_sum_list_n.append(vf_sum_n)
    VFth_list = np.append(VFth_list,1.0)
    data.append(VFth_list)
    data.append(circuit_with_vfth)
    data.append(VF_sum_list)
    data.append(VF_sum_list_n)
    data.append(num_of_VFO_list)
    data.append(num_of_VF5_list)
    data.append(num_of_VF6_list)
    data.append(num_of_VF7_list)
    data.append(num_of_VF8_list)
    data.append(num_of_VF9_list)

    worksheet.write_row('I2', title_data, title_cells_format)
    worksheet.write_column('I3', data[0],cells_format)
    worksheet.write_column('J3', data[1],cells_format)
    worksheet.write_column('K3', data[2],cells_format)
    worksheet.write_column('L3', data[3],cells_format)
    worksheet.write_column('M3', data[4],cells_format)
    worksheet.write_column('N3', data[5],cells_format)
    worksheet.write_column('O3', data[6],cells_format)
    worksheet.write_column('P3', data[7],cells_format)
    worksheet.write_column('Q3', data[8],cells_format)
    worksheet.write_column('R3', data[9],cells_format)
    

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart({'type': 'line'})
    chart2 = workbook.add_chart({'type': 'column'})

    # Configure the first series.
    chart1.add_series({
        'name':       '=Sheet1!$K$2',
        'categories': f'=Sheet1!$J$3:$J${max_row + 2}',
        'values':     f'=Sheet1!$K$3:$K${max_row + 2}',
        'line':   {'color': f'{color1}', 'width': 1.25},
        'smooth':     True,
    })
    chart1.add_series({
        'name':       '=Sheet1!$L$2',
        'categories': f'=Sheet1!$J$3:$J${max_row + 2}',
        'values':     f'=Sheet1!$L$3:$L${max_row + 2}',
        'line':   {'color': f'{color3}', 'width': 1.25},
        'smooth':     True,
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Sum of all VFs in each circuit with VFths','name_font':{'color':f'{font_color}'}})
    chart1.set_x_axis({'name': 'circuit with VFths','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'}})
    chart1.set_y_axis({'name': 'Result of Sum','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'},'major_gridlines':{'visible': True,'line':{'color':f'{border_color}','width': 1}}})
    chart1.set_legend({'font': {'color':f'{font_color}','bold': 1}})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_chartarea({'border': {'color': '#FFFFFF'},'fill':{'color': f'{cell_bg_color}'}})
    chart1.set_plotarea({'fill':   {'color': f'{cell_bg_color}'}})
    chart1.set_size({'width': 644, 'height': 370})


    # Configure the first series.
    chart2.add_series({
        'name':       '=Sheet1!$M$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$M$3:$M${max_row + 2}',
        'fill':   {'color': f'{color1}'}
    })
    chart2.add_series({
        'name':       '=Sheet1!$N$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$N$3:$N${max_row + 2}',
        'fill':   {'color': f'{color2}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$O$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$O$3:$O${max_row + 2}',
        'fill':   {'color': f'{color3}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$P$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$P$3:$P${max_row + 2}',
        'fill':   {'color': f'{color4}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$Q$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$Q$3:$Q${max_row + 2}',
        'fill':   {'color': f'{color5}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$R$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$R$3:$R${max_row + 2}',
        'fill':   {'color': f'{color6}'}
        
    })

    # Add a chart title and some axis labels.
    chart2.set_title ({'name': 'Number of nodes in each VFs','name_font':{'color':f'{font_color}'}})
    chart2.set_x_axis({'name': 'VFs','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'}})
    chart2.set_y_axis({'name': 'Number of node','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'},'major_gridlines':{'visible': True,'line':{'color':f'{border_color}','width': 1}}})
    chart2.set_legend({'font': {'color':f'{font_color}','bold': 1}})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart2.set_chartarea({'border': {'color': '#FFFFFF'},'fill':{'color': f'{cell_bg_color}'}})
    chart2.set_plotarea({'fill':   {'color': f'{cell_bg_color}'}})
    chart2.set_size({'width': 644, 'height': 370})

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1)
    worksheet.insert_chart('A2', chart2)

    workbook.close()

def my_get_CN(eDN,NoVN,VFth,benchmark_file_name,iteration,gates_delay_dictionary ,get_cn_write_xlsx_flag):
    # for excell file
    row = 0
    workbook = None
    worksheet = None
    if get_cn_write_xlsx_flag:
        newpath = f'Outputs/{benchmark_file_name}' 
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_my_get_CN_iteration_{iteration}.xlsx')
        worksheet = workbook.add_worksheet()
    def create_vulnerability_factor_xlsx_file(circuit_dictionary,iteration_node,VFth):
        # Create a workbook and add a worksheet.
        nonlocal worksheet
        nonlocal row

        # Add a bold format to use to highlight cells.
        header_cells_format = workbook.add_format()
        cells_format = workbook.add_format()
        key_gate_cells_format = workbook.add_format()
        vulnerability_cells_format = workbook.add_format()
        footer_cells_format = workbook.add_format()
        sheet_format = workbook.add_format()

        # Color variables
        border_color = '#ffffff'
        cell_bg_color = '#f2f2f2'
        header_bg_color = '#E6E6E6'
        font_color = '#616161'
        vulnerability_bg_color = '#FFCCCC'
        key_gate_color = '#CCECFF'
        data_bar_color = '#FFADAE'

        cells_format.set_border()
        cells_format.set_border_color(border_color)
        cells_format.set_bg_color(cell_bg_color)
        cells_format.set_font_color(font_color)
        cells_format.set_align('left')
        cells_format.set_font_size(12)
        
        footer_cells_format.set_border()
        footer_cells_format.set_border_color(border_color)
        footer_cells_format.set_bg_color(header_bg_color)
        footer_cells_format.set_font_color(font_color)
        footer_cells_format.set_align('left')
        footer_cells_format.set_font_size(12)
        footer_cells_format.set_bold()
        
        key_gate_cells_format.set_border()
        key_gate_cells_format.set_border_color(border_color)
        key_gate_cells_format.set_bg_color(key_gate_color)
        key_gate_cells_format.set_font_color(font_color)
        key_gate_cells_format.set_align('left')
        key_gate_cells_format.set_font_size(12)
        
        vulnerability_cells_format.set_border()
        vulnerability_cells_format.set_border_color(border_color)
        vulnerability_cells_format.set_bg_color(vulnerability_bg_color)
        vulnerability_cells_format.set_font_color(font_color)
        vulnerability_cells_format.set_align('left')
        vulnerability_cells_format.set_font_size(12)

        header_cells_format.set_align('center')
        header_cells_format.set_align('vcenter')
        header_cells_format.set_bg_color(header_bg_color)
        header_cells_format.set_font_size(14)
        header_cells_format.set_font_color(font_color)
        header_cells_format.set_bold()

        sheet_format.set_bg_color(cell_bg_color)

        
        
        headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)')
        worksheet.set_column('A1:XFD1048576',30,sheet_format)
        
        # {benchmark_file_name} Real Algorithm Vulnerability Factor
        if iteration_node == None:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Original Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        else:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Node = {iteration_node} Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        # Write some data headers.
        row += 1
        col = 0
        
        for header_data in headers_data:
            worksheet.set_row(row,20)
            worksheet.write(row,col,header_data,header_cells_format)
            col += 1

        # Start from the second cell below the headers
        row += 1
        col = 0
        first_row = row
        # Iterate over the data and write it out row by row.
        for node in circuit_dictionary:
            if 'K' in circuit_dictionary[node]['Type']:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],key_gate_cells_format)
                worksheet.write(row, col + 1, int(node), key_gate_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],key_gate_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],key_gate_cells_format)
            elif abs(circuit_dictionary[node]['VF_F']) >= VFth:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],vulnerability_cells_format)
                worksheet.write(row, col + 1, int(node), vulnerability_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],vulnerability_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],vulnerability_cells_format)
            else:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],cells_format)
                worksheet.write(row, col + 1, int(node), cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],cells_format)
            worksheet.write(row, col + 4, circuit_dictionary[node]['VF_F'],cells_format)
            row += 1
        worksheet.write(row, col, 'Sum Of VF :',footer_cells_format)
        worksheet.write(row, col + 1, f'=SUMIF(E{first_row + 1}:E{row},">0")+ABS(SUMIF(E{first_row + 1}:E{row},"<0"))',footer_cells_format)
        worksheet.write(row + 1, col, f'Number Of VF Above {VFth}',footer_cells_format)
        worksheet.write(row + 1, col + 1, f'=COUNTIF(E{first_row + 1}:E{row},">={VFth}")+COUNTIF(E{first_row + 1}:E{row},"<=-{VFth}")',footer_cells_format)
            
        worksheet.conditional_format(f'E{first_row + 1}:E{row}', {'type': 'data_bar','bar_color':
            f'{data_bar_color}','bar_negative_color_same': True,'bar_no_border': True})
        row += 3
    if get_cn_write_xlsx_flag:
        create_vulnerability_factor_xlsx_file(eDN,None,VFth)
    NoVNR_max = 0
    CN_temp = None
    for node in eDN:
        if eDN[node]['Type'] != 'PI' and eDN[node]['Type'] != 'PI(PO)' and abs(eDN[node]['VF_F']) >= 0.5:
            eDN_temp = Replace_Gate(deepcopy(eDN),node,gates_delay_dictionary)
            if get_cn_write_xlsx_flag:
                create_vulnerability_factor_xlsx_file(eDN_temp,node,VFth)
            VF_List_temp = [node for node in eDN_temp if abs(eDN_temp[node]['VF_F']) >= VFth]
            NoVN_temp = len(VF_List_temp)
            if ((NoVN - NoVN_temp) > NoVNR_max):
                NoVNR_max = (NoVN - NoVN_temp)
                CN_temp = node
    if get_cn_write_xlsx_flag:
        workbook.close()
    return CN_temp

def my_real_algorithm(circuit_dictionary,VFth,NoR,benchmark_file_name,gates_delay_dictionary,real_algorithm_write_xlsx_flag = False,get_cn_write_xlsx_flag = False):
    # for excell file
    row = 0
    workbook = None
    worksheet = None
    if real_algorithm_write_xlsx_flag:
        newpath = f'Outputs/{benchmark_file_name}' 
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_my_real_algorithm.xlsx')
        worksheet = workbook.add_worksheet()
    def create_vulnerability_factor_xlsx_file(circuit_dictionary,iteration,VFth):
        # Create a workbook and add a worksheet.
        nonlocal worksheet
        nonlocal row

        # Add a bold format to use to highlight cells.
        header_cells_format = workbook.add_format()
        cells_format = workbook.add_format()
        key_gate_cells_format = workbook.add_format()
        vulnerability_cells_format = workbook.add_format()
        footer_cells_format = workbook.add_format()
        sheet_format = workbook.add_format()

        # Color variables
        border_color = '#ffffff'
        cell_bg_color = '#f2f2f2'
        header_bg_color = '#E6E6E6'
        font_color = '#616161'
        vulnerability_bg_color = '#FFCCCC'
        key_gate_color = '#CCECFF'
        data_bar_color = '#FFADAE'

        cells_format.set_border()
        cells_format.set_border_color(border_color)
        cells_format.set_bg_color(cell_bg_color)
        cells_format.set_font_color(font_color)
        cells_format.set_align('left')
        cells_format.set_font_size(12)
        
        footer_cells_format.set_border()
        footer_cells_format.set_border_color(border_color)
        footer_cells_format.set_bg_color(header_bg_color)
        footer_cells_format.set_font_color(font_color)
        footer_cells_format.set_align('left')
        footer_cells_format.set_font_size(12)
        footer_cells_format.set_bold()
        
        key_gate_cells_format.set_border()
        key_gate_cells_format.set_border_color(border_color)
        key_gate_cells_format.set_bg_color(key_gate_color)
        key_gate_cells_format.set_font_color(font_color)
        key_gate_cells_format.set_align('left')
        key_gate_cells_format.set_font_size(12)
        
        vulnerability_cells_format.set_border()
        vulnerability_cells_format.set_border_color(border_color)
        vulnerability_cells_format.set_bg_color(vulnerability_bg_color)
        vulnerability_cells_format.set_font_color(font_color)
        vulnerability_cells_format.set_align('left')
        vulnerability_cells_format.set_font_size(12)

        header_cells_format.set_align('center')
        header_cells_format.set_align('vcenter')
        header_cells_format.set_bg_color(header_bg_color)
        header_cells_format.set_font_size(14)
        header_cells_format.set_font_color(font_color)
        header_cells_format.set_bold()

        sheet_format.set_bg_color(cell_bg_color)

        headers_data = ('Types','Nodes','Zero Probabilities (P0)','One Probabilities (P1)','Vulnerability Factors (VF)')
        worksheet.set_column('A1:XFD1048576',30,sheet_format)
        
        
        if iteration == None:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Original Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        else:
            worksheet.set_row(row,30)
            worksheet.merge_range(f'A{row + 1}:E{row + 1}', f'Iteration = {iteration} Benchmark {benchmark_file_name} Real Algorithm Vulnerability Factor For VFth = {VFth}', header_cells_format)
        # Write some data headers.
        row += 1
        col = 0
        
        for header_data in headers_data:
            worksheet.set_row(row,20)
            worksheet.write(row,col,header_data,header_cells_format)
            col += 1

        # Start from the second cell below the headers
        row += 1
        col = 0
        first_row = row
        # Iterate over the data and write it out row by row.
        for node in circuit_dictionary:
            if 'K' in circuit_dictionary[node]['Type']:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],key_gate_cells_format)
                worksheet.write(row, col + 1, int(node), key_gate_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],key_gate_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],key_gate_cells_format)
            elif abs(circuit_dictionary[node]['VF_F']) >= VFth:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],vulnerability_cells_format)
                worksheet.write(row, col + 1, int(node), vulnerability_cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],vulnerability_cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],vulnerability_cells_format)
            else:
                worksheet.write(row, col, circuit_dictionary[node]['Type'],cells_format)
                worksheet.write(row, col + 1, int(node), cells_format)
                worksheet.write(row, col + 2, circuit_dictionary[node]['P0_F'],cells_format)
                worksheet.write(row, col + 3, circuit_dictionary[node]['P1_F'],cells_format)
            worksheet.write(row, col + 4, circuit_dictionary[node]['VF_F'],cells_format)
            row += 1
        worksheet.write(row, col, 'Sum Of VF :',footer_cells_format)
        worksheet.write(row, col + 1, f'=SUMIF(E{first_row + 1}:E{row},">0")+ABS(SUMIF(E{first_row + 1}:E{row},"<0"))',footer_cells_format)
        worksheet.write(row + 1, col, f'Number Of VF Above {VFth}',footer_cells_format)
        worksheet.write(row + 1, col + 1, f'=COUNTIF(E{first_row + 1}:E{row},">={VFth}")+COUNTIF(E{first_row + 1}:E{row},"<=-{VFth}")',footer_cells_format)
            
        worksheet.conditional_format(f'E{first_row + 1}:E{row}', {'type': 'data_bar','bar_color':
            f'{data_bar_color}','bar_negative_color_same': True,'bar_no_border': True})
        row += 3    
    eDN = circuit_dictionary
    if real_algorithm_write_xlsx_flag:
        create_vulnerability_factor_xlsx_file(eDN,None,VFth)
    for j in range(NoR):
        VF_list = [node for node in eDN if abs(eDN[node]['VF_F']) >= VFth]
        NoVN = len(VF_list)
        CN = my_get_CN(deepcopy(eDN),NoVN,VFth,benchmark_file_name,j + 1,gates_delay_dictionary,get_cn_write_xlsx_flag)
        if (NoVN != 0 and CN != None):
            eDN = Replace_Gate(deepcopy(eDN),CN,gates_delay_dictionary)
            if real_algorithm_write_xlsx_flag:
                create_vulnerability_factor_xlsx_file(eDN,j + 1,VFth)
        else:
            break
    if real_algorithm_write_xlsx_flag:
        workbook.close()
    return eDN

def my_real_algorithm_for_VFth_range(circuit_dictionary,NoR,benchmark_file_name,gates_delay_dictionary):

    newpath = f'Outputs/{benchmark_file_name}' 
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    workbook = xlsxwriter.Workbook(f'Outputs/{benchmark_file_name}/{benchmark_file_name}_my_real_algorithm_for_VFth_range.xlsx')
    worksheet = workbook.add_worksheet()
    
    # Add a bold format to use to highlight cells.
    header_cells_format = workbook.add_format()
    cells_format = workbook.add_format()
    title_cells_format = workbook.add_format()
    sheet_format = workbook.add_format()

    row = 0

    border_color = '#ffffff'
    cell_bg_color = '#f2f2f2'
    header_bg_color = '#E6E6E6'
    font_color = '#616161'
    color1 = '#EE5C85'
    color2 = '#FFEB63'
    color3 = '#57ADE0'
    color4 = '#9CEA5B'
    color5 = '#FFAD63'
    color6 = '#8857E0'


    
    header_cells_format.set_align('vcenter')
    header_cells_format.set_bg_color(header_bg_color)
    header_cells_format.set_font_size(14)
    header_cells_format.set_font_color(font_color)
    header_cells_format.set_bold()

    cells_format.set_border()
    cells_format.set_border_color(border_color)
    cells_format.set_bg_color(cell_bg_color)
    cells_format.set_font_color(font_color)
    cells_format.set_align('left')
    cells_format.set_font_size(12)
        
    title_cells_format.set_border()
    title_cells_format.set_align('center')
    title_cells_format.set_border_color(border_color)
    title_cells_format.set_bg_color(header_bg_color)
    title_cells_format.set_font_color(font_color)
    title_cells_format.set_align('left')
    title_cells_format.set_font_size(12)
    title_cells_format.set_bold()

    sheet_format.set_bg_color(cell_bg_color)

    # Add the worksheet data that the charts will refer to.
    worksheet.set_column('A1:XFD1048576',30,sheet_format)
    worksheet.set_row(row,30)
    worksheet.merge_range('A1:R1', f'Compare {benchmark_file_name} Number And Sum Of Vulnerability Factor Between VFth_min = 0.5 And VFth_max = 0.9', header_cells_format)
       
    title_data = ('VFth','Circuit with VFth','Sum all absolut VFs in each VFth','Sum all Negative VFs in each VFth','Original circuit','Circuit with VFth = 0.5','Circuit with VFth = 0.6','Circuit with VFth = 0.7','Circuit with VFth = 0.8','Circuit with VFth = 0.9')
    VFth_list = np.arange(0.5,1,0.1)
    circuit_with_vfth = ['Original','VFth = 0.5','VFth = 0.6','VFth = 0.7','VFth = 0.8','VFth = 0.9']
    max_row = len(VFth_list) + 1
    data = []
    VF_sum_list = []
    VF_sum_list_n = []
    num_of_VFO_list = []
    num_of_VF5_list = []
    num_of_VF6_list = []
    num_of_VF7_list = []
    num_of_VF8_list = []
    num_of_VF9_list = []
    flag5 = True
    flag6 = True
    flag7 = True
    flag8 = True
    flag9 = True
    eDN = circuit_dictionary
    vf_list = [abs(eDN[node]['VF_F']) for node in eDN]
    vf_list_n = [abs(eDN[node]['VF_F']) for node in eDN if (eDN[node]['VF_F']) < 0]
    vf_list5 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.5]
    vf_list6 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.6]
    vf_list7 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.7]
    vf_list8 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.8]
    vf_list9 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.9]
    vf5_gt_c = len(vf_list5)
    vf6_gt_c = len(vf_list6)
    vf7_gt_c = len(vf_list7)
    vf8_gt_c = len(vf_list8)
    vf9_gt_c = len(vf_list9)
    num_of_VF9 = vf9_gt_c
    num_of_VF8 = vf8_gt_c - vf9_gt_c
    num_of_VF7 = vf7_gt_c - vf8_gt_c
    num_of_VF6 = vf6_gt_c - vf7_gt_c
    num_of_VF5 = vf5_gt_c - vf6_gt_c
    num_of_VFO_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
    vf_sum = sum(vf_list)
    vf_sum_n = sum(vf_list_n)
    VF_sum_list.append(vf_sum)
    VF_sum_list_n.append(vf_sum_n)

    for VFth in VFth_list:
        eDN = my_real_algorithm(deepcopy(circuit_dictionary),VFth,NoR,benchmark_file_name,deepcopy(gates_delay_dictionary))
        vf_list = [abs(eDN[node]['VF_F']) for node in eDN]
        vf_list_n = [abs(eDN[node]['VF_F']) for node in eDN if (eDN[node]['VF_F']) < 0]
        vf_list5 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.5]
        vf_list6 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.6]
        vf_list7 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.7]
        vf_list8 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.8]
        vf_list9 = [abs(eDN[node]['VF_F']) for node in eDN if abs(eDN[node]['VF_F']) >= 0.9]
        vf5_gt_c = len(vf_list5)
        vf6_gt_c = len(vf_list6)
        vf7_gt_c = len(vf_list7)
        vf8_gt_c = len(vf_list8)
        vf9_gt_c = len(vf_list9)
        num_of_VF9 = vf9_gt_c
        num_of_VF8 = vf8_gt_c - vf9_gt_c
        num_of_VF7 = vf7_gt_c - vf8_gt_c
        num_of_VF6 = vf6_gt_c - vf7_gt_c
        num_of_VF5 = vf5_gt_c - vf6_gt_c
        if flag5:
            num_of_VF5_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag5 = False
        elif flag6:
            num_of_VF6_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag6 = False
        elif flag7:
            num_of_VF7_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag7 = False
        elif flag8:
            num_of_VF8_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag8 = False
        elif flag9:
            num_of_VF9_list.extend([num_of_VF5,num_of_VF6,num_of_VF7,num_of_VF8,num_of_VF9])
            flag9 = False
        vf_sum = sum(vf_list)
        vf_sum_n = sum(vf_list_n)
        VF_sum_list.append(vf_sum)
        VF_sum_list_n.append(vf_sum_n)
    VFth_list = np.append(VFth_list,1.0)
    data.append(VFth_list)
    data.append(circuit_with_vfth)
    data.append(VF_sum_list)
    data.append(VF_sum_list_n)
    data.append(num_of_VFO_list)
    data.append(num_of_VF5_list)
    data.append(num_of_VF6_list)
    data.append(num_of_VF7_list)
    data.append(num_of_VF8_list)
    data.append(num_of_VF9_list)

    worksheet.write_row('I2', title_data, title_cells_format)
    worksheet.write_column('I3', data[0],cells_format)
    worksheet.write_column('J3', data[1],cells_format)
    worksheet.write_column('K3', data[2],cells_format)
    worksheet.write_column('L3', data[3],cells_format)
    worksheet.write_column('M3', data[4],cells_format)
    worksheet.write_column('N3', data[5],cells_format)
    worksheet.write_column('O3', data[6],cells_format)
    worksheet.write_column('P3', data[7],cells_format)
    worksheet.write_column('Q3', data[8],cells_format)
    worksheet.write_column('R3', data[9],cells_format)
    

    # Create a new chart object. In this case an embedded chart.
    chart1 = workbook.add_chart({'type': 'line'})
    chart2 = workbook.add_chart({'type': 'column'})

    # Configure the first series.
    chart1.add_series({
        'name':       '=Sheet1!$K$2',
        'categories': f'=Sheet1!$J$3:$J${max_row + 2}',
        'values':     f'=Sheet1!$K$3:$K${max_row + 2}',
        'line':   {'color': f'{color1}', 'width': 1.25},
        'smooth':     True,
    })
    chart1.add_series({
        'name':       '=Sheet1!$L$2',
        'categories': f'=Sheet1!$J$3:$J${max_row + 2}',
        'values':     f'=Sheet1!$L$3:$L${max_row + 2}',
        'line':   {'color': f'{color3}', 'width': 1.25},
        'smooth':     True,
    })
    
    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Sum of all VFs in each circuit with VFths','name_font':{'color':f'{font_color}'}})
    chart1.set_x_axis({'name': 'circuit with VFths','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'}})
    chart1.set_y_axis({'name': 'Result of Sum','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'},'major_gridlines':{'visible': True,'line':{'color':f'{border_color}','width': 1}}})
    chart1.set_legend({'font': {'color':f'{font_color}','bold': 1}})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_chartarea({'border': {'color': '#FFFFFF'},'fill':{'color': f'{cell_bg_color}'}})
    chart1.set_plotarea({'fill':   {'color': f'{cell_bg_color}'}})
    chart1.set_size({'width': 644, 'height': 370})


    # Configure the first series.
    chart2.add_series({
        'name':       '=Sheet1!$M$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$M$3:$M${max_row + 2}',
        'fill':   {'color': f'{color1}'}
    })
    chart2.add_series({
        'name':       '=Sheet1!$N$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$N$3:$N${max_row + 2}',
        'fill':   {'color': f'{color2}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$O$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$O$3:$O${max_row + 2}',
        'fill':   {'color': f'{color3}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$P$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$P$3:$P${max_row + 2}',
        'fill':   {'color': f'{color4}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$Q$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$Q$3:$Q${max_row + 2}',
        'fill':   {'color': f'{color5}'}
        
    })
    chart2.add_series({
        'name':       '=Sheet1!$R$2',
        'categories': f'=Sheet1!$I$3:$I${max_row + 2}',
        'values':     f'=Sheet1!$R$3:$R${max_row + 2}',
        'fill':   {'color': f'{color6}'}
        
    })

    # Add a chart title and some axis labels.
    chart2.set_title ({'name': 'Number of nodes in each VFs','name_font':{'color':f'{font_color}'}})
    chart2.set_x_axis({'name': 'VFs','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'}})
    chart2.set_y_axis({'name': 'Number of node','name_font':{'color':f'{font_color}'}, 'num_font':{'color':f'{font_color}'},'line':{'color':f'{font_color}'},'major_gridlines':{'visible': True,'line':{'color':f'{border_color}','width': 1}}})
    chart2.set_legend({'font': {'color':f'{font_color}','bold': 1}})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart2.set_chartarea({'border': {'color': '#FFFFFF'},'fill':{'color': f'{cell_bg_color}'}})
    chart2.set_plotarea({'fill':   {'color': f'{cell_bg_color}'}})
    chart2.set_size({'width': 644, 'height': 370})

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D2', chart1)
    worksheet.insert_chart('A2', chart2)

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

if __name__ == '__main__': 
    main()