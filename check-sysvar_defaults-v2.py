#!/usr/bin/env python

import openpyxl
import re
import linecache

def get_sysvar_info_from_table(sysvar_table_file_path):

    # Reads the prepared excel table and gets the complete information of the system variable
    sysvar_list = []
    wb = openpyxl.load_workbook(sysvar_table_file_path)
    ws = wb.active

    sysvar_from_table = list(ws.values)
    del sysvar_from_table[0]

    for i in sysvar_from_table:
        new_dict = {'VARIABLE_NAME': i[0],
                    'VARIABLE_SCOPE': i[1],
                    'DEFAULT_VALUE': str(i[2]),
                    'CURRENT_VALUE': str(i[3]),
                    'VALUE_RANGE': '[{}, {}]'.format(str(i[4]), str(i[5])),
                    'POSSIBLE_VALUES': str(i[6]),
                    'IS_NOOP': i[7]}
        if new_dict['IS_NOOP'] == 'NO':
            sysvar_list.append(new_dict)
        else:
            continue

    return sysvar_list


def get_sysvar_info_from_doc(sysvar_doc_file_path):

    # Reads the doc sysvar doc file and collects the doc information of sysvars
    POS = []
    lineNum = 0
    docs_sysvar_info_sum = []
    docs_sysvar_info_sum_dedup = []

    with open(sysvar_doc_file_path,'r',encoding='utf-8') as fp:
        lineNum = 0
        for line in fp:
            lineNum += 1
            if re.search('###', line):
                POS.append(lineNum)
            else:
                continue

    for i in POS:
        if (POS.index(i)+1) < len(POS):
            docs_sysvar_info = {}
            for line in range(i, POS[POS.index(i)+1]-1):
                line_content = linecache.getline(sysvar_doc_file_path, line)
                if re.search(r'###', line_content):
                    doc_sysvar_name = re.sub(r'(###\s)(`)?([`\w]*)(\s<span.*)?$', r'\3', line_content).strip()
                    docs_sysvar_info['variable_name'] = doc_sysvar_name
                elif re.search(r'- Default value:', line_content):
                    doc_sysvar_default_value = re.sub(r'(- Default value\: )(.+)$', r'\2', line_content).strip()
                    docs_sysvar_info['default_value'] = doc_sysvar_default_value.strip('`')
                elif re.search(r'- Scope:', line_content):
                    docs_sysvar_scope = re.sub(r'(- Scope: )(.+)$', r'\2', line_content).strip()
                    docs_sysvar_info['scope'] = docs_sysvar_scope.replace(' | ', ',')
                elif re.search(r'(- Range: )(`?)\[(.+),(\s)?(.+)\](`?)', line_content):
                    docs_sysvar_minv = re.sub(r'(- Range: )(`?)(\[)(.+)(,)(\s?)(.+)(\])(`?)', r'\4', line_content).strip()
                    docs_sysvar_maxv = re.sub(r'(- Range: )(`?)(\[)(.+)(,)(\s?)(.+)(\])(`?)', r'\7', line_content).strip()
                    docs_sysvar_info['range'] = '[' + docs_sysvar_minv + ', ' + docs_sysvar_maxv + ']'
                else:
                    continue
                docs_sysvar_info_sum.append(docs_sysvar_info)

        else:
            pass

    for li in docs_sysvar_info_sum:
        if li not in docs_sysvar_info_sum_dedup:
            docs_sysvar_info_sum_dedup.append(li)
        else:
            continue
    return docs_sysvar_info_sum_dedup

def main():

    for docs_dict_item in get_sysvar_info_from_doc(sysvar_doc_file_path):
        for table_dict_item in get_sysvar_info_from_table(sysvar_table_file_path):
            if docs_dict_item['variable_name'] == table_dict_item['VARIABLE_NAME']:
                if 'default_value' in docs_dict_item and 'DEFAULT_VALUE' in table_dict_item:
                    if docs_dict_item['default_value'] == table_dict_item['DEFAULT_VALUE']:
                        continue
                    else:
                        print('Default value error' + '\t\t\t' + docs_dict_item['variable_name'] + '\t\t\t'
                            + docs_dict_item['default_value'] + '\t\t\t' + table_dict_item['DEFAULT_VALUE']
                            )
                else:
                    continue

                if 'scope' in docs_dict_item and 'VARIABLE_SCOPE' in table_dict_item:
                    if docs_dict_item['scope'] == table_dict_item['VARIABLE_SCOPE']:
                        continue
                    else:
                        print('Scope error' + '\t\t\t' + docs_dict_item['variable_name'] + '\t\t\t'
                            + docs_dict_item['scope'] + '\t\t\t' + table_dict_item['VARIABLE_SCOPE']
                            )
                else:
                    continue

                if 'range' in docs_dict_item and 'VALUE_RANGE' in table_dict_item:
                    if docs_dict_item['range'] == table_dict_item['VALUE_RANGE']:
                        continue
                    else:
                        print('Range error' + '\t\t\t' + docs_dict_item['variable_name'] + '\t\t\t'
                            + docs_dict_item['range'] + '\t\t\t' + table_dict_item['VALUE_RANGE']
                            )
                else:
                    continue

            else:
                continue

    return


if __name__ == '__main__':

    sysvar_table_file_path = '/Users/shawntom/Downloads/sys-var-640-1102.xlsx'
    sysvar_doc_file_path = '/Users/shawntom/Documents/GitHub/upstream/docs/system-variables.md'

    main()














