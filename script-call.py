# input path: /home/raihan/Downloads/Input_data_newJson
# output path: /home/raihan/PycharmProjects/BadhanDOC
import argparse
import sys
import os

parser = argparse.ArgumentParser(description="Call different scripts. The call should be in this order '-d -x1 -x2 -p'")
parser.add_argument('-d', '--lzac_doc', type=str, metavar='', required=False, help="Call lzac_doc.py: value must be 'doc'", nargs='+')
parser.add_argument('-x1', '--xlsx1', type=str, metavar='', required=False, help="Call xlsx_part1.py: value must be 'xl1'", nargs='+')
parser.add_argument('-x2', '--xlsx2', type=str, metavar='', required=False, help="Call xlsx_part2.py: value must be 'xl2'", nargs='+')
parser.add_argument('-p', '--path', type=str, metavar='', required=False, help='Enter path', nargs='+')
args = parser.parse_args()
temp_args = []
new_temp_args = []
for val in sys.argv:
    temp_args.append(val)
    new_temp_args.append(val)

if args.lzac_doc != None and args.lzac_doc[0] == 'doc':
    if len(temp_args) == 3:
        temp_args.pop()
        temp_args.pop()
    if len(temp_args) > 3:
        while len(temp_args) > 3:
            temp_args.pop()
        del temp_args[1]
        del temp_args[1]
        del new_temp_args[1]
        del new_temp_args[1]
    if os.path.isdir(new_temp_args[-2]):
        temp_args.append(new_temp_args[-2])
        temp_args.append(new_temp_args[-1])
    sys.argv = temp_args
    import lzak_docs

temp_args = []
for val in new_temp_args:
    temp_args.append(val)

if args.xlsx1 != None and args.xlsx1[0] == 'xl1':
    if len(temp_args) == 3:
        temp_args.pop()
        temp_args.pop()
    if len(temp_args) > 3:
        while len(temp_args) > 3:
            temp_args.pop()
        del temp_args[1]
        del temp_args[1]
        del new_temp_args[1]
        del new_temp_args[1]
    if os.path.isdir(new_temp_args[-2]):
        temp_args.append(new_temp_args[-2])
        temp_args.append(new_temp_args[-1])
    sys.argv = temp_args
    import xlsx_part1

temp_args = []
for val in new_temp_args:
    temp_args.append(val)

if args.xlsx2 != None and args.xlsx2[0] == 'xl2':
    if len(temp_args) == 3:
        temp_args.pop()
        temp_args.pop()
    if len(temp_args) > 3:
        while len(temp_args) > 3:
            temp_args.pop()
        del temp_args[1]
        del temp_args[1]
        del new_temp_args[1]
        del new_temp_args[1]
    if os.path.isdir(new_temp_args[-2]):
        temp_args.append(new_temp_args[-2])
        temp_args.append(new_temp_args[-1])
    sys.argv = temp_args
    import xlsx_part2

    # try:
    #     # python 3.4+ should use builtin unittest.mock not mock package
    #     from unittest.mock import patch
    # except ImportError:
    #     from mock import patch
    #
    # def get_setup_file():
    #     parserN = argparse.ArgumentParser()
    #     parserN.add_argument('-f')
    #     argsN = parserN.parse_args()
    #     print(argsN)
    #     return argsN.path
    #
    # def test_parse_args():
    #     testargs = ["path", "-p", "/home/fenton/project/setup.py"]
    #     with patch.object(sys, 'argv', testargs):
    #         setup = get_setup_file()
    #         assert setup == "/home/fenton/project/setup.py"
    #
    # if __name__ == '__main__':
    #     test_parse_args()
