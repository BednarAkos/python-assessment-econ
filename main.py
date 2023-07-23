import argparse
import pptmaker


parser = argparse.ArgumentParser(description='Commands for the CLI app')
"""Commands usable in the command line"""
parser.add_argument('-c', '--config_path', action='store', help="shows path", type=str, nargs='?', const='')  # Take a string as path
parser.add_argument('-e', '--exit', action='store_true', help="exit without doing anything")  # Just exits
# -h as "help" command gives you the command options

all_arguments = parser.parse_args()  #parsing the arguments
data = []  # list for handling the json file

print(all_arguments)

if all_arguments.config_path == "":
    print("need")
else:
    data = pptmaker.loader(all_arguments.config_path)  # load the json file as a list of dictionary
    ppt = pptmaker.ppt_processor(data)  # this assembles the ppt
    ppt.save('test.pptx')  # create output file

if all_arguments.exit:
    print("Thank you for playing")
