#################################################
# Doom
# Script to adjust read an RFE email
#
# v0.02
# # for ticket #1
#
# Rodrigo Nobrega
# 20180514-
#################################################

# import modules
# import datetime
# import os


# global variables
DIRECTORY = './examples/'


# define Timeshift class
class ReadRFE(object):
    """
    Reads contents of an RFE Outlook message file
    """
    def __init__(self, file):
        self.contents = open('{}{}'.format(DIRECTORY, file), 'r', encoding = 'ISO-8859-1')

    def output(self):
        for i in self.contents:
            print(i)


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('                     Reads and processes contents of RFE')
    print('=============================================================================\n')
    rfe = ReadRFE('SI-81772.msg')
    rfe.output()


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
