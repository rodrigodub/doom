#################################################
# Doom
# Script to adjust read an RFE email
#
# v0.04
# # for ticket #1
#
# Rodrigo Nobrega
# 20180514-20180516
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
        self.contents = self.readcontents(file)

    def readcontents(self, file):
        a = ''
        b = open(r'{}{}'.format(DIRECTORY, file), 'r', encoding = 'ISO-8859-1')
        # b = open(r'{}{}'.format(DIRECTORY, file), 'r', encoding='ANSI')
        for i in b:
            a = a + r'{}\n'.format(i)
        b.close()
        return a

    def outputcustomer(self):
        for i in self.contents:
            if 'Customer Name' in i:
                print(i)


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('                     Reads and processes contents of RFE')
    print('=============================================================================\n')
    rfe = ReadRFE('SI-81772.msg')
    print(rfe.contents)
    # rfe.outputcustomer()


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
