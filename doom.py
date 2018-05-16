#################################################
# Doom
# Script to adjust read an RFE email
#
# v0.06
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
OUTPUTRFE = 'Customer Name:\n{}\n\n' \
            'Site:\n\n\n' \
            'Who has identified the requirement?:\n\n\n' \
            'Role:\n\n\n' \
            'acQuire RFE Owner:\n\n\n' \
            'Original SI#:\n\n\n' \
            'GIM Suite Version:\n\n\n' \
            'Description:\n\n\n' \
            'Workaround:\n\n\n' \
            'Additional Comments:\n\n\n'


# define Timeshift class
class ReadRFE(object):
    """
    Reads contents of an RFE Outlook message file
    """
    def __init__(self, file):
        self.contents = self.readcontents(file)
        self.customer = self.outputcustomer()


    def readcontents(self, file):
        a = ''
        b = open(r'{}{}'.format(DIRECTORY, file), 'r', encoding='ISO-8859-1')
        for i in b:
            a = a + '{}'.format(i.replace('\x00', ''))
        b.close()
        return a

    def outputcustomer(self):
        return self.contents.split('Customer Name')[1].split('Site')[0].replace('\n', '').strip()


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('                     Reads and processes contents of RFE')
    print('=============================================================================\n')
    rfe = ReadRFE('SI-81772.msg')
    print(OUTPUTRFE.format(rfe.customer))


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
