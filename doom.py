#################################################
# Doom
# Script to adjust read an RFE email
#
# v0.10
# for tickets #4, #5
#
# Rodrigo Nobrega
# 20180514-20180517
#################################################

# import modules
# import datetime
# import os


# global variables
# DIRECTORY = './examples/'
DIRECTORY = 'C:/Users/rnobrega/Desktop/'
OUTPUTRFE = 'Customer Name:\n{}\n\n' \
            'Site:\n{}\n\n' \
            'Who has identified the requirement?:\n{}\n\n' \
            'Role:\n{}\n\n' \
            'acQuire RFE Owner:\n{}\n\n' \
            'Original SI#:\n{}\n\n' \
            'GIM Suite Version:\n{}\n\n' \
            'Description:\n{}\n{}\n\n' \
            'Workaround:\n{}\n\n' \
            'Additional Comments:\n{}\n\n'


# define Timeshift class
class ReadRFE(object):
    """
    Reads contents of an RFE Outlook message file
    """
    def __init__(self, file):
        self.contents = self.readcontents(file)
        self.customer = self.returnstring('Customer Name', 'Site')
        self.site = self.returnstring('Site', 'Who has identified')
        self.user = self.returnstring('Who has identified the requirement?', 'Role')
        self.role = self.returnstring('Role', 'GIM Suite Version')
        self.version = self.returnstring('GIM Suite Version', 'Triage Details')
        self.owner = self.returnstring('acQuire RFE Owner', 'SI#')
        self.si = self.returnstring('SI#', 'RFE Summary')
        self.userstory = self.returnstring('User Story', 'Description')
        self.description = self.returnstring('Description', 'Problem Type')
        self.workaround = self.returnstring('Workaround(s)\n\n\n\n \n\n\n\n[Implication]', 'Additional Comments')
        self.additionalcomments = self.returnstring('Additional Comments', 'v4.1')
        self.impact1 = self.returnstring('Select Impact level (as rated by the customer)\n\n\n\n \n\n\n\n[Implication]', 'Describe the impact\n\n\n\n \n\n\n\n[Implication]')
        self.impact2 = self.returnstring('Describe the impact\n\n\n\n \n\n\n\n[Implication]', 'Rate the level of impact to acQuire business\n\n\n\n(if applicable)')
        self.impact3 = self.returnstring('Rate the level of impact to acQuire business\n\n\n\n(if applicable)', 'Can you quantify the impact:')
        self.impact4 = self.returnstring('Can you quantify the impact:', 'Workaround(s)')

    def readcontents(self, file):
        a = ''
        b = open(r'{}{}'.format(DIRECTORY, file), 'r', encoding='ISO-8859-1')
        for i in b:
            a = a + '{}'.format(i.replace('\x00', ''))
        b.close()
        return a

    def returnstring(self, fromtring, tostring):
        return self.contents.split(fromtring)[1].split(tostring)[0].replace('\n', '').strip()


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('                     Reads and processes contents of RFE')
    print('=============================================================================\n')
    f = input('RFE email filename (*.msg): ')
    print('\n-----------------------------------------------------------------------------\n')
    rfe = ReadRFE(f)
    print(OUTPUTRFE.format(rfe.customer, rfe.site, rfe.user, rfe.role, rfe.owner, rfe.si, rfe.version,
                           rfe.userstory, rfe.description, rfe.workaround, rfe.additionalcomments))


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
