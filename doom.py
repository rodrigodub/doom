#################################################
# Doom
# Script to adjust read an RFE email
#
# v0.15
# for ticket #8
#
# Rodrigo Nobrega
# 20180514-20180523
#################################################

# import modules
# import datetime
import os


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
class Readreport(object):
    """
    Reads contents of an RFE Outlook message file
    """
    def __init__(self, file, option):
        # report details
        self.inputfilename = file
        self.outputfilename = file.replace('.msg', '.txt')
        self.contents = self.readcontents()
        self.option = option
        # RFE details
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
        # Bug details

    def readcontents(self):
        a = ''
        b = open(r'{}{}'.format(DIRECTORY, self.inputfilename), 'r', encoding='ISO-8859-1')
        for i in b:
            a = a + '{}'.format(i.replace('\x00', ''))
        b.close()
        return a

    def returnstring(self, fromstring, tostring):
        return self.contents.split(fromstring)[1].split(tostring)[0].replace('\n', '').strip()

    def outputfile(self):
        f = open(r'{}{}'.format(DIRECTORY, self.outputfilename), 'w')
        f.write(OUTPUTRFE.format(self.customer, self.site, self.user, self.role, self.owner, self.si, self.version,
                                 self.userstory, self.description, self.workaround, self.additionalcomments))
        f.close()


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('                     Reads and processes contents of RFE')
    print('=============================================================================\n')
    file = input('Email filename (*.msg) : ')
    option = input('Is this a Bug (B) or RFE (R) report? : ')
    print('\n-----------------------------------------------------------------------------\n')
    if option.upper() in ('B', 'R'):
        report = Readreport(file, option.upper())
        report.outputfile()
        os.system('start {}{}'.format(DIRECTORY, report.inputfilename))
        os.system('start {}{}'.format(DIRECTORY, report.outputfilename))
    else:
        print('The report must be a Bug or an RFE.\nProgram finished.')
    # report = Readreport(f)
    # print(OUTPUTRFE.format(report.customer, report.site, report.user, report.role, report.owner, report.si, report.version,
    #                        report.userstory, report.description, report.workaround, report.additionalcomments))


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
