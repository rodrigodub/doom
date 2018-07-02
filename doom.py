#################################################
# Doom
# Script to adjust read an RFE email
#
# v1.06
# for ticket #
#
# Rodrigo Nobrega
# 20180514-20180702
#################################################

# import modules
import os


# global variables
OUTPUTRFE = 'Summary:\n{}\n' \
            '-----------------------------------------------------------------------------\n' \
            'Customer Name:\n{}\n\n' \
            'Site:\n{}\n\n' \
            'Who has identified the requirement?:\n{}\n\n' \
            'Role:\n{}\n\n' \
            'acQuire RFE Owner:\n{}\n\n' \
            'Original SI#:\n{}\n\n' \
            'GIM Suite Version:\n{}\n\n' \
            'Description:\n{}\n{}\n\n' \
            'Workaround:\n{}\n\n' \
            'Additional Comments:\n{}\n\n'
OUTPUTBUG = 'Summary:\n{}\n' \
            '-----------------------------------------------------------------------------\n' \
            'acQuire Contact (Bug Owner):\n{}\n\n' \
            'Customer:\n{}\n\n' \
            'Database:\n{}\n\n' \
            'Description:\n{}\n{}\n\n' \
            'Replication Steps:\n{}\n\n' \
            'Workaround:\n{}\n\n' \
            'Additional Comments:\n{}\n\n'


# define Readreport class
class Readreport(object):
    """
    Reads contents of an RFE Outlook message file
    """
    def __init__(self, file):
        # report details
        self.inputfilename = file
        self.outputfilename = file.replace('.msg', '.txt')
        self.contents = self.readcontents()
        self.option = self.returnoption()
        # RFE details
        if self.option == 'R':
            try:
                self.customer = self.returnstring('Customer Name', 'Site')
            except IndexError:
                self.customer = ''
            try:
                self.site = self.returnstring('Site', 'Who has identified')
            except IndexError:
                self.site = ''
            try:
                self.user = self.returnstring('Who has identified the requirement?', 'Role')
            except IndexError:
                self.user = ''
            try:
                self.role = self.returnstring('Role', 'GIM Suite Version')
            except IndexError:
                self.role = ''
            try:
                self.version = self.returnstring('GIM Suite Version', 'Triage Details')
            except IndexError:
                self.version = ''
            try:
                self.owner = self.returnstring('acQuire RFE Owner', 'SI#')
            except IndexError:
                self.owner = ''
            try:
                self.si = self.returnstring('SI#', 'RFE Summary')
            except IndexError:
                self.si = ''
            try:
                self.userstory = self.returnstring('User Story', 'Description')
            except IndexError:
                self.userstory = ''
            try:
                self.description = self.returnstring('Description', 'Problem Type')
            except IndexError:
                self.description = ''
            try:
                self.workaround = self.returnstring('Workaround(s)\n\n\n\n \n\n\n\n[Implication]', 'Additional Comments')
            except IndexError:
                self.workaround = ''
            try:
                self.additionalcomments = self.returnstring('Additional Comments', 'v4.1')
            except IndexError:
                self.additionalcomments = ''
            try:
                self.impact1 = self.returnstring('Select Impact level (as rated by the customer)\n\n\n\n \n\n\n\n[Implication]', 'Describe the impact\n\n\n\n \n\n\n\n[Implication]')
            except IndexError:
                self.impact1 = ''
            try:
                self.impact2 = self.returnstring('Describe the impact\n\n\n\n \n\n\n\n[Implication]', 'Rate the level of impact to acQuire business\n\n\n\n(if applicable)')
            except IndexError:
                self.impact2 = ''
            try:
                self.impact3 = self.returnstring('Rate the level of impact to acQuire business\n\n\n\n(if applicable)', 'Can you quantify the impact:')
            except IndexError:
                self.impact3 = ''
            try:
                self.impact4 = self.returnstring('Can you quantify the impact:', 'Workaround(s)')
            except IndexError:
                self.impact4 = ''
        # Bug details
        if self.option == 'B':
            try:
                self.userstory = self.returnstring('Summary', 'Issue Type:')
            except IndexError:
                self.userstory = ''
            try:
                self.customer = self.returnstring('Client:', 'Problem Description:')
            except IndexError:
                self.customer = ''
            try:
                self.owner = self.returnstring('acQuire Contact:', 'Client:')
            except IndexError:
                self.owner = ''
            try:
                self.database = self.returnstring('Database Information:', 'Replication Steps:')
            except IndexError:
                self.database = ''
            try:
                self.description = self.returnstring('Problem Description:', 'Database Information:')
            except IndexError:
                self.description = ''
            try:
                self.replication = self.returnstring('Replication Steps:', 'Workaround (If applicable):')
            except IndexError:
                self.replication = ''
            try:
                self.workaround = self.returnstring('Workaround (If applicable):', 'Additional Comments:')
            except IndexError:
                self.workaround = ''
            try:
                self.additionalcomments = self.returnstring('Additional Comments:', 'acQuire Technology Solutions Pty Ltd')
            except IndexError:
                self.additionalcomments = ''

    def readcontents(self):
        a = ''
        b = open(r'{}'.format(self.inputfilename), 'r', encoding='ISO-8859-1')
        for i in b:
            a = a + r'{}'.format(i.replace('\x00', '').encode('utf-8'))
        b.close()
        return a

    def returnstring(self, fromstring, tostring):
        return self.contents.split(fromstring)[1].split(tostring)[0].replace('\n', '').strip().replace(r"\n'b'\n'b'\n'b'\n'b'", '')

    def outputfile(self):
        f = open(r'{}'.format(self.outputfilename), 'w')
        if self.option == 'R':
            f.write(OUTPUTRFE.format(self.userstory, self.customer, self.site, self.user, self.role, self.owner, self.si,
                                 self.version, self.userstory, self.description, self.workaround, self.additionalcomments))
        elif self.option == 'B':
            f.write(OUTPUTBUG.format(self.userstory, self.owner, self.customer, self.database, self.userstory,
                                     self.description, self.replication, self.workaround, self.additionalcomments))
        f.close()

    def returnoption(self):
        if 'REQUEST FOR ENHANCEMENT (RFE)' in self.contents:
            return 'R'
        elif 'Bug User Story' in self.contents:
            return 'B'


# main loop
def main():
    print('\n=============================================================================')
    print('                                    Doom')
    print('            Reads and processes contents of RFE and Bug reports')
    print('=============================================================================\n')
    directory = input('Email directory : ').replace('\\', '/')
    if directory[-1:] != '/':
        directory += '/'
    filename = input('Email filename (*.msg) : ')
    file = '{}{}'.format(directory, filename)
    print('\n-----------------------------------------------------------------------------\n')
    report = Readreport(file)
    report.outputfile()
    os.system('start {}'.format(report.inputfilename))
    os.system('start {}'.format(report.outputfilename))
    print('Opening files.\nProgram finished.')


# main, calling main loop
if __name__ == '__main__':
    # test()
    main()
