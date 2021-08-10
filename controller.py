'''
Formbook Beacon Decode POC
Author: James Slaughter
CIRT Malware Analysis Challenge
'''

'''
controller.py - This file is responsible for keeping global settings available through class properties
'''

#python imports
import imp
import sys

'''
controller
Class: This class is is responsible for keeping global settings available through class properties
'''
class controller:
    '''
    Constructor
    '''
    def __init__(self):

        self.debug = False
        self.xls = ''
        self.wb = ''
        

