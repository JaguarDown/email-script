#!/usr/bin/python

class ConsoleColor:
    OKGREEN = '\033[92m'
    FAILRED = '\033[91m'
    ENDC = '\033[0m'
    OKBLUE = '\033[94m'

    def ok(self):
        print "[" + self.OKGREEN + "OK" + self.ENDC + "]"

    def fail(self):
        print "[" + self.FAILRED + "FAIL" + self.ENDC + "]"

    def color_print(self, text):
        print self.OKBLUE + text + self.ENDC
