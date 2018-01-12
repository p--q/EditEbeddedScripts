# -*- coding: utf-8 -*-
from . import consts2
def fromsubmod(sheet):
	sheet["A4"].setString("fromsubmod: {}".format(consts2.LISTSHEET["name"]))
	