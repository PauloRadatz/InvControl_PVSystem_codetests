# -*- coding: iso-8859-15 -*-

import os
import timeit
# import DSS
# import tkFileDialog
# import tkMessageBox
# import datetime
import pandas as pd
import DSS

import Reports
import Scenario
import numpy as np
import operator
import itertools
import sys
import traceback


class Methodology:

    def __init__(self):

        self.dss = DSS.DSS(1)
        self.objReports = Reports.Reports(self)
        self.scenario = None

    def set_scenario(self, scenario):
        self.scenario = scenario
        self.objReports.set_scenario(scenario)

    def run(self):

        self.dss.dssObj.ClearAll()

        if self.scenario.file_dss is None:
            return False
        self.dss.dssText.Command = "Compile [" + self.scenario.file_dss + "]"

        self.objReports.export_reports()




