# -*- coding: iso-8859-15 -*-

import os
import timeit
import sys
# import DSS
# import tkFileDialog
# import tkMessageBox
# import datetime
import pandas as pd
import DSS
import operator
import traceback
import math



class Scenario:
    true_list = ["1", "Yes", "yes", 1, 1.0]

    def __init__(self, row, objMethodology, pvsystem_folder):

        self.objMethodology = objMethodology

        self.id = int(row["Scenario ID"])
        self.OpenDSS = row["OpenDSS"]
        self.file_dss = pvsystem_folder + "\\" + row["Master DSS"]

        self.scenario_name = self.file_dss.split("\\")[-1].split(".")[0]

        self.kVA = float(row["kVA Rating"])
        self.Pmpp = float(row["Pmpp"])
        self.pctPmpp = float(row["pctPmpp"])
        self.kvarlimit = float(row["kvar Limit"])
        self.Vbase = float(row["V Base"])

        self.master_dir = pvsystem_folder + "/" + os.path.dirname(row["Master DSS"])

        self.scenario_output_dir = self.master_dir + r"\results_" + self.OpenDSS

        if not os.path.exists(self.scenario_output_dir):
            os.makedirs(self.scenario_output_dir)

        if type(row["VV Curve Y"]) is str:
            self.VV_y = [float(i) for i in row["VV Curve Y"].split(" ")]

        if type(row["VV Curve X"]) is str:
            self.VV_x = [float(i) for i in row["VV Curve X"].split(" ")]

        if type(row["VW Curve Y"]) is str:
            self.VW_y = [float(i) for i in row["VW Curve Y"].split(" ")]

        if type(row["VW Curve X"]) is str:
            self.VW_x = [float(i) for i in row["VW Curve X"].split(" ")]

        self.sf = row["Smart Function"]

        if row["Smart Function"] == "VV" or row["Smart Function"] == "VV_VW" or row["Smart Function"] == "VV_DRC":
            self.VV = True
        else:
            self.VV = False

        if row["Smart Function"] == "VW" or row["Smart Function"] == "VV_VW":
            self.VW = True
        else:
            self.VW = False

        if row["Smart Function"] == "DRC" or row["Smart Function"] == "VV_DRC":
            self.DRC = True
        else:
            self.DRC = False

        if self.DRC:
            self.arDRC = float(row["arDRC"])

        if row["Qava"] in Scenario.true_list:
            self.Qava = True
        else:
            self.Qava = False

        self.Pava = row["PY"]

        if row["Q Option"] == "LPF":
            self.Q_LPF = True
            self.alpha = math.exp(-1* float(row["stepsize"]) / float(row["tau"]))
        else:
            self.Q_LPF = False

        if row["Q Option"] == "RF":
            self.Q_RF = True
            self.step = float(row["stepsize"])
            self.delta_Q = float(row["delta Q"])
        else:
            self.Q_RF = False

        if row["P Option"] == "LPF":
            self.P_LPF = True
            self.alpha = math.exp(-1* float(row["stepsize"]) / float(row["tau"]))
        else:
            self.P_LPF = False

        if row["P Option"] == "RF":
            self.P_RF = True
            self.step = float(row["stepsize"])
            self.delta_P = float(row["delta P"])
        else:
            self.P_RF = False

        if row["Vref"] == "avg":
            self.Vref = "avg"
        elif row["Vref"] == "ravg":
            self.Vref = "ravg"
        else:
            self.Vref = "rated"

        self.objMethodology.set_scenario(self)

        self.df_results = pd.DataFrame()


    def run(self):

        self.objMethodology.set_scenario(self)

        self.objMethodology.run()