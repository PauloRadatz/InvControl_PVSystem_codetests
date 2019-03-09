# -*- coding: iso-8859-15 -*-

import os
import timeit
# import DSS
# import tkFileDialog
# import tkMessageBox
import pandas as pd
import sys
import Methodology
import Scenario
import traceback
from tkinter import filedialog
from tkinter import *

class Main:

    ScenarioList = []  

    def __init__(self):

        # inicial

        # os.chdir("..")
        # os.chdir("Input")

        # scenario_input_dir = os.path.abspath(os.curdir)

        # scenarios = r"\Scenarios_VVVW.csv"
        # scenarios = r"\Scenarios_VV_LPF.csv"
        # scenarios = r"\Scenarios_VV_RF.csv"
        # scenarios = r"\Scenarios_VVDRC.csv"
        # scenarios = r"\Scenarios_VV.csv"
        # scenarios = r"\Scenarios_DRC.csv"
        # scenarios = r"\Scenarios_VW.csv"
        # scenarios = r"\Scenarios_VVVW.csv"

        print("Please select the Scenarios file (*.csv file).\n")
        scenario_file = filedialog.askopenfilename()

        #scenario_file = scenario_input_dir + scenarios

        df_scenarios = pd.read_csv(scenario_file, engine='python', sep=",")

        objMethodology = Methodology.Methodology()

        os.chdir(os.path.dirname(os.path.abspath(scenario_file)))
        os.chdir("..")
        os.chdir("..")
        os.chdir("PVSystem")

        pvsystem_folder = os.path.abspath(os.curdir)

        print("Reading input file with scenarios")
        for index, row in df_scenarios.iterrows():

            if str(row["Scenario ID"]) == "End":
                exit()

            scenario_name = row["Master DSS"].split("\\")[-1]
            print("\nScenario " + str(row["Scenario ID"]) + ". Scenario name: " + scenario_name)
            objScenario = Scenario.Scenario(row, objMethodology, pvsystem_folder)

            objScenario.run()


if __name__ == '__main__':
    Main()
    # raw_input("Pressione Enter para fechar a janela...")


