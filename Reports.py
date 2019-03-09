# encoding: utf-8


import math
import pandas as pd
import numpy as np
import traceback


class Reports:

    result_columns = ["hour", " Irradiance", "Ptotal", "Qtotal", "Stotal", "v1 (pu)", "Vreg", "QVV Calculated", "QDRC Calculated", "Q LPF or RF Calculated", "Qtotal Calculated", "kW_out_desired", "PVW Calculated", "PVW LPF or RF Calculated", "Ptotal Calculated", "volt-var", "volt-watt", "DRC", "VV_DRC"]

    def __init__(self, objMethodology):

        self.scenario = None
        self.objMethodology = objMethodology

        self.df_OpenDSS_power_results = pd.DataFrame()
        self.df_OpenDSS_voltage_results = pd.DataFrame()

        self.df_calculated_power_results = pd.DataFrame()


    def set_scenario(self, scenario):
        """ Seta o cenário ativo para o objeto GeraRel"""
        self.scenario = scenario
        self.drop_columns()

    def reads_OpenDSS_power_results(self):

        df_power_temp = pd.read_csv(self.scenario.master_dir + "/" + str(self.scenario.scenario_name) + "_Mon_pv_powers_1.csv", engine='python')

        self.df_OpenDSS_power_results["hour"] = df_power_temp["hour"]
        self.df_OpenDSS_power_results["Ptotal"] = -1 * (df_power_temp[" P1 (kW)"] + df_power_temp[" P2 (kW)"] + df_power_temp[" P3 (kW)"])
        self.df_OpenDSS_power_results["Qtotal"] = -1 * (df_power_temp[" Q1 (kvar)"] + df_power_temp[" Q2 (kvar)"] + df_power_temp[" Q3 (kvar)"])
        self.df_OpenDSS_power_results["Stotal"] = (self.df_OpenDSS_power_results["Ptotal"]**2 + self.df_OpenDSS_power_results["Qtotal"]**2)**(1/2)
        self.df_OpenDSS_power_results["Qava"] = (self.scenario.kVA**2 - self.df_OpenDSS_power_results["Ptotal"]**2)**(1/2)

    def reads_OpenDSS_voltage_results(self):

        df_voltage_temp = pd.read_csv(self.scenario.master_dir + "/" + str(self.scenario.scenario_name) + "_Mon_pv_currents_1.csv", engine='python').dropna()
        self.df_OpenDSS_voltage_results["hour"] = df_voltage_temp["hour"]
        self.df_OpenDSS_voltage_results["v1 (pu)"] = df_voltage_temp[" V1"] / self.scenario.Vbase

    def reads_OpenDSS_internal_results(self):

        df_internal_temp = pd.read_csv(self.scenario.master_dir + "/" + str(self.scenario.scenario_name) + "_Mon_pv_v_1.csv", engine='python')

        self.df_OpenDSS_internal_results = df_internal_temp[["hour", " Irradiance", "Vreg" ,"Vavg (DRC)", "volt-var", "volt-watt", "DRC", "VV_DRC", "kW_out_desired"]]

    def drop_columns(self):

        if self.scenario.sf == "VV":
            self.droped_columns = ["QDRC Calculated", "PVW Calculated", "PVW LPF or RF Calculated", "volt-watt", "DRC", "VV_DRC", "Ptotal Calculated", "kW_out_desired"]
        elif self.scenario.sf == "DRC":
            self.droped_columns = ["QVV Calculated", "PVW Calculated", "PVW LPF or RF Calculated", "volt-watt", "volt-var", "VV_DRC", "Ptotal Calculated", "kW_out_desired"]
        elif self.scenario.sf == "VV_DRC":
            self.droped_columns = ["PVW Calculated", "PVW LPF or RF Calculated", "volt-watt", "volt-var", "DRC", "Ptotal Calculated", "kW_out_desired"]
        elif self.scenario.sf == "VW":
            self.droped_columns = ["QVV Calculated", "QDRC Calculated", "Q LPF or RF Calculated", "Qtotal Calculated", "volt-var", "DRC", "VV_DRC", "Ptotal Calculated"]
        elif self.scenario.sf == "VV_VW":
            self.droped_columns = ["QDRC Calculated", "DRC", "VV_DRC", "Ptotal Calculated"]

    def calculates_power_results(self):

        Qtotal = pd.DataFrame()
        Ptotal = pd.DataFrame()


        self.df_calculated_power_results["hour"] = self.df_OpenDSS_voltage_results["hour"]
        self.df_calculated_power_results["PVW Calculated"] = np.nan
        self.df_calculated_power_results["QVV Calculated"] = np.nan
        self.df_calculated_power_results["QDRC Calculated"] = np.nan
        self.df_calculated_power_results["Qtotal Calculated"] = np.nan
        self.df_calculated_power_results["Ptotal Calculated"] = np.nan
        self.df_calculated_power_results["Q LPF or RF Calculated"] = np.nan
        self.df_calculated_power_results["PVW LPF or RF Calculated"] = np.nan
        self.df_calculated_power_results["delta Q"] = np.nan
        Qtotal["Qtotal"] = self.df_calculated_power_results["Qtotal Calculated"]
        Ptotal["Ptotal"] = self.df_calculated_power_results["Ptotal Calculated"]

        if self.scenario.Qava:
            Qbase = self.df_OpenDSS_power_results["Qava"]
        else:
            Qbase = self.scenario.kvarlimit

        if self.scenario.Pava == "PAVAILABLEPU":
            Pbase = self.scenario.Pmpp * self.df_OpenDSS_internal_results[" Irradiance"]
        elif self.scenario.Pava == "PCTPMPPPU":
            Pbase = self.scenario.Pmpp * self.scenario.pctPmpp / 100.0
        elif self.scenario.Pava == "PMPPPU":
            Pbase = self.scenario.Pmpp

        if self.scenario.VV:
            self.get_VV_curve()

            if self.scenario.Vref == "rated":
                self.df_calculated_power_results["QVV Calculated"] = (self.df_OpenDSS_voltage_results["aVV"] * self.df_OpenDSS_voltage_results["v1 (pu)"] + self.df_OpenDSS_voltage_results["bVV"]) * Qbase
            else:
                self.df_calculated_power_results["QVV Calculated"] = (self.df_OpenDSS_voltage_results["aVV"] * self.df_OpenDSS_internal_results["Vreg"] + self.df_OpenDSS_voltage_results["bVV"]) * Qbase

            Qtotal["Qtotal"] = self.df_calculated_power_results["QVV Calculated"]
        if self.scenario.VW:
            self.get_VW_curve()
            self.df_calculated_power_results["PVW Calculated"] = (self.df_OpenDSS_voltage_results["aVW"] * self.df_OpenDSS_voltage_results["v1 (pu)"] + self.df_OpenDSS_voltage_results["bVW"]) * Pbase
            Ptotal["Ptotal"] = self.df_calculated_power_results["PVW Calculated"]
        if self.scenario.DRC:

            self.df_calculated_power_results["QDRC Calculated"] = -1 * self.scenario.arDRC * Qbase * (self.df_OpenDSS_voltage_results["v1 (pu)"] - self.df_OpenDSS_internal_results["Vavg (DRC)"])
            self.df_calculated_power_results["QDRC Calculated"].iloc[0] = 0.0
            Qtotal["Qtotal"] = self.df_calculated_power_results["QDRC Calculated"]

        if self.scenario.VV and self.scenario.DRC:
            Qtotal["Qtotal"] = self.df_calculated_power_results["QVV Calculated"] + self.df_calculated_power_results["QDRC Calculated"]

        if self.scenario.Q_LPF:
            self.get_Q_LPF(Qtotal)

        elif self.scenario.Q_RF:
            self.df_calculated_power_results["delta Q"] = Qbase * self.scenario.delta_Q
            self.get_Q_RF(Qtotal)

        else:
            self.df_calculated_power_results["Qtotal Calculated"] = Qtotal

        if self.scenario.P_LPF:
            self.get_P_LPF()

        elif self.scenario.P_RF:
            self.df_calculated_power_results["delta P"] = Pbase * self.scenario.delta_P
            self.get_P_RF()
        else:
            self.df_calculated_power_results["Ptotal Calculated"] = Ptotal



    def creates_results_report(self):
        df = [self.df_OpenDSS_power_results, self.df_OpenDSS_voltage_results, self.df_OpenDSS_internal_results, self.df_calculated_power_results]
        df_results_temp = pd.concat(df, axis=1)[Reports.result_columns].drop(self.droped_columns, axis=1)

        self.scenario.df_results = df_results_temp.iloc[:, ~df_results_temp.columns.duplicated()]

    def export_reports(self):

        self.reads_OpenDSS_power_results()
        self.reads_OpenDSS_voltage_results()
        self.reads_OpenDSS_internal_results()

        self.calculates_power_results()

        self.creates_results_report()

        self.scenario.df_results.to_csv(self.scenario.scenario_output_dir + "/" + self.scenario.scenario_name + ".csv", index=False, sep=",")

    def get_VV_curve(self):

        coef_dic = {}

        aVV_list = []
        bVV_list = []

        for i in range(len(self.scenario.VV_x)):

            if i != 0:
                m = (self.scenario.VV_y[i] - self.scenario.VV_y[i-1]) / (self.scenario.VV_x[i] - self.scenario.VV_x[i-1])
                b = self.scenario.VV_y[i] - m * self.scenario.VV_x[i]

                coef_dic[i] = [m, b]


        for index, voltage in self.df_OpenDSS_voltage_results.iterrows():
            m = np.nan
            b = np.nan
            for i in range(len(self.scenario.VV_x)):

                if i != 0:
                    if voltage["v1 (pu)"] >= self.scenario.VV_x[i-1] and voltage["v1 (pu)"] < self.scenario.VV_x[i]:

                        m = coef_dic[i][0]
                        b = coef_dic[i][1]

            aVV_list.append(m)
            bVV_list.append(b)

        self.df_OpenDSS_voltage_results["aVV"] = aVV_list
        self.df_OpenDSS_voltage_results["bVV"] = bVV_list

    def get_VW_curve(self):

        coef_dic = {}

        aVW_list = []
        bVW_list = []

        for i in range(len(self.scenario.VW_x)):

            if i != 0:
                m = (self.scenario.VW_y[i] - self.scenario.VW_y[i-1]) / (self.scenario.VW_x[i] - self.scenario.VW_x[i-1])
                b = self.scenario.VW_y[i] - m * self.scenario.VW_x[i]

                coef_dic[i] = [m, b]


        for index, voltage in self.df_OpenDSS_voltage_results.iterrows():
            m = np.nan
            b = np.nan
            for i in range(len(self.scenario.VW_x)):


                if i != 0:
                    if voltage["v1 (pu)"] >= self.scenario.VW_x[i - 1] and voltage["v1 (pu)"] < self.scenario.VW_x[i]:
                        m = coef_dic[i][0]
                        b = coef_dic[i][1]

            aVW_list.append(m)
            bVW_list.append(b)

        self.df_OpenDSS_voltage_results["aVW"] = aVW_list
        self.df_OpenDSS_voltage_results["bVW"] = bVW_list


    def get_Q_LPF(self, Q):

        Q_LPF_list = []
        Q_prior = 0.0

        for index, q in Q.iterrows():

            if index == 0:
                Q_LPF_list.append(q["Qtotal"] * (1 - self.scenario.alpha))
            else:
                Q_LPF_list.append(q["Qtotal"] * (1 - self.scenario.alpha) + Q_prior * self.scenario.alpha)

            Q_prior = self.df_OpenDSS_power_results.iloc[index]["Qtotal"]

        self.df_calculated_power_results["Q LPF or RF Calculated"] = Q_LPF_list
        self.df_calculated_power_results["Qtotal Calculated"] = self.df_calculated_power_results["Q LPF or RF Calculated"]

    def get_P_LPF(self):

        P_LPF_list = []
        P_prior = 0.0

        for index, p in self.df_OpenDSS_internal_results.iterrows():

            if index == 0:
                P_LPF_list.append(p["kW_out_desired"] * (1 - self.scenario.alpha))
            else:
                P_LPF_list.append(p["kW_out_desired"] * (1 - self.scenario.alpha) + P_prior * self.scenario.alpha)

            P_prior = self.df_OpenDSS_power_results.iloc[index]["Ptotal"]

        self.df_calculated_power_results["PVW LPF or RF Calculated"] = P_LPF_list
        self.df_calculated_power_results["Ptotal Calculated"] = self.df_calculated_power_results["PVW LPF or RF Calculated"]

    def get_Q_RF(self, Q):

        Q_RF_list = []
        Q_RF_total_list = []
        Q_prior = 0.0

        for index, q1 in Q.iterrows():

            delta_Q = self.df_calculated_power_results.iloc[index]["delta Q"]

            if index == 0:
                if q1["Qtotal"] < 0:
                    q = - delta_Q * self.scenario.step
                    temp = max(q1["Qtotal"],q)
                    Q_RF_total_list.append(temp)
                    temp2 = q
                else:
                    q = delta_Q * self.scenario.step
                    temp = min(q1["Qtotal"], q)
                    Q_RF_total_list.append(temp)
                    temp2 = q
            else:
                if (q1["Qtotal"] - Q_prior) < 0:
                    q = - delta_Q * self.scenario.step
                    temp = Q_prior + max((q1["Qtotal"] - Q_prior), q)
                    Q_RF_total_list.append(temp)
                    temp2 = Q_prior + q
                else:
                    q = delta_Q * self.scenario.step
                    temp = Q_prior + min((q1["Qtotal"] - Q_prior), q)
                    Q_RF_total_list.append(temp)
                    temp2 = Q_prior + q

            Q_RF_list.append(temp2)

            Q_prior = self.df_OpenDSS_power_results.iloc[index]["Qtotal"]

        self.df_calculated_power_results["Q LPF or RF Calculated"] = Q_RF_list
        self.df_calculated_power_results["Qtotal Calculated"] = Q_RF_total_list

    def get_P_RF(self):

        P_RF_list = []
        P_RF_total_list = []
        P_prior = 0.0

        for index, p1 in self.df_OpenDSS_internal_results.iterrows():

            delta_P = self.df_calculated_power_results.iloc[index]["delta P"]

            if index == 0:
                p =  delta_P * self.scenario.step
                temp = min(p1["kW_out_desired"],p)
                P_RF_total_list.append(temp)
                temp2 = p

            else:
                if (p1["kW_out_desired"] - P_prior) < 0:
                    p = - delta_P * self.scenario.step
                    temp = P_prior + max((p1["kW_out_desired"] - P_prior), p)
                    P_RF_total_list.append(temp)
                    temp2 = P_prior + p
                else:
                    p = delta_P * self.scenario.step
                    temp = P_prior + min((p1["kW_out_desired"] - P_prior), p)
                    P_RF_total_list.append(temp)
                    temp2 = P_prior + p

            P_RF_list.append(temp2)

            P_prior = self.df_OpenDSS_power_results.iloc[index]["Ptotal"]

        self.df_calculated_power_results["PVW LPF or RF Calculated"] = P_RF_list
        self.df_calculated_power_results["Ptotal Calculated"] = P_RF_total_list