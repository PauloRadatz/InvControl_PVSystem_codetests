# -*- coding: iso-8859-15 -*-



import ctypes
import win32com.client
from win32com.client import makepy
# from numpy import *
from pylab import *
from comtypes import automation
import os
import sys

class DSS(object):

    def __init__(self, tipo):
        """ Cria um novo objeto DSS """

        if tipo == 1:
            # These variables provide direct interface into OpenDSS
            # sys.argv = ["makepy", r"OpenDSSEngine.DSS"]
            # makepy.main()  # ensures early binding and improves speed

            # Create a new instance of the DSS
            self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

            # Start the DSS
            if not self.dssObj.Start(0):
                ctypes.windll.user32.MessageBoxW(0, u"Nao foi possível realizar a conexao com o OpenDSS", u"Conexao OpenDSS", 0)

            # Acesso direto para as interfaces do OpenDSS
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssMeters = self.dssCircuit.Meters
            self.dssMonitors = self.dssCircuit.Monitors
            self.dssPDElement = self.dssCircuit.PDElements
            self.dssSource = self.dssCircuit.Vsources
            self.dssTransformer = self.dssCircuit.Transformers
            self.dssLines = self.dssCircuit.Lines
            self.dssLoads = self.dssCircuit.Loads
            self.dssRegulators = self.dssCircuit.RegControls
            self.dssCapControls = self.dssCircuit.CapControls
            self.dssTopology = self.dssCircuit.Topology

        else:
            print(str(os.getcwd()))
            # Importa a DLL
            # os.chdir("..")  # Sai de "Fontes" e vai para "OpenDSS_ANEEL"
            # os.chdir("Fontes")
            os.chdir(os.path.dirname(sys.argv[0]))
            print(os.getcwd())
            self.dssObj = ctypes.WinDLL("OpenDSSDirect.dll")
            self.AllocateMemory()

            if int(self.dssObj.DSSI(ctypes.c_int32(3), ctypes.c_int32(0))) == 1:
                versao = ctypes.c_char_p(self.dssObj.DSSS(ctypes.c_int32(1), "".encode('ascii')))
                print(u"Conexao com o OpenDSSDirect.dll realizada com sucesso! Versao: " + versao.value)

            # Desativa formulários
            self.dssObj.DSSI(ctypes.c_int32(8), ctypes.c_int32(0))

    def AllocateMemory(self):
        """ Método que arruma o problema de alocação de memória para interfaces do tipo Float """
        self.dssObj.TransformersF.restype = ctypes.c_double
        self.dssObj.BUSF.restype = ctypes.c_double
        self.dssObj.DSSLoadsF.restype = ctypes.c_double
        self.dssObj.DSSLoadsS.restype = ctypes.c_char_p
        self.dssObj.DSSS.restype = ctypes.c_char_p
        # self.dssObj.DSSPut_Command = ctypes.c_char_p
        self.dssObj.CircuitS.restype = ctypes.c_char_p
        self.dssObj.TopologyS.restype = ctypes.c_char_p
        self.dssObj.MetersS.restype = ctypes.c_char_p
        self.dssObj.RegControlsS.restype = ctypes.c_char_p
        self.dssObj.DSSLoadsS.restype = ctypes.c_char_p
        self.dssObj.TransformersS.restype = ctypes.c_char_p
        self.dssObj.DSSLoadsF.restype = ctypes.c_double
        self.dssObj.CircuitS.restype = ctypes.c_char_p
        self.dssObj.GeneratorsS.restype = ctypes.c_char_p

    def ClearAll(self):
        """ Método que limpa a memória do OpenDSSDirect.dll"""
        self.dssObj.DSSI(ctypes.c_int32(1), ctypes.c_int32(0))

    def Text(self, comando):
        """ Método que envia um comando pela OpenDSSDirect.dll """
        self.dssObj.DSSPut_Command(comando.encode('ascii'))
        # resposta = ctypes.c_char_p(self.dssObj.DSSPut_Command(comando.encode('ascii')))
        # resposta = ctypes.c_char_p(self.dssObj.DSSPut_Command(comando.encode('ascii')))
        # return resposta.value.decode('ascii')
        # return resposta.value  # não é necessário decodificar.

    def Set_Topology_ActiveBranchName(self, nome):
        """ Método que seta um ramo ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.TopologyS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Get_Topology_ActiveBranchName(self):
        """ Método que retorna o nome do ramo ativo"""
        resposta = ctypes.c_char_p(self.dssObj.TopologyS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def Topology_BackwardBranch(self):
        """ Método que torna o ramo à montante ativo. Retorna o index do elemento se encontrou. Caso não haja mais, retorna 0"""
        resposta = int(self.dssObj.TopologyI(ctypes.c_int32(7), ctypes.c_int32(0)))
        return resposta

    def Set_Transfomer_ActiveName(self, nome):
        """ Método que seta um transformador ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.TransformersS(ctypes.c_int32(3), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Set_Transformer_kVA(self, kva):
        """ Método que seta a potência aparente do transformador ativo"""
        resposta = float(self.dssObj.TransformersF(ctypes.c_int32(11), ctypes.c_double(kva)))
        return resposta

    def Get_Transformer_kVA(self):
        """ Método que retorna a potência aparente do transformador ativo"""
        resposta = float(self.dssObj.TransformersF(ctypes.c_int32(10), ctypes.c_double(0)))
        return resposta

    def Get_Solution_Converged(self):
        """ Método que indica se a solução do circuito convergiu ou não. Retorna 1, caso positivo e 0, caso negativo"""
        resposta = int(self.dssObj.SolutionI(ctypes.c_int32(38), ctypes.c_int32(0)))
        return resposta

    def Get_Meters_AllNames(self):
        """ Método que retorna uma lista com o nome de todos os medidores do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def Set_Meters_ActiveName(self, nome):
        """ Método que seta um medidor ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.MetersS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Get_Meters_ActiveName(self):
        """ Método que retorna o nome do medidor ativo"""
        resposta = ctypes.c_char_p(self.dssObj.MetersS(ctypes.c_int32(0), ctypes.c_int32(0)))
        # return resposta.value.decode('ascii')
        return str(resposta.value)


    def Get_Meters_RegisterNames(self):
        """ Método que retorna uma lista com o nome de todos os registros do medidor ativo"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(1), variant_pointer)
        return variant_pointer.contents.value

    def Get_Meters_RegisterValues(self):
        """ Método que retorna uma lista com os valores do registros do medidor ativo"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.MetersV(ctypes.c_int(2), variant_pointer)
        return variant_pointer.contents.value

    def Get_Circuit_AllBusNames(self):
        """ Método que retorna uma lista com os nomes de todas as barras do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int(7), variant_pointer)
        return variant_pointer.contents.value

    def Get_Circuit_AllNodeNames(self):
        """ Método que retorna uma lista com os nomes de todas os nós do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int(10), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def Set_Circuit_ActiveBus_Name(self, nome):
        """ Método que torna uma barra ativa pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(4), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Get_Bus_kVBase(self):
        """ Método que retorna o kVBase da barra ativa"""
        resposta = float(self.dssObj.BUSF(ctypes.c_int32(0), ctypes.c_double(0)))
        return resposta

    def Get_Loads_AllNames(self):
        """ Método que retorna uma lista com os nomes de todas as cargas do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.DSSLoadsV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def Get_Loads_Count(self):
        """ Método que retorna a quantidade de cargas no circuito"""
        resposta = int(self.dssObj.DSSLoads(ctypes.c_int32(4), ctypes.c_int32(0)))
        return resposta

    def Set_Loads_First(self):
        """ Método que ativa a primeira carga registrada no circuito"""
        resposta = int(self.dssObj.DSSLoads(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta

    def Set_Loads_Next(self):
        """ Método que ativa a próxima carga registrada no circuito, tendo como referência a carga atualmente ativa"""
        resposta = int(self.dssObj.DSSLoads(ctypes.c_int32(1), ctypes.c_int32(0)))
        return resposta

    def Get_CktElement_BusNames(self):
        """ Método que retorna uma lista com o nome das barras à qual o elemento ativo se conecta"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CktElementV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def Get_CktElement_NumPhases(self):
        """ Método que retorna a quantidade de fases do elemento ativo"""
        resposta = int(self.dssObj.CktElementI(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta

    def Get_Loads_kV(self):
        """ Método que retorna o valor de kV da carga ativa"""
        resposta = float(self.dssObj.DSSLoadsF(ctypes.c_int32(2), ctypes.c_double(0)))
        return resposta

    def Get_Transformer_AllNames(self):
        """ Método que retorna uma lista com os nomes de todos os Transformadores"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.TransformersV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def Set_Circuit_ActiveElement_Name(self, nome):
        """ Método que torna um elemento ativo pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(3), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Get_Circuit_AllBusVmagPu(self):
        """ Método que retorna uma lista com as tensões nodais em pu de todos os nós do circuito"""
        # Declarando um ponteiro para uma variável do tipo "variant"
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int(9), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def Monitors_ResetAll(self):
        """ Método que reseta todos os monitores do circuito"""
        resposta = int(self.dssObj.MonitorsI(ctypes.c_int32(3), ctypes.c_int32(0)))
        return resposta

    def Set_Loads_ActiveName(self, nome):
        """ Método que torna uma carga ativa pelo nome"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(1), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def Set_Loads_kW(self, kW):
        """ Método que seta o valor de kW da carga ativa"""
        resposta = float(self.dssObj.DSSLoadsF(ctypes.c_int32(1), ctypes.c_double(kW)))
        return resposta

    def Set_Loads_daily(self, nome):
        """ Método que seta o valor do parâmetro daily da carga ativa"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(5), nome.encode('ascii')))
        return resposta.value.decode('ascii')
        return variant_pointer.contents.value

    def Meters_ResetAll(self):
        """ Método que limpa todos os medidores"""
        resposta = self.dssObj.MetersI(ctypes.c_int32(3), ctypes.c_int32(0))
        return resposta

    def Monitors_ResetAll(self):
        """ Método que limpa todos os monitores"""
        resposta = self.dssObj.MonitorsI(ctypes.c_int32(3), ctypes.c_int32(0))
        return resposta

    def set_VSource_First(self):
        """ Ativa o primeiro elemento Vsource"""
        resposta = self.dssObj.VSourcesI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_VSource_AllNames(self):
        """ Método que retorna o nome de todas as fontes"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.VsourcesV(ctypes.c_int32(0), variant_pointer)
        return variant_pointer.contents.value


    def set_RegControl_First(self):
        """ Ativa o primeiro elemento RegControl"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_RegControl_Next(self):
        """ Ativa o próximo elemento RegControl"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_RegControl_Count(self):
        """ Método que retorna a quantidade de RegControls"""
        resposta = self.dssObj.RegControlsI(ctypes.c_int32(12), ctypes.c_int32(0))
        return resposta

    def get_RegControl_Transformer(self):
        """ Método que retorna o nome do transformador controlado pelo RegControl ativado"""
        resposta = ctypes.c_char_p(self.dssObj.RegControlsS(ctypes.c_int32(4), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_RegControl_Name(self):
        """ Método que retorna o nome do RegControl ativado"""
        resposta = ctypes.c_char_p(self.dssObj.RegControlsS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def set_Load_First(self):
        """ Ativa o primeiro elemento Load"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_Load_Next(self):
        """ Ativa o próximo elemento Load"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_Load_Count(self):
        """ Método que retorna a quantidade de Loads"""
        resposta = self.dssObj.DSSLoads(ctypes.c_int32(4), ctypes.c_int32(0))
        return resposta

    def get_Load_Name(self):
        """ Método que retorna o nome da carga ativada"""
        resposta = ctypes.c_char_p(self.dssObj.DSSLoadsS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_Load_kV(self):
        """ Método que retorna o kV da carga ativada"""
        resposta = float(self.dssObj.DSSLoadsF(ctypes.c_int32(2), ctypes.c_double(0)))
        return resposta

    def get_Load_AllNames(self):
        """ Método que retorna o nome de todas as cargas"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.DSSLoadsV(ctypes.c_int32(0), variant_pointer)
        return variant_pointer.contents.value

    def get_CktElement_BusNames(self):
        """ Método que retorna os nomes das barras do elemento ativo"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CktElementV(ctypes.c_int32(0), variant_pointer)
        return variant_pointer.contents.value

    def get_CktElement_NumPhases(self):
        """ Método que retorna a contidade de fases do elemento ativo ---- Confirmar se o terminal precisa estar ativo"""
        resposta = self.dssObj.CktElementI(ctypes.c_int32(2), ctypes.c_int32(0))
        return resposta

    def set_Transformer_First(self):
        """ Ativa o primeiro elemento Transformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(8), ctypes.c_int32(0))
        return resposta

    def set_Transformer_Next(self):
        """ Ativa o próximo elementoTransformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(9), ctypes.c_int32(0))
        return resposta

    def get_Transformer_Count(self):
        """ Método que retorna a quantidade de Transformer """
        resposta = self.dssObj.TransformersI(ctypes.c_int32(10), ctypes.c_int32(0))
        return resposta

    def get_Transformer_Name(self):
        """ Método que retorna o nome do transformer ativado"""
        resposta = ctypes.c_char_p(self.dssObj.TransformersS(ctypes.c_int32(2), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')

    def get_Transformer_AllNames(self):
        """ Método que retorna os nomes dos transformeres"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.TransformersV(ctypes.c_int(0), variant_pointer)
        return variant_pointer.contents.value

    def get_Transformer_NumWindings(self):
        """ Método que retorna a quantidade de enrolamentos do transformador ativado"""
        resposta = self.dssObj.TransformersI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def get_Circuit_AllBusVMagPu(self):
        """ Método que retorna as tensões das barras em pu"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(9), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def get_Circuit_AllNodeNames(self):
        """ Método que get os nomes dos nós"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(10), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def get_Circuit_AllBusNames(self):
        """ Método que get os nomes das barras"""
        variant_pointer = ctypes.pointer(automation.VARIANT())
        self.dssObj.CircuitV(ctypes.c_int32(7), variant_pointer, ctypes.c_int32(0))
        return variant_pointer.contents.value

    def set_Circuit_SetActiveBusi(self, i):
        """ Método que ativa a barra i"""
        resposta = self.dssObj.CircuitI(ctypes.c_int32(9), ctypes.c_int32(i))
        return resposta

    def set_Circuit_SetActiveElement(self, nome):
        """ Método que ativa a barra i"""
        resposta = ctypes.c_char_p(self.dssObj.CircuitS(ctypes.c_int32(0), nome.encode('ascii')))
        return resposta.value.decode('ascii')

    def get_Bus_kVBase(self):
        """ Método que retorna o kVBase da barra ativa"""
        resposta = float(self.dssObj.BUSF(ctypes.c_int32(0), ctypes.c_double(0)))
        return resposta

    def set_Generators_First(self):
        """ Ativa o primeiro elemento generator"""
        resposta = self.dssObj.GeneratorsI(ctypes.c_int32(0), ctypes.c_int32(0))
        return resposta

    def set_Generators_Next(self):
        """ Ativa o próximo elemento generator"""
        resposta = self.dssObj.GeneratorsI(ctypes.c_int32(1), ctypes.c_int32(0))
        return resposta

    def get_Generators_Count(self):
        """ Método que retorna a quantidade de generator"""
        resposta = self.dssObj.GeneratorsI(ctypes.c_int32(6), ctypes.c_int32(0))
        return resposta

    def get_Generators_Name(self):
        """ Método que retorna o nome do gerador ativado"""
        resposta = ctypes.c_char_p(self.dssObj.GeneratorsS(ctypes.c_int32(0), ctypes.c_int32(0)))
        return resposta.value.decode('ascii')
