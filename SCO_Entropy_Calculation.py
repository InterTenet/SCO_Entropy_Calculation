import numpy as np
import pandas as pd

R = 8.31 # Constante dos gases
Kb = 1.38*pow(10, -23) # Constante de Boltzmann
h = 6.63*pow(10, -34) # Constante de Planck
c = 2.9*pow(10, 8) # Velocidade da luz

T = pd.read_excel(r'C:\Users\Romulo\Desktop\UFRJ\Quantica\projeto\Data\Data.xlsx', 'Temperatures', skiprows=1, header=None)
temperature = np.array(T[0])
vibration = pd.read_excel(r'C:\Users\Romulo\Desktop\UFRJ\Quantica\projeto\Data\Data.xlsx', 'Vibrations')
modes = np.linspace(1,len(vibration), len(vibration)) # Modos vibracionais
DATA_IND = []; DATA_DIF = []

def S(T,v):
    return R*(h*v/(Kb*T*(np.exp((h*v)/(Kb*T))-1))-np.log(1-np.exp(-h*v/(Kb*T))))

for j in range(0, len(modes)):
    mode = int(modes[j])
    dataline_ind = []
    dataline_dif = []
    dataline_ind.append(mode)
    dataline_dif.append(mode)
    for i in temperature:
        s_l = round(S(i, vibration.iloc[j, 0]*100*c), 2)
        s_h = round(S(i, vibration.iloc[j, -1]*100*c), 2)
        dif_s = round((s_h-s_l), 2)
        dataline_ind.append(s_l)
        dataline_ind.append(s_h)
        dataline_dif.append(dif_s)
    DATA_IND.append(dataline_ind)
    DATA_DIF.append(dataline_dif)

DF_IND = pd.DataFrame(DATA_IND)
DF_DIF = pd.DataFrame(DATA_DIF)
TEMPERATURE_DIF = temperature.tolist()
TEMPERATURE_IND = sorted(TEMPERATURE_DIF + TEMPERATURE_DIF)
TEMPERATURE_DIF.insert(0, 'Mode')
TEMPERATURE_IND.insert(0, 'Mode')

with pd.ExcelWriter(r'C:\Users\Romulo\Desktop\UFRJ\Quantica\projeto\Data\entropy_data.xlsx') as archive:
    DF_IND.to_excel(archive, sheet_name='Individual Entropy', index=False, header = TEMPERATURE_IND)
    DF_DIF.to_excel(archive, sheet_name='Entropy Difference', index=False, header = TEMPERATURE_DIF)