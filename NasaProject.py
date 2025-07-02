import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
df = pd.readcsv('NasaFile.xlsx')

class emcAndemifacility:


    def __init__(self, location, dimensions, frequencyRange, fieldStrengths, anechoicChamberType, standardNumber):
        self.location = location
        self.dimensions = dimensions
        self.frequencyRange = frequencyRange
        self.fieldStrengths = fieldStrengths
        self.anechoicChamberType = anechoicChamberType
        self.standardNumber = standardNumber
    
    def getLocation(self):
        return self.location
    def getDimensions(self):
        return self.dimensions
    def getFrequency(self):
        return self.frequencyRange
    def getFieldStrengths(self):
        return self.fieldStrengths
    def getAnechoicChamberType(self):
        return self.anechoicChamberType
    def getStandardNumber(self):
        return self.standardNumber

daytonTBrownEMIroom1 = emcAndemifacility("usa", "6x6x3.7", "10 kHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, UK DEF STAN 59-41")
daytonTBrownEMIroom2 = emcAndemifacility("usa", "9.8x9.1x3.7", "11 kHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, UK DEF STAN 59-41")
daytonTBrownEMIroom3 = emcAndemifacility("usa", "4.6x5.5x3.0", "12 kHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, UK DEF STAN 59-41")
daytonTBrownEMIroom4 = emcAndemifacility("usa", "6x5x3", "13 kHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, UK DEF STAN 59-41")
daytonTBrownEMIroom5 = emcAndemifacility("usa", "6x9.8x3", "14 kHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, UK DEF STAN 59-41")
daytonTBrownEMIroom6 = emcAndemifacility("usa", "6x6x4.5", "3 MHz-40 GHz", "1.4", "semi", "MIL-STD-461, DO-16O, EN 55011, EN 55022, EN 55032, CISPR 11, CISPR 32")
ngAzusaB183 = emcAndemifacility("NG Azusa - B183", "22' x 23' x 10'", "unknown", "unknown", "fully", "unknown")
ngAzusaB200 = emcAndemifacility("NG Azusa - B200", "20' x 20' x 20'", "unknown", "unknown", "unknown", "unknown")
ngAzusaB200aTent = emcAndemifacility("NG Azusa - B200a Tent", "20' x 20' x 20'", "unknown", "unknown", "fully", "unknown")
ngSpacePark1 = emcAndemifacility("NG Space Park", "9.1 x 7.6 x 5.2", "10 KHz - 40GHz", "0.2 kV/m", "semi", "MIL-STD-461 Rev. A to G")
ngSpacePark2 = emcAndemifacility("NG Space Park", "8.5 x 4.9 x 4.9", "10 KHz - 40GHz", "0.2 kV/m", "semi", "MIL-STD-461 Rev. A to G")
ngSpacePark3 = emcAndemifacility("NG Space Park", "6.7 x 6.1 x 4.9", "10 KHz - 40GHz", "0.2 kV/m", "semi", "MIL-STD-461 Rev. A to G")
tartuObservatory = emcAndemifacility("University of Tartu, Tartu Observatory", "4.28 m x 3.08 m x 2.55 m, test distance 1 m", "30 MHz ... 18 GHz", "max 0.05 kV/m", "semi", "IEC/EN 61000-4-3; ECSS-E-ST-20-07C ; MIL-STD-461G")

EMCAndEMIFacilities = [daytonTBrownEMIroom1,daytonTBrownEMIroom2, daytonTBrownEMIroom3, daytonTBrownEMIroom4, daytonTBrownEMIroom5, daytonTBrownEMIroom6,ngAzusaB183,ngAzusaB200,ngAzusaB200aTent,ngSpacePark1,ngSpacePark2,ngSpacePark3ddd]
