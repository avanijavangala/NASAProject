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
#dimensions are in meters
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


class vibrationFacility:
    def __init__(self, location, slipTabledimensions,headExpanderDimensions, testType, maxForce, randomVibrationForce, frequencyRange, peakToPeakMaxDisplacement,maxBareTableAcceleration, maxLoadCapacity):
        self.location = location
        self.slipTabledimensions = slipTabledimensions
        self.headExpanderDimensions = headExpanderDimensions
        self.testType = testType
        self.maxForce = maxForce
        self.randomVibrationForce = randomVibrationForce
        self.frequencyRange = frequencyRange
        self.peakToPeakMaxDisplacement = peakToPeakMaxDisplacement
        self.maxBareTableAcceleration = maxBareTableAcceleration
        self.maxLoadCapacity = maxLoadCapacity


    def getSlipTableDimensions(self):
        return self.slipTabledimensions
    def getHeadExpanderDimensions(self):
        return self.headExpanderDimensions
    def getTestType(self):
        return self.testType
    def getMaxForce(self):
        return self.maxForce
    def getRandomVibrationForce(self):
        return self.randomVibrationForce
    def getFrequencyRange(self):
        return self.frequencyRange
    def getPeakToPeakMaxDisplacement(self):
        return self.peakToPeakMaxDisplacement
    def getMaxBareTableAcceleration(self):
        return self.maxBareTableAcceleration
    def getMaxLoadCapacity(self):
        return self.maxLoadCapacity

daytonTBrown1 = vibrationFacility("USA", "36x36 m", "36x36 m", "any", "178 kN", "89 kN", "3–3000 Hz", "2 mm", "200 G", "500 kg")
daytonTBrown2 = vibrationFacility("USA", "36x36 m", "36x36 m", "Any", "178 kN", "89 kN", "3–3000 Hz", "1 mm", "200 G", "500 kg")
daytonTBrown3 = vibrationFacility("USA", "36x36 m", "36x36 m", "Any", "212 kN", "106 kN", "3–3000 Hz", "3 mm", "200 G", "500 kg")
daytonTBrown4 = vibrationFacility("USA", "any size", "32x96 m", "Any", "354 kN", "177 kN", "3–3000 Hz", "2 mm", "200 G", "10000 kg")
daytonTBrown5 = vibrationFacility("USA", "60x60 m", "32x96 and 60x60 m", "Any", "133–490 kN", "245 kN", "0.1–3000 Hz", "3–10.5 mm", "50–200 G", "5000–10000 kg")
daytonTBrown6 = vibrationFacility("USA", "36x36 m", "36x36 m", "Any", "66 kN", "66 kN", "5–2300 Hz", "2 mm", "100 G", "500 kg")
eliteElectronicEngineering = vibrationFacility("USA", "2.4x2.4 m", "1.8x2.4 m", "sine, random", "347 kN", "unknown", "unknown", "unknown", "220 G", "2993 kg")
ngAzusa1 = vibrationFacility("USA", "48x48 in", "36x36 in", "All", "20000 F-lbs", "18000 F-lbs", "4–2500 Hz", "2 in", "50 G", "1000 lbs")
ngAzusa2 = vibrationFacility("USA", "48x48 in", "36x36 in", "All", "20000 F-lbs", "18000 F-lbs", "4–2500 Hz", "2 in", "50 G", "1000 lbs")
ngAzusa3 = vibrationFacility("USA", "60x60 in", "60x60 in", "All", "50000 F-lbs", "48000 F-lbs", "4–2500 Hz", "3 in", "100 G", "4000 lbs")
ngSpaceParkT5500SCIF = vibrationFacility("USA", "120x119 in (JWST) or custom", "116 in diam (JWST) or custom", "All", "55000 F-lbs (sine)", "48000 F-lbs (random)", "5–2000 Hz", "2 in", "100 G", "4000 lbs")
ngSpaceParkT4000 = vibrationFacility("USA", "75x73.5 in", "63 in diam", "All", "40000 F-lbs (sine)", "36000 F-lbs (random)", "5–2000 Hz", "2 in", "100 G", "4000 lbs")
ngSpaceParkT1000 = vibrationFacility("USA", "33x33 in", "25 in diam", "All", "21500 F-lbs (sine)", "18000 F-lbs (random)", "5–2000 Hz", "2 in", "100 G", "1500 lbs")
ngSpaceParkC150 = vibrationFacility("USA", "33x33 in", "25 in diam", "All", "17000 F-lbs (sine)", "15000 F-lbs (random)", "5–2000 Hz", "1 in", "100 G", "1000 lbs")
universityTartuObservatory = vibrationFacility("Observatooriumi 1, Tõravere, 61602, Estonia", "No slip table", "16 cm", "Sinusoidal vibration, Random vibration", "1.62 kN", "1.09 kN", "5–4000 Hz", "25.4 mm", "73 G", "50 kg")

vibrationTestFacilities = [daytonTBrown1, daytonTBrown2, daytonTBrown3, daytonTBrown4, daytonTBrown5, daytonTBrown6, eliteElectronicEngineering, ngAzusa1, ngAzusa2, ngAzusa3, ngSpaceParkT5500SCIF, ngSpaceParkT4000, ngSpaceParkT1000, ngSpaceParkC150, universityTartuObservatory]


class shockFacility:
    def __init__(self, location, tableSize, testType, maxWeight, maxAcceleration, maxDisplacement, force, operatingFrequencyRange):

        self.location = location
        self.tableSize = tableSize
        self.testType = testType
        self.maxWeight = maxWeight
        self.maxAcceleration = maxAcceleration
        self.maxDisplacement = maxDisplacement
        self.force = force
        self.operatingFrequencyRange = operatingFrequencyRange

    def getLocation(self):
        return self.location

    def getTableSize(self):
        return self.tableSize

    def getTestType(self):
        return self.testType

    def getMaxWeight(self):
        return self.maxWeight

    def getMaxAcceleration(self):
        return self.maxAcceleration

    def getMaxDisplacement(self):
        return self.maxDisplacement

    def getForce(self):
        return self.force

    def getOperatingFrequencyRange(self):
        return self.operatingFrequencyRange

daytonTBrownVerticalHorizontalMips = vibrationFacility("USA", "3x3 m", "Pyro", "100 kg", "10000 G", "unknown", "20000 kN", "unknown")
daytonTBrownLansmont = vibrationFacility("USA", "1x1 m", "Free Fall", "50 kg", "5000–10000 G", "10–20 mm", "unknown", "20000 Hz")
daytonTBrownMontery = vibrationFacility("USA", "3x3 m", "Free Fall", "1000 kg", "500 G", "10 mm", "unknown", "20000 Hz")
eliteElectronicEngineering = vibrationFacility("USA", "2.5 m", "classical, free fall", "2993 kg", "30000 G", "unknown", "900 kN", "unknown")
ngAzusa = vibrationFacility("USA", "24x24 in", "Pendulum", "100 lbs", "5000 G", "N/A", "N/A", "10 kHz")
ngSpaceParkPendulum = vibrationFacility("USA", '12"x12" small, 24"x24" mid, 27"x20" large', "Pendulum", "~120 lbs", "8000 Gpk", "N/A", "N/A", "20 Hz – 10 kHz")
ngSpaceParkBungee = vibrationFacility("USA", '6"x6"', "Bungee", "<15 lbs", "8000 Gpk", "N/A", "N/A", "20 Hz – 10 kHz")
universityTartuObservatory = vibrationFacility("Observatooriumi 1, Tõravere, 61602, Estonia", "1x1 m", "mechanical impact (metal-to-metal pendulum hammer system)", "26 kg", "30000 G SRS", "N/A", "N/A", "10–10000 Hz")

shockTestingFacilities = [daytonTBrownVerticalHorizontalMips, daytonTBrownLansmont, daytonTBrownMontery, eliteElectronicEngineering, ngAzusa, ngSpaceParkPendulum, ngSpaceParkBungee, universityTartuObservatory]