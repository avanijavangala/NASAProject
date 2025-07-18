import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
df = pd.readcsv('NasaFile.xlsx')

class   RadiationTesting:


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




EMCAndEMIFacilities = [daytonTBrownEMIroom1,daytonTBrownEMIroom2, daytonTBrownEMIroom3, daytonTBrownEMIroom4, daytonTBrownEMIroom5, daytonTBrownEMIroom6,ngAzusaB183,ngAzusaB200,ngAzusaB200aTent,ngSpacePark1,ngSpacePark2,ngSpacePark3]


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

daytonTBrownVerticalHorizontalMips = shockFacility("USA", "3x3 m", "Pyro", "100 kg", "10000 G", "unknown", "20000 kN", "unknown")
daytonTBrownLansmont = shockFacility("USA", "1x1 m", "Free Fall", "50 kg", "5000–10000 G", "10–20 mm", "unknown", "20000 Hz")
daytonTBrownMontery = shockFacility("USA", "3x3 m", "Free Fall", "1000 kg", "500 G", "10 mm", "unknown", "20000 Hz")
eliteElectronicEngineering = shockFacility("USA", "2.5 m", "classical, free fall", "2993 kg", "30000 G", "unknown", "900 kN", "unknown")
ngAzusa = shockFacility("USA", "24x24 in", "Pendulum", "100 lbs", "5000 G", "N/A", "N/A", "10 kHz")
ngSpaceParkPendulum = shockFacility("USA", '12"x12" small, 24"x24" mid, 27"x20" large', "Pendulum", "~120 lbs", "8000 Gpk", "N/A", "N/A", "20 Hz – 10 kHz")
ngSpaceParkBungee = shockFacility("USA", '6"x6"', "Bungee", "<15 lbs", "8000 Gpk", "N/A", "N/A", "20 Hz – 10 kHz")
universityTartuObservatory = shockFacility("Observatooriumi 1, Tõravere, 61602, Estonia", "1x1 m", "mechanical impact (metal-to-metal pendulum hammer system)", "26 kg", "30000 G SRS", "N/A", "N/A", "10–10000 Hz")

shockTestingFacilities = [daytonTBrownVerticalHorizontalMips, daytonTBrownLansmont, daytonTBrownMontery, eliteElectronicEngineering, ngAzusa, ngSpaceParkPendulum, ngSpaceParkBungee, universityTartuObservatory]

class acousticFacility:
    def __init__(self, location, dimensions, testVolume, soundPressure, frequencyRange):
        self.location = location
        self.dimensions=dimensions
        self.testVolume = testVolume
        self.soundPressure = soundPressure
        self.frequencyRange = frequencyRange
    
    def getLocation(self):
        return location

    def getDimensions(self):
        return dimensions

    def getTestVolume(self):
        return testVolume

    def getSoundPressure(self):
        return soundPressure
    
    def getFrequencyRange(self):
        return frequencyRange

daytonTBrown = acousiticFacility("USA",	"40x40m", "unknown", "146 dB",	"20000 Hz")
ngAzusaDirectFieldAcousticFacility = acousticFacility("USA", "scalable system", "N/A", "147 dB", "10 kHz")
ngSpacePark = acousticFacility("USA", "7.9 x 9.7 x 18.9 m", "1448 m3", "154 dB", "20 Hz to 10 kHz")

acousticTestingFacilities = [daytonTBrown, ngAzusaDirectFieldAcousticFacility, ngSpacePark]

class centrifugeFacility:
    def __init__(self, location, armRadius, tableSize, maxPayloadMass, maxAcceleration, numberOfSLipRings, numberOfElectricalContacts):
        self.location=location
        self.armRadius = armRadius
        self.tableSize = tableSize
        self.maxPayloadMass = maxPayloadMass
        self.maxAcceleration = maxAcceleration
        self.numberOfSLipRings = numberOfSLipRings
        self.numberOfElectricalContacts = numberOfElectricalContacts

    def getLocation(self):
        return self.location

    def getArmRadius(self):
        return self.armRadius

    def getTableSize(self):
        return self.tableSize

    def getMaxPayloadMass(self):
        return self.maxPayloadMass

    def getMaxAcceleration(self):
        return self.maxAcceleration

    def getNumberOfSlipRings(self):
        return self.numberOfSlipRings

    def getNumberOfElectricalContacts(self):
        return self.numberOfElectricalContacts

daytonTBrown = centrifugeFacility("USA", "32 m", "3x3 m", "2000 kg", "200 G", "60", "unknown")
eliteElectronicEngineering = centrifugeFacility("USA", "0.9 m", "0.6x0.6 m", "unknown", "50 G", "28", "unknown")
istanbulTechnicalUniversity = centrifugeFacility("Turkey", "1.5 m", "unknown", "50 kg", "1–50 G", "unknown", "unknown")

centrifugeTestingFacilities = [daytonTBrown, eliteElectronicEngineering, instanbulTechnicalUniversity]



class RadiationTesting:
    def __init__(self, facilityName, location, testType, sourceOrParticles, doseRate, energy, contactEmail, webPage):
        self.facilityName = facilityName
        self.location = location
        self.testType = testType
        self.sourceOrParticles = sourceOrParticles
        self.doseRate = doseRate
        self.energy = energy
        self.contactEmail = contactEmail
        self.webPage = webPage

    def getFacilityName(self):
        return self.facilityName
    def getLocation(self):
        return self.location
    def getTestType(self):
        return self.testType
    def getSourceOrParticles(self):
        return self.sourceOrParticles
    def getDoseRate(self):
        return self.doseRate
    def getEnergy(self):
        return self.energy
    def getContactEmail(self):
        return self.contactEmail
    def getWebPage(self):
        return self.webPage
    boeing_gamma_source_g =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "TID", "Gamma Source G", "220", "~1.25", "arthur.a.rugtvedt@boeing.com")
boeing_gamma_source_e =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "TID", "Gamma Source E", "40", "~1.25", "arthur.a.rugtvedt@boeing.com")
boeing_gamma_source_s =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "TID", "Gamma Source S", "1", "~1.25", "arthur.a.rugtvedt@boeing.com")
boeing_open_field_gamma_irradiator =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "TID", "Open Field Gamma Irradiator", "0.1", "~1.25", "arthur.a.rugtvedt@boeing.com")

boeing_dynamitron =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "SEE", "Dynamitron", "", "2.5", "arthur.a.rugtvedt@boeing.com")
boeing_cretch =   RadiationTesting("Boeing Radiation Effects Lab", "USA", "SEE", "CRETCh", "", "0.75", "arthur.a.rugtvedt@boeing.com")

northrop_pulsed_laser =   RadiationTesting("Northrop Grumman SEE - Pulsed Laser", "USA", "Single Event Effects", "Laser", "", "", "jonathan.avila2@ngc.com, jeffrey.warner@ngc.com")
northrop_beacon_irradiator_1 =   RadiationTesting("Northrop Grumman ELDRS - Beacon Room Irradiator 1", "USA", "Total Ionizing Dose - Enhanced Low Dose Rate Sensitivity", "Co60 Gamma Rays", "≤0.01", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")
northrop_beacon_irradiator_2 =   RadiationTesting("Northrop Grumman ELDRS - Beacon Room Irradiator 2", "USA", "Total Ionizing Dose - Enhanced Low Dose Rate Sensitivity", "Co60 Gamma Rays", "≤0.01", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")
northrop_hopewell_ds20 =   RadiationTesting("Northrop Grumman ELDRS - Hopewell DS20 Room Irradiator", "USA", "Total Ionizing Dose - Enhanced Low Dose Rate Sensitivity", "Co60 Gamma Rays", "≤0.01", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")
northrop_shepherd_142 =   RadiationTesting("Northrop Grumman ELDRS - Shepherd 142 Irradiator", "USA", "Total Ionizing Dose - Enhanced Low Dose Rate Sensitivity", "Co60 Gamma Rays", "≤0.01", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")

northrop_gammacell_irradiator_1 =   RadiationTesting("Northrop Grumman TID HDR - Gammacell 220 Irradiator #1", "USA", "Total Ionizing Dose - Enhanced Low Dose Rate Sensitivity", "Co60 Gamma Rays", "<0.15", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")
northrop_gammacell_irradiator_3 =   RadiationTesting("Northrop Grumman TID HDR - Gammacell 220 Irradiator #3", "USA", "Total Ionizing Dose - High Dose Rate", "Co60 Gamma Rays", "50-300", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")
northrop_gammacell_irradiator_4 =   RadiationTesting("Northrop Grumman TID HDR - Gammacell 220 Irradiator #4", "USA", "Total Ionizing Dose - High Dose Rate", "Co60 Gamma Rays", "50-300", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")

northrop_fxr_febetron_705 =   RadiationTesting("USA", "Prompt Dose (Nuclear Hardness & Survivability)", "X-Rays", "1E7-1E11 @ 22ns PWHM", "", "jonathan.avila2@ngc.com, jared.myers@ngc.com")

radef_heavy_ions =   RadiationTesting("Finland", "Heavy ions", "K-130 cyclotron, from B- Au", "from 5 up to coulpes of 1e5 ions/cm^2/s", "10, 16.3 and 22 MeV/n", "Heikki.i.Kettunen@jyu.fi, https://www.jyu.fi/accelerator/radef")

radef_protons =   RadiationTesting("Finland", "Protons", "K-130 cyclotron, protons", "from 1e4 up to coulpes of 1e9 protons/cm^2/s", "from 0.4 up to 55 MeV", "Heikki.i.Kettunen@jyu.fi, https://www.jyu.fi/accelerator/radef")

radef_gamma_rays =   RadiationTesting("Finland", "Gammarays", "Electron accelerator (pulssed photon beam)", "100-600 Rad /min = (1.7-10 Rad/s)", "6 MV and 15 MV Bremsstrahlung radiation", "Heikki.i.Kettunen@jyu.fi, https://www.jyu.fi/accelerator/radef")

radef_electrons =   RadiationTesting("Finland", "Electrons", "Electron accelerator (pulssed electron beam)", "100-1000 Rad /min = (1.7-17 Rad/s)", "6, 9, 12, 16 and 20 MeV", "Heikki.i.Kettunen@jyu.fi, https://www.jyu.fi/accelerator/radef")

class AltitudeChamber:
    def __init__(self, facilityName, location, dimensions, temperatureRange, maxAltitude, heatingCoolingRates, contactEmail):
        self.facilityName = facilityName
        self.location = location
        self.dimensions = dimensions
        self.temperatureRange = temperatureRange
        self.maxAltitude = maxAltitude
        self.heatingCoolingRates = heatingCoolingRates
        self.contactEmail = contactEmail

    def getFacilityName(self):
        return self.facilityName
    def getLocation(self):
        return self.location
    def getDimensions(self):
        return self.dimensions
    def getTemperatureRange(self):
        return self.temperatureRange
    def getMaxAltitude(self):
        return self.maxAltitude
    def getHeatingCoolingRates(self):
        return self.heatingCoolingRates
    def getContactEmail(self):
        return self.contactEmail
dayton_20ft = AltitudeChamber("USA", "6x2.4", "-40 to 100", "22.8", "2", "mmay@dtb.com")
dayton_6ft = AltitudeChamber("USA", "1.2x1.8", "-65 to 157", "22.8", "5", "mmay@dtb.com")
dayton_4ft = AltitudeChamber("USA", "1.2x1.2", "-65 to 158", "22.8", "5", "mmay@dtb.com")
dayton_tvac = AltitudeChamber("USA", "1.5x1.5", "-100 to 156", "10-6torr", "2", "mmay@dtb.com")

class SensorsTesting:
    def __init__(self, facilityName, location, testType, equipment, wavelengthRange, contactEmail):
        self.facilityName = facilityName
        self.location = location
        self.testType = testType
        self.equipment = equipment
        self.wavelengthRange = wavelengthRange
        self.contactEmail = contactEmail

    def getFacilityName(self):
        return self.facilityName
    def getLocation(self):
        return self.location
    def getTestType(self):
        return self.testType
    def getEquipment(self):
        return self.equipment
    def getWavelengthRange(self):
        return self.wavelengthRange
    def getContactEmail(self):
        return self.contactEmail
elite_electronic_engineering = SensorsTesting("USA", "EMI, environmental stress", "Chambers", "", "michael.cosentino@elitetest.com")

airborne_sensor_facility = SensorsTesting("USA", "radiometry/ spectral sensitivity", "eMAS, PICARD, MASTER, AMS, DMS, DCS, POS", "350 – 14000", "")

