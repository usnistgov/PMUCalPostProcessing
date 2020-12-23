% Analyse ETS data

clear

DUT = CcsInfo;
DUT.getDataPath('Open ETS Data folder')
DUT.readData;
DUT.readParams;

ETS = AnalysisETS;
ETS = ETS.EtsTimestamps(DUT,0.002);
ETS = ETS.InterlaceData(DUT);