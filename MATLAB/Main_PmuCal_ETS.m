% Analyse ETS data

clear

DUT = CcsInfo;
DUT.getDataPath('Open ETS Data folder')
DUT.readData;
DUT.readParams;

ETS = AnalysisETS;
ETS = ETS.EtsTimestamps(DUT,0.002);
ETS = ETS.InterleaveData(DUT);

% create a time vector
for i = 1:length(DUT.data(10).Timestamp)
    x = DUT.data(10).Timestamp(i)