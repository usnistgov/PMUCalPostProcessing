% Calculates the uncertainty of a PMU calibration system using standard
% IEEE ICAP output formats of raw data from a PMU calibrator.
%
% The "CUT" is the PMU Calibrator Under Test
% The "CCS" is the Calibrator Calibration System.
%
clear

% Create a Calibrator Under Test object
CUT = CalInfo;  % create an instance of the CalInfo class
CUT.getDataPath('Open CUT Data folder'); 
CUT.readData;

% Create a Calobrator Clibration System object
CCS = CcsInfo;  % create an instance of the CcsInfo class
CCS.getDataPath('Open CCS Data folder'); 
CCS.readData;
CCS.readParams;



ANALYSIS = AnalysePmuCalCal;
[CUT, CCS] = ANALYSIS.alignData(CUT, CCS);
ANALYSIS = ANALYSIS.calcCutCcsDiffs(CUT, CCS);
ANALYSIS = ANALYSIS.comparePmuResults(CUT,CCS);


        
    