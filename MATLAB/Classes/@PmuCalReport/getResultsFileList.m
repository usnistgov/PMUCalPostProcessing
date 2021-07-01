function self = getResultsFileList(self)
%getResultsFileList Prompts the user for the results file path
% Prompts the user for the path to the PMU Calibrator results files then gets  lists of results and parameter files

prompt = sprintf('Choose the folder of all results from F0 = %d, Fs = %d', self.F0, self.Fs);
rawDataPath = uigetdir(self.resultPath,prompt);
paramNames = {};
dataNames = self.getfn(rawDataPath,'.csv');
%for i = 1:numel(dataNames)
i = 1;
while i <= numel(dataNames)
    if contains(dataNames(i),'Parameters')
        paramNames{end+1}=dataNames(i);
        dataNames(i) = [];
    else
        i = i+1;
    end
end
self.paramFiles = paramNames;
self.dataFiles = dataNames;

% getting the specifications for the selected PMU configuration
absPath = mfilename('fullpath');
absPath = extractBefore(absPath,'\MATLAB');
relPath = sprintf('Spec_%dF0_%dFs_%s.csv',self.F0,self.Fs,self.PmuClass);
absPath = fullfile(absPath,'Specifications',relPath);
try
    self.specifications = readtable(absPath);
catch
    error ('Failed to read  specifications file %s,',absPath);
end

% Dialog to verify the correct PMU configuration
dlg = dlgConfig(self, num2str(self.F0), num2str(self.Fs), self.PmuClass, self.resultPath);
waitfor(dlg)

% TODO: Update the .ini file with the latest Raw Data Path

end