function self = makeReportFileName(self)
% creates the file name for the reports file once the
% resultFileLists are populated

C = readcell(cell2mat(self.paramFiles{1}));
C = C';
T = cell2table(C(2:end,:));
T.Properties.VariableNames = C(1,:);

if size(T,1) > 2
    self.ResultType = 'NIST';
    name = sprintf('%dF0_%dFs_%s.xlsx',T.F0(1),self.Fs,self.PmuClass);
else
    self.ResultType = 'Fluke';
    name = sprintf('%dF0_%dFs_%s.xlsx',T.F0,T.Fs,self.PmuClass);
    self.Fs = T.Fs;
end
self.ReportFile = fullfile(self.resultPath,'Reports',name);
end

