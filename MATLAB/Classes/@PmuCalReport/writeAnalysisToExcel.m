function writeAnalysisToExcel(self)
%Loop through all the data and param files, creating and adding data to sheets

i = 1;
nextLine = 1;
while (i <= numel(self.dataFiles))
    [sheetName,influenceFactor] = self.makeSheetName(self.paramFiles{1,i});
    [new] = self.activeSheet(sheetName);           % create a new sheet or set active
    
    % Most test types just have one line but step test is special
    iFactor = influenceFactor.Properties.VariableNames{1,1};
    switch iFactor
        case {'KxS','KaS'}
            [i] = writeStepResultsToExcel(self,i,influenceFactor);
            
        otherwise
            [nextLine] = writeResultsToExcel (self,nextLine,i,new,influenceFactor);
            
    end
    i = i+1;
end
self.hExcel.Save;
end
