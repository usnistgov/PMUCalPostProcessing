function plotExcelAnalysis(self)

nSheets = self.hExcel.hSheets.Count;

for i = nSheets:-1:1
    self.hExcel.ActivateSheet(i);
    sheetName = self.hExcel.hSheet.Name;
    T = readtable(self.ReportFile,'Sheet',sheetName,'PreserveVariableNames',true);
    iFactor = T.Properties.VariableNames;
    
    [T,xLabel] = getXLabel(T,self.F0);
    
    % Get the pass/fail limits
    Lim = self.getLimitsFromSheetName(sheetName);
    
    % Create a new sheet
    self.hExcel.NewSheet(strcat(sheetName,' plots'))

    
    switch sheetName
        case {'phase step pos','phase step neg','amplitude step pos','amplitude step neg'}
            plotStepResults(self,T,Lim,xLabel)            
            
        otherwise
            plotResults(self,T,Lim,xLabel)
            
    end
        
    
    
end
end

% --------------------------------------
function [T,xLabel] = getXLabel(T, F0)
xLabel = strings(2,1);
iFactor = T.Properties.VariableNames;

switch iFactor{1}
    case 'Fin'
        xLabel(:) = "Frequency (Hz)";
    case 'Kv'
        xLabel(1) = "Voltage Index (% of nominal)";
        xLabel(2) = "Current Index (% of nominal)";
    case 'Fh'
        xLabel(:) = "Interfering Signal Frequency (Hz)";
        % if it is interharmonics, we want to find the gap
        % on either side of the nominal frequency
        mask = T.Fh<F0;
        tNaN = NaN(1,width(T));
        tNaN = array2table(tNaN);
        tNaN.Properties.VariableNames = T.Properties.VariableNames;
        T = [T(mask,:);tNaN;T(~mask,:)];
    case {'Fa' 'Fx'}
        xLabel(:) = "Modulation Frequency (Hz)";
        
    otherwise
        xLabel(:)= "Time (s)";
end
end
