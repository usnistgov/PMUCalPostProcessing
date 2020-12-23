% Analysis class for PMU Calibrator Calibration
classdef AnalysePmuCalCal
%%    
   properties
       CutCcsDiffs  % TVE, FE, RFE of the CUT references minus the CCS references
       PmuResults   % Using the PMU values from the CUT, and the Reference values from the CCS, get the PMU calibration Recults
       PmuResultCompare  % Compare the PMU results from the CUT to the ones found in PmuResults above
   end
%%   
methods(Access=public)
    
    %  calculates the differences of the reference values in terms of %
    %  TVE, FE and RFE
    function self = calcCutCcsDiffs(self,CUT, CCS)
        % Loop through all the pages of data
        for i = 1 : length(CCS.data)
            
            % verify the timestamps are all the same and throw an error if not
            if ~isequal(CUT.data(i).Timestamp,CCS.data(i).Timestamp) error('Unequal timstamps found in CUT and CCS data on page %d', i); end
            
            %for all the data elements, subtract them, calculate he TVE, FE
            %or RFE as appropriate and store the results.
            fnRef = fieldnames(CCS.data(i).REF);
            self.CutCcsDiffs.data(i).Timestamp = CCS.data(i).Timestamp;
            % Loop through all the reference values
            for ii = 1:numel(fnRef)
                %disp(fnRef{ii});
                cut = CUT.data(i).REF.(fnRef{ii});
                ccs = CCS.data(i).REF.(fnRef{ii});
                
                if fnRef{ii} ~= "Freq" && fnRef{ii} ~= "ROCOF"
                    % Calculate the TVE, ME and PE for the reference phasor values
                    % TVE
                    self.CutCcsDiffs.data(i).(fnRef{ii}).TVE = sqrt(((real(cut)-real(ccs)).^2+(imag(cut)-imag(ccs)).^2)./(real(ccs).^2+imag(ccs).^2))*100;
                    % ME
                    self.CutCcsDiffs.data(i).(fnRef{ii}).ME = (abs(cut)-abs(ccs))./abs(ccs);
                    % PE in radians
                    self.CutCcsDiffs.data(i).(fnRef{ii}).PE = angle(cut)-angle(ccs);
                elseif fnRef{ii} == "Freq"
                    self.CutCcsDiffs.data(i).(fnRef{ii}).FE = cut - ccs;
                elseif  fnRef{ii} == "ROCOF"
                    self.CutCcsDiffs.data(i).(fnRef{ii}).RFE = cut - ccs;
                end
            end
        end
    end
    
    
    % Calculate the PMU results using the CUT's PMU values and the CCS's
    % reference values, then find the difference between those and the CUTs
    % results.  This checks both the CUTs reference values and the result calculations.
    % If the CUT properly calculates the results, these should match up with the calculated
    % differences in the reference values computed by calcCutCcsDiffs.
    function self = comparePmuResults(self, CCS, CUT)
        self = self.calcPmuResults(CCS,CUT);
        self = self.compareResults(CUT);
    end
    
end
 
    %%
    methods(Access = public)
        
        % Calculate the PMU results using the CUT's PMU values and the CCS's
        function self = calcPmuResults(self, CCS, CUT)
            % Loop through all the pages of data
            for i = 1 : length(CCS.data)
                
                % verify the timestamps are all the same and throw an error if not
                if ~isequal(CUT.data(i).Timestamp,CCS.data(i).Timestamp) error('Unequal timstamps found in CUT and CCS data on page %d', i); end
                
                %for all the data elements, subtract them, calculate he TVE, FE
                %or RFE as appropriate and store the results.
                fnRef = fieldnames(CUT.data(i).PMU);
                self.PmuResults.data(i).Timestamp = CUT.data(i).Timestamp;
                % Loop through all the reference values
                for ii = 1:numel(fnRef)
                    %disp(fnRef{ii});
                    pmu = CUT.data(i).PMU.(fnRef{ii});
                    ccs = CCS.data(i).REF.(fnRef{ii});
                    
                    if fnRef{ii} ~= "Freq" && fnRef{ii} ~= "ROCOF"
                        % Calculate the TVE, ME and PE for the reference phasor values
                        % TVE
                        self.PmuResults.data(i).(fnRef{ii}).TVE = sqrt(((real(pmu)-real(ccs)).^2+(imag(pmu)-imag(ccs)).^2)./(real(ccs).^2+imag(ccs).^2))*100;
                        % ME
                        self.PmuResults.data(i).(fnRef{ii}).ME = ((abs(pmu)-abs(ccs))./abs(ccs))*100;
                        % PE in radians
                        self.PmuResults.data(i).(fnRef{ii}).PE = angle(pmu)-angle(ccs);
                    elseif fnRef{ii} == "Freq"
                        self.PmuResults.data(i).FE = pmu - ccs;
                    elseif  fnRef{ii} == "ROCOF"
                        self.PmuResults.data(i).RFE = pmu - ccs;
                    end
                end
            end
        end
        
        % Find the difference between the calculated results in PmuResults
        % and the results in the CUT's Results
        function self = compareResults(self, CUT)
            % PmuResultCompare will hold the difference between the CUTs
            % calculated PMU results and the PMUResults calculated above
            for i = 1 : length(CUT.data)
                fnRes = fieldnames(CUT.data(i).Results);
                self.PmuResultCompare.data(i).Timestamp = CUT.data(i).Timestamp;
                % Loop through all the PMU values
                for ii = 1:numel(fnRes)
                    %disp(fnRes{ii});
                    if fnRes{ii} ~= "FE" && fnRes{ii} ~= "RFE"
                        fnVal = fieldnames(CUT.data(i).Results.(fnRes{ii}));
                        for j = 1:numel(fnVal)
                            %disp(fnVal{j});
                            cut = CUT.data(i).Results.(fnRes{ii}).(fnVal{j});
                            calc = self.PmuResults.data(i).(fnRes{ii}).(fnVal{j});
                            self.PmuResultCompare.data(i).(fnRes{ii}).(fnVal{j})=cut-calc;
                        end
                    else
                           cut = CUT.data(i).Results.(fnRes{ii});
                           calc = self.PmuResults.data(i).(fnRes{ii});
                           self.PmuResultCompare.data(i).(fnRes{ii})=cut-calc;                         
                    end
                end
                
            end
        end
        
    end
    

   
   %%
   methods(Static)
       
       % Static function finds the commmon timestamps in each page of data
       % and returns data with only the common timestamps
       function [CUT, CCS] =  alignData(CUT, CCS)
           for i = 1:length(CCS.data)
               [~,idx] = intersect(CCS.data(i).Timestamp,CUT.data(i).Timestamp,'rows');
               CCS.data(i).Timestamp = CCS.data(i).Timestamp(idx(1):idx(end));
               CUT.data(i).Timestamp = CUT.data(i).Timestamp(idx(1):idx(end));
               if isempty(idx) error('Error: no common timestamps were found on the %d th page during alignData', i); end
               
               %loop through all the fields of all the structures
               fnData = fieldnames(CCS.data(i));
               fnData = fnData(2:end);
               for ii = 1:numel(fnData)                   
                   %disp(fnData{ii})                   
                   %Loop through all the results structures
                   fnRes = fieldnames(CCS.data(i).(fnData{ii}));
                   for j = 1:numel(fnRes)
                       %disp(fnRes{j})
                       %loop through all the values in the results
                       if fnRes{j} ~= "FE" && fnRes{j} ~= "RFE" && fnData{ii} ~= "PMU" && fnData{ii} ~= "REF"
                           fnVal = fieldnames(CCS.data(i).(fnData{ii}).(fnRes{j}));
                           for jj = 1:numel(fnVal)
                               %disp(fnVal{jj})
                               CCS.data(i).(fnData{ii}).(fnRes{j}).(fnVal{jj}) = CCS.data(i).(fnData{ii}).(fnRes{j}).(fnVal{jj})(idx(1):idx(end));
                               CUT.data(i).(fnData{ii}).(fnRes{j}).(fnVal{jj}) = CUT.data(i).(fnData{ii}).(fnRes{j}).(fnVal{jj})(idx(1):idx(end));
                           end
                       else
                           CCS.data(i).(fnData{ii}).(fnRes{j}) = CCS.data(i).(fnData{ii}).(fnRes{j})(idx(1):idx(end));
                           CUT.data(i).(fnData{ii}).(fnRes{j}) = CUT.data(i).(fnData{ii}).(fnRes{j})(idx(1):idx(end));
                           
                       end
                   end
               end
           end
       end
       
   end
    
end