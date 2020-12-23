% Class for the PMU Calibrator Calibration System (CCS)
% Child of the CalInfo Class, adds the Parameters

classdef CcsInfo < CalInfo
    
    properties
        Parameters
        eTestType
        F0
    end
    
    %% Public Methods
    methods(Access=public)
        
        % Read all the  params files from the folder,
        function readParams(self)
            [~,paramFiles] = self.getFileNames();
            for i = 1 : length(paramFiles)
                C = readcell(fullfile(paramFiles(i).folder,paramFiles(i).name));
                C = C';
                self.Parameters(i).eTestType = C{2,1};
                self.Parameters(i).F0 = C{2,2};
                T = cell2table(C); 
                T = T(:,3:end);
                T.Properties.VariableNames = C(1,3:end);
                T = T(2:end,:);
                fn = T(1,:);
                for j = 1:numel(fn)
                    self.Parameters(i).Params=T;
                end
            end
        end
    end
end