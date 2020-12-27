% Analysis Class for PMU Calibration ETS step test
classdef AnalysisETS
    %%
    properties
        timeStep
        data
        parameters        
    end
    
    %%
    methods(Access=public)       
        
        %create a single contigous timestamp vector from the first set of
        %DUT timestmps
        function self = EtsTimestamps (self,DUT,timeStep)
            self.timeStep = timeStep;
                self.data.Timestamp(:,1)= DUT.data(1).Timestamp;
            for i = 2:length(DUT.data)
                self.data.Timestamp(:,i)=self.data.Timestamp(:,1)+((i-1)*timeStep);
            end
            dim = size(self.data.Timestamp); len=dim(1)*dim(2);
            self.data.Timestamp = reshape(self.data.Timestamp',[len 1]);
            %plot(self.data.Timestamp)
        end
        
        % interleave the DUT data
        function self = InterleaveData(self,DUT)
            len = length(DUT.data);
            % interleave results
            fnRes = fieldnames(DUT.data(1).Results);
            for i = 1:numel(fnRes)
                disp(fnRes{i})
                if fnRes{i} ~= "FE" && fnRes{i} ~= "RFE"
                    fnType = fieldnames(DUT.data(1).Results.(fnRes{i}));
                    for ii = 1:numel(fnType)
                        M = zeros(length(DUT.data(1).Results.(fnRes{i}).(fnType{ii})),len);
                        for j = 1:len
                            M(:,j) = DUT.data(j).Results.(fnRes{i}).(fnType{ii});
                        end
                        dim = size(M); dim = dim(1)*dim(2);
                        self.data.Results.(fnRes{i}).(fnType{ii})=reshape(M',[dim 1]);
                    end
                else
                    M = zeros(length(DUT.data(1).Results.(fnRes{i})),len);
                    for j = 1:len
                        M(:,j) = DUT.data(j).Results.(fnRes{i});
                    end
                    dim = size(M); dim = dim(1)*dim(2);
                    self.data.Results.(fnRes{i})=reshape(M',[dim 1]);
                end
                
            end
            
            % interleave PMU reports
            fnRep = fieldnames(DUT.data(1).PMU);
            for i = 1:numel(fnRep)
                disp(fnRep)
                M = zeros(length(DUT.data(i).PMU.(fnRep{i})),len);
                for j = 1:len
                    M(:,j) =DUT.data(j).PMU.(fnRep{i});
                end
                dim = size(M); dim = dim(1)*dim(2);
                self.data.PMU.(fnRep{i})=reshape(M',[dim 1]);
            end
            
            % interleave reference values 
            fnRep = fieldnames(DUT.data(1).REF);
            for i = 1:numel(fnRep)
                disp(fnRep)
                M = zeros(length(DUT.data(i).REF.(fnRep{i})),len);
                for j = 1:len
                    M(:,j) =DUT.data(j).REF.(fnRep{i});
                end
                dim = size(M); dim = dim(1)*dim(2);
                self.data.REF.(fnRep{i})=reshape(M',[dim 1]);
            end            
        end
 %% Save interleaved data to a .csv file named for the first file in the 
 % DUT's data folder
        function SaveInterleave(self, DUT)          
            % Make a table of the timestamps
            T = array2table(self.data.Timestamp);
            T.Properties.VariableNames= {'Timestamp'};
            
            % Make a table of the results
            t = table();    %initializa an empty table 
            fnPhase = fieldnames(self.data.Results);
            for i = 1:numel(fnPhase)
                if fnPhase{i} ~= "FE" && fnPhase{i} ~= "RFE"
                    fnRes = fieldnames(self.data.Results.(fnPhase{i}));
                    for ii = 1:numel(fnRes)
                        temp = array2table(self.data.Results.(fnPhase{i}).(fnRes{ii}));  
                        temp.Properties.VariableNames = cellstr(strcat((fnPhase{i}),'_',(fnRes{ii})));
                        t = [t, temp];
                    end                                   
                else
                    temp = array2table(self.data.Results.(fnPhase{i}));
                    temp.Properties.VariableNames = cellstr(fnPhase{i});
                    t = [t, temp];
                end
            end
            T = [T, t];
                                    
            % Table of PMU values
            T_pmu =  struct2table(self.data.PMU);
            vNames = T_pmu.Properties.VariableNames;
            for i = 1:numel(vNames)
                vNames{i} = strcat('PMU_', vNames{i});
            end
            T_pmu.Properties.VariableNames = vNames;
            
            % Table of REF values
            T_ref =  struct2table(self.data.REF);
            vNames = T_ref.Properties.VariableNames;
            for i = 1:numel(vNames)
                vNames{i} = strcat('REF_', vNames{i});
            end
            T_ref.Properties.VariableNames = vNames;
            
            % write  the table to a .csv file
            T = [T,T_pmu,T_ref];
            path = DUT.dataPath;
            names = dir(path);
            
            % search for the first non directory file 
            for i = 1:length(names)
               if ~names(i).isdir 
                   [~,name,ext]=fileparts(fullfile(names(i).folder,names(i).name));
                   if ext == ".csv"
                       if ~(contains(name,'_Parameters'))
                           break
                       end
                   end
               end
            end
            name = strcat(path,'\ETS_', name, '.csv');
            writetable(T,name)
        end
        
 %%   
        
            
                    
    end
    
end


        