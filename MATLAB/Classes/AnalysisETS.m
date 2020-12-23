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
        function self = InterlaceData(self,DUT)            
            len = length(DUT.data);
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
        end
    
            
                    
    end
    
end


        