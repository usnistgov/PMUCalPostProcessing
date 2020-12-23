% Class for PMU Calibrator results
% Calculates the uncertainty of a PMU calibration system using standard
% IEEE ICAP output formats of raw data from a PMU calibrator.
%
% The "CUT" is the PMU Calibrator Under Test
% The "CCS" is the Calibrator Calibration System.
%
classdef CalInfo < handle
    
    properties
        dataPath
        data
    end
%% Public Methods    
    methods(Access=public)
        % Ask the user for the path to the (CUT or CCS) data
        function getDataPath(self,title)
            self.dataPath = uigetdir('.',title);
        end
        
        % Read all the files in the path and initialize the data property
        function readData(self)
            [dataFiles, ~] = self.getFileNames();
            %self.data = zeros(length(dataFiles),1);
            for i = 1:length(dataFiles)
                self.getData(i,dataFiles(i));
            end        
        end
        
    end
    
%% Protected Methods
    methods(Access=protected)
        % gets the file info for the data and parameters files
        function [dataFiles, paramFiles] = getFileNames(self)
            names = dir(self.dataPath);
            ii = 1; jj = 1;     %init a file pointer
            for i = 1:length(names)
                if ~names(i).isdir
                    [~,name,ext]=fileparts(fullfile(names(i).folder,names(i).name));
                    if ext == ".csv"
                        if contains(name,'_Parameters')
                            paramFiles(ii) = names(i); ii = ii+1;
                        else
                            dataFiles(jj) = names(i); jj = jj+1;
                        end
                    end
                end
            end
        end
        
        % gets one result structure from one data file
        function self = getData(self,i,file)
            M = readmatrix(fullfile(file.folder,file.name));
           
            self.data(i).Timestamp = M(:,1);
            % TVE, ME, PE
            self.data(i).Results.VA = self.getResult(M,2);
            self.data(i).Results.VB = self.getResult(M,5);
            self.data(i).Results.VC = self.getResult(M,8);
            self.data(i).Results.Vp = self.getResult(M,11);
            self.data(i).Results.IA = self.getResult(M,14);
            self.data(i).Results.IB = self.getResult(M,17);
            self.data(i).Results.IC = self.getResult(M,20);
            self.data(i).Results.Ip = self.getResult(M,23);
            % FE, RFE
            self.data(i).Results.FE =M(:,26);
            self.data(i).Results.RFE =M(:,27);
            % PMU Raw data            
            self.data(i).PMU.VA = self.getComplex(M,29);
            self.data(i).PMU.VB = self.getComplex(M,31);
            self.data(i).PMU.VC = self.getComplex(M,33);
            self.data(i).PMU.Vp = self.getComplex(M,35);
            self.data(i).PMU.IA = self.getComplex(M,37);
            self.data(i).PMU.IB = self.getComplex(M,39);
            self.data(i).PMU.IC = self.getComplex(M,41);
            self.data(i).PMU.Ip = self.getComplex(M,43);
            self.data(i).PMU.Freq = M(:,45);
            self.data(i).PMU.ROCOF = M(:,46);
            % Reference Raw Data
            self.data(i).REF.VA = self.getComplex(M,48);
            self.data(i).REF.VB = self.getComplex(M,50);
            self.data(i).REF.VC = self.getComplex(M,52);
            self.data(i).REF.Vp = self.getComplex(M,54);
            self.data(i).REF.IA = self.getComplex(M,56);
            self.data(i).REF.IB = self.getComplex(M,58);
            self.data(i).REF.IC = self.getComplex(M,60);
            self.data(i).REF.Ip = self.getComplex(M,62);
            self.data(i).REF.Freq = M(:,64);
            self.data(i).REF.ROCOF = M(:,65);           
        end
        
        % this gets the test result (TVE, ME, PE) section from the data files
        function result = getResult(self,M,i)
            result.TVE= M(:,i);
            result.ME = M(:,i+1);
            result.PE = M(:,i+2);
        end
        
        %this gets the PMU and reference complex numbers 
        function result = getComplex(self,M,i)
            result = M(:,i);
            result = result + 1i*M(:,i+1);
        end
            
        
    end
    
    
end


