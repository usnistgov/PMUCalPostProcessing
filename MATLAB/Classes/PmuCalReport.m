% Report Generation Class for PMUCal Uncertainty Calculator

% The goal is to write data and plots directly into sheets of Excel
% spreadsheets that will later be used to populate MS word calibration
% reports.

classdef PmuCalReport < handle
    
    properties
        hExcel  % handle to the Excel activeX object
        resultPath
        paramFiles
        dataFiles
        PmuClass = 'M'
        ReportFile
        vNom = 70
        iNom = 5
        Fs = 50
        F0 = 50
        ResultType
        
        % figure default sizes
        figPos = [1664 1848 360 250];   % figure position
        axPos = [0.1 0.15 0.6 0.75];     % axis position
        lgdPos = [0.71 0.425 0.27 0.2475]; % legend position
        
        specifications %specifications read from .xml file for the current PMU configuration

        
        
    end
    
    %% Constructor
    methods
        
        function self = PmuCalReport(varargin)            
            appDataPath = fullfile(getenv('APPDATA'),'PmuCal');
            if ~exist(appDataPath, 'dir' )
                mkdir(appDataPath)
            end
            name = 'PmuCalReport.ini';
            
            % input arguments
            for i = 1:2:nargin
                switch varargin{i}
                    case 'vNom'
                        self.vNom = varargin{i+1};
                    case 'iNom'
                        self.iNom = varargin{i+1};
                    case 'Fs'
                        self.Fs = varargin{i+1};
                    case 'F0'
                        self.F0 = varargin{i+1};
                    case 'PmuClass'
                        self.PmuClass = varargin{i+1};
                    case 'Reset'    % delete the .ini file
                        b = varargin{i+1};
                        if b=="t" || b=="T"|| b=="true" || b=="True"
                            if exist(fullfile(appDataPath,name),'file')
                                delete(fullfile(appDataPath,name))
                            end
                        end
                    otherwise
                        warning('Unrecognized parameter %s',varargin{i});
                end
                
            end
            
            
            % program .ini file
            if ~exist(fullfile(appDataPath,name),'file')
                self.resultPath = uigetdir(fullfile(getenv('USERPROFILE'),'Documents'),'Path to PMU results');
                structure = struct('ResultsPath',struct('ResultsPath',self.resultPath));
                self.struct2Ini(fullfile(appDataPath,name),structure);
            else
                structure = self.ini2struct(fullfile(appDataPath,name));
                self.resultPath = structure.ResultsPath.ResultsPath;
            end
                                       
        end
    end
    
    %% Public Methods
    methods (Access = public)
        
        %---------------
        % get two cell arrays of all the raw data files and parameter files from the folder
        % for one configuration
        function self = getResultsFileList(self)
            prompt = sprintf('Choose the folder of all results from F0 = %d, Fs = %', self.F0, self.Fs);
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
            
          
            %prompt = sprintf('Choose the Class for selected F0 = %d, Fs = %d results',self.F0,self.Fs);
            %self.PmuClass = questdlg(prompt,'Choose PMU Class','M','P','M');
            
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
             dlg = dlgConfig(self, num2str(self.F0), num2str(self.Fs), self.PmuClass);
             waitfor(dlg)
            
            % TODO: Update the .ini file with the latest Raw Data Path
                        
        end
        
        % ----------------
        % creates the file name for the reports file once the
        % resultFileLists are populated
        function self = makeReportFileName(self)
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
        
        % --------------------------------
        function makeNewReportFile(self)
            
            self.hExcel = excelActiveX;      % consruct an excel object            
            self.hExcel.Connect;
            self.hExcel.Visible(1);
            self.hExcel.AddBook;            % Add a new workbook
            self.hExcel.SaveAs(self.ReportFile);
        end
            
         
        
        % ---------------------------------
        % loop through the data file list and write summary data to Excel
        function writeAnalysisToExcel(self)
            
            %Loop through all the data and param files, creating and adding
            %data to sheets
            for i = 1:numel(self.dataFiles)
                [sheetName,influenceFactor] = self.makeSheetName(self.paramFiles{1,i});
                [new] = self.activeSheet(sheetName);           % create a new sheet or set active
                [Hdr,Vals] = self.calcTableLine (self.dataFiles(i));
                Vals = [influenceFactor{1,1:numel(influenceFactor)}, Vals];
                
                % if the sheet is new, write the Header to the new sheet
                if new == true                    
                    Hdr = [influenceFactor.Properties.VariableNames{1,:},Hdr];   % concat
                    nLine = numel(Hdr); % length of the header line
                    nLine = self.hExcel.num2letters(nLine);
                    rng = strcat('A1:',nLine,'1');
                    self.hExcel.AddRange('Hdr','Cells',rng{1});
                    self.hExcel.WriteRange('Hdr',Hdr(1,:));
                    nextLine = 2;
                end
                
                % Write the values to the next line
                nLine = numel(Vals);
                nLine = self.hExcel.num2letters(nLine);
                rng = strcat('A',string(nextLine),':',nLine,string(nextLine));
                self.hExcel.AddRange('NextLine','Cells',rng{1});
                self.hExcel.WriteRange('NextLine',Vals);
                nextLine = nextLine+1;
            end
            
            self.hExcel.Save;
            
        end
        
        %----------------------------------------
        % loop through the sheets and create a set of plots for each sheet.
        function plotExcelAnalysis(self)
            
           nSheets = self.hExcel.hSheets.Count;
           for i = nSheets:-1:1
              self.hExcel.ActivateSheet(i);
              sheetName = self.hExcel.hSheet.Name;
              T = readtable(self.ReportFile,'Sheet',sheetName,'PreserveVariableNames',true);
              Lim = self.getLimitsFromSheetName(sheetName);
              
              % Create a new sheet 
              self.hExcel.NewSheet(strcat(sheetName,' plots'))              
              fig = self.plotVoltageTVE(T,Lim(1));
              self.hExcel.WriteData(fig,'B2')
              close(fig)
              fig = self.plotCurrentTVE(T,Lim(1));
              self.hExcel.WriteData(fig,'H2')
               close(fig)
              fig = self.plotFrequencyError(T,Lim(2));
              self.hExcel.WriteData(fig,'B16')
              close(fig)
              fig = self.plotRocofError(T,Lim(3));
              self.hExcel.WriteData(fig,'H16')
              close(fig)
            
              
           end
        end
       
        
        
    end
    
    
    %% Static Methods
    methods(Static)   
        
        
        function [Hdr,Vals] = calcTableLine(dataFile)
           C = readcell(cell2mat(dataFile)); 
           % Get the header row 
           hdr = string(C(1,:));
           idx = find(hdr=="RFE");       
           C = C(:,2:idx);      % only use from the first TVE to the RFE
           hdr = hdr(2:idx);
           % Make an array of header strings
           Hdr = strings(length(hdr),4);
           Vals = zeros(length(hdr),4);
           for i = 1:numel(hdr)
               Hdr(i,:) = [strcat("Mean",hdr(i)),strcat("St Dev",hdr(i)),strcat("Min",hdr(i)),strcat("Max",hdr(i))];
               Vals(i,1) = mean(cell2mat(C(2:end,i)));
               Vals(i,2) = std(cell2mat(C(2:end,i)));
               Vals(i,3) = max(cell2mat(C(2:end,i)));
               Vals(i,4) = min(cell2mat(C(2:end,i)));
           end
           Vals = reshape(Vals',[1,4*i]);
           Hdr = reshape(Hdr',[1,4*i]);
        end        
        
        
        % write a structure to a .ini file
        % Dirk Lohse (2021). struct2ini (
        % https://www.mathworks.com/matlabcentral/fileexchange/22079-struct2ini), 
        % MATLAB Central File Exchange. Retrieved January 15, 2021.
        function struct2Ini(file,structure)
            fid = fopen(file,'w');
            
            sects = fieldnames(structure);
            
            for i = 1:numel(sects)
                sect = char(sects(i));
                fprintf(fid,'\n[%s]\n',sect);
                mem = structure.(sect);
                if ~isempty(mem)
                    memNames = fieldnames(mem);
                    for j = 1:numel(memNames)
                        memName = char(memNames(j));
                        memVal = structure.(sect).(memName);
                        fprintf(fid,'%s=%s\n',memName,memVal);
                    end                    
                end
            end
            fclose(fid);
        end
        
        % read a structure from a .ini file
        %freeb (2021). ini2struct
        %(https://www.mathworks.com/matlabcentral/fileexchange/45725-ini2struct),
        %MATLAB Central File Exchange. Retrieved January 15, 2021.
        function Struct = ini2struct(FileName)
            % Parses .ini file
            % Returns a structure with section names and keys as fields.
            %
            % Based on init2struct.m by Andriy Nych
            % 2014/02/01
            f = fopen(FileName,'r');                    % open file
            while ~feof(f)                              % and read until it ends
                s = strtrim(fgetl(f));                  % remove leading/trailing spaces
                if isempty(s) || s(1)==';' || s(1)=='#' % skip empty & comments lines
                    continue
                end
                if s(1)=='['                            % section header
                    Section = genvarname(strtok(s(2:end), ']'));
                    Struct.(Section) = [];              % create field
                    continue
                end
                
                [Key,Val] = strtok(s, '=');             % Key = Value ; comment
                Val = strtrim(Val(2:end));              % remove spaces after =
                
                if isempty(Val) || Val(1)==';' || Val(1)=='#' % empty entry
                    Val = [];
                elseif Val(1)=='"'                      % double-quoted string
                    Val = strtok(Val, '"');
                elseif Val(1)==''''                     % single-quoted string
                    Val = strtok(Val, '''');
                else
                    Val = strtok(Val, ';');             % remove inline comment
                    Val = strtok(Val, '#');             % remove inline comment
                    Val = strtrim(Val);                 % remove spaces before comment
                    
                    [val, status] = str2num(Val);       
                    if status, Val = val; end           % convert string to number(s)
                end
                
                if ~exist('Section', 'var')             % No section found before
                    Struct.(genvarname(Key)) = Val;
                else                                    % Section found before, fill it
                    Struct.(Section).(genvarname(Key)) = Val;
                end
            end
            fclose(f);                        
        end
        
 
        

        
        % Get the sheet names and influcene factors from Fluke formatted
        % parameter files
        function [sheetName, influenceFactor] = FlukeParams(T)
            
            % anonymous functions
            % get the index of a Variable in a Table
            tabIdx = @(T,var) find(strcmp(T.Properties.VariableNames,var),1);
            % get the decimal value of a binary array
            bindec = @(b) uint16(bin2dec(sprintf('%d',b)));
            % make the sheet names
            
            switch T.eTestType
                case 0    % SteadyState
                    b = [T.Kv~=100, T.Ki2~= 100, T.Kh~=0, T.Ki1~=0];
                    %b = uint16(bin2dec(sprintf('%d',b)));
                    b = bindec(b);
                    switch b
                        case 0
                            sheetName = 'frequency range';
                            idx=tabIdx(T,'Fin');
                            influenceFactor = T(1,idx);
                        case {8, 6, 4}
                            sheetName = 'signal magnitude';
                            idx = tabIdx(T,'Kv');
                            influenceFactor = T(1,idx);
                            idx = tabIdx(T,'Ki2');
                            influenceFactor = [influenceFactor, T(1,idx)];
                        case 1
                            sheetName = 'out of band interference';
                            idx = tabIdx(T,'Ki1');
                            influenceFactor = T(1,idx);
                        case 2
                            sheetName = 'harmonic distortion';
                            idx = tabIdx(T,'Kh');
                            influenceFactor = T(1,idx);
                        otherwise
                            disp(T);
                            error('Illegal parameters')
                    end
                    
                case 1
                    sheetName = 'frequency ramp';
                    idx = tabIdx(T,'Fin');
                    influenceFactor = T(1,idx);
                    idx = tabIdx(T,'dF');
                    influenceFactor = [influenceFactor, T(1,idx)];
                    
                case 2
                    b = [T.Kx~=0, T.Ka~=0];
                    %b = uint16(bin2dec(sprintf('%d',b)));
                    b = bindec(b);
                    switch b
                        case 1
                            sheetName = 'phase modulation';
                            idx = tabIdx(T,'Ka');
                            influenceFactor = T(1,idx);
                        case 2
                            sheetName = 'amplitude modulation';
                            idx = tabIdx(T,'Kx');
                            influenceFactor = T(1,idx);
                        otherwise
                            sheetName = 'combined modulation';
                            idx = tabIdx(T,'Kx');
                            influenceFactor = T(1,idx);
                            idx = tabIdx(T,'Ka');
                            influenceFactor = [influenceFactor, T(1,idx)];
                    end
                    
                case 3
                    b = [T.Ka~=0, T.Kx~=0];
                    b = bindec(b);
                    switch b
                        case  1
                            sheetName ='amplitude step';
                            idx = tabIdx(T,'Kx');
                            influenceFactor = T(1,idx);
                        case 2
                            sheetName ='phase step';
                            idx = tabIdx(T,'Ka');
                            influenceFactor = T(1,idx);
                        otherwise
                            sheetName = 'combined step';p
                            idx = tabIdx(T,'Kx');
                            influenceFactor = T(1,idx);
                            idx = tabIdx(T,'Ka');
                            influenceFactor = [influenceFactor, T(1,idx)];
                    end
                otherwise
                    error('Unrecognized test type: %d',T.eTestType)
            end
            
        end
        
        
    end
    

    %% private methods    
    methods (Access = public)
        
        %-----------------------
        % from a parameter file, return the name of the excel sheet to write
        function [sheetName, influenceFactor] = makeSheetName(self,paramsFile)
            C = readcell(cell2mat(paramsFile));    % read params file to a cell array
            C = C';
            T = cell2table(C(2:end,:));     % cell array to a table
            
            % a problem with some of the paramfiles: duplicate parameter
            % names
            A = C(1,:);
            [~,~,X] = unique(A(:),'stable');
            Y = hist(X,unique(X));
            Y = (Y~=1);
            idx = find(Y==1);
            for i = 1:numel(idx)
                dupIdx = find(X==idx(i));
                for ii = 1:numel(dupIdx)
                    A{dupIdx(ii)} = sprintf('%s%d',A{dupIdx(ii)},ii);
                end
            end
            T.Properties.VariableNames = A;
            if size(T,1)>2
                [sheetName, influenceFactor] = self.NistParams(T);
            else
                [sheetName, influenceFactor] = self.FlukeParams(T);
            end
        end
                    
    % --------------------------
    % Either create a new sheet or set the sheet active
    function [new] = activeSheet(self,sheetName)
        % first, find out if a sheet exists
        new = false;
        [nbSheets,shList] = self.hExcel.ListSheets;
        
        % if there is only one sheet, and it's name is "Sheet1", rename
        % the sheet.
        if nbSheets == 1
            if shList{2} == "Sheet1"
                self.hExcel.GetSheet('Sheet1');
                self.hExcel.NameSheet(sheetName);
                new = true;
                return;
            end
        end
        
        % Check to see if the sheet already exist
        if numel(find(strcmp(shList(:,2),sheetName))) == 0
            self.hExcel.GetSheet(shList{end,2});
            self.hExcel.NewSheet(sheetName);    % make one if it does not exist
            new = true;
        end
        self.hExcel.GetSheet(sheetName)  % get the sheet
    end
    
    % Recursively collect all files from subfolders
    function filenames = getfn(self,folder,pattern)
        getfnrec(folder,pattern)
        
        idx = ~cellfun(@isempty, regexp(filenames,pattern));
        filenames =filenames(idx);
        
        % This nested function recursively goes through all subfolders
        % and collects all filenames within them
        function getfnrec(path,pattern)
            d = dir(path);
            filenames = {d(~[d.isdir]).name};
            filenames = strcat(path,filesep,filenames);
            
            dirnames = {d([d.isdir]).name};
            dirnames = setdiff(dirnames,{'.','..'});
            for i = 1:numel(dirnames)
                fulldirname = [path filesep dirnames{i}];
                filenames = [filenames, self.getfn(fulldirname,pattern)];
            end
        end
    end
    
    % Get the sheet names and influence factors from NIST formatted
    % parameter files
    function [sheetName, influenceFactor] = NistParams(self,T)
        
        % anonymous functions
        % get the index of a Variable in a Table
        tabIdx = @(T,var) find(strcmp(T.Properties.VariableNames,var),1);
        % get the decimal value of a binary array
        bindec = @(b) uint16(bin2dec(sprintf('%d',b)));
        % make the sheet names
        
        switch T.eTestType(1)
            case 0  %SteadyState
                b = [T.Kh(1)~=0, mean(T.Xm(1:3))~=self.vNom, mean(T.Xm(4:6))~=self.iNom];
                b = bindec(b);
                
                % there is a special case here.  If the current active sheet 
                % is "signal magnitude, and Kv and Ki are both nominal, then this is 
                % a signal magnitude test case.
                if b == 0;
                    sheetName = self.hExcel.hSheet.Name;
                    if sheetName == "signal magnitude"
                        b = 1;
                    end
                end
                
                switch b
                    case 0
                        sheetName = 'frequency range';
                        idx = tabIdx(T,'Fin');
                        influenceFactor = T(1,idx);
                    case {1,2,3}
                        sheetName = 'signal magnitude';
                        influenceFactor = table(mean((T.Xm(1:3))/self.vNom)*100,(mean(T.Xm(4:6))/self.iNom)*100);
                        influenceFactor.Properties.VariableNames = {'Kv','Ki'};
                    case 4
                        idx = tabIdx(T,'Fh');
                        influenceFactor = T(1,idx);
                        switch mod(T.Fh(1),T.Fin(1))
                            case 0
                                sheetName = 'harmonic distortion';
                                
                            otherwise
                                sheetName = 'out of band interference';
                        end
                end
                
            case 1 %Ramp
                sheetName = 'frequency ramp';
                idx = tabIdx(T,'Fin');
                influenceFactor = T(1,idx);
                idx = tabIdx(T,'dF');
                influenceFactor = [influenceFactor, T(1,idx)];
                
            case 2 % modulation
                b = [T.Kx(1)~=0, T.Ka(1)~=0];
                b = bindec(b);
                switch b
                    case 1
                        sheetName = 'phase modulation';
                        idx = tabIdx(T,'Ka');
                        influenceFactor = T(1,idx);
                    case 2
                        sheetName = 'amplitude modulation';
                        idx = tabIdx(T,'Ka');
                        influenceFactor = T(1,idx);
                    otherwise
                        sheetName = 'combined modulation';
                        idx = tabIdx(T,'Ka');
                        influenceFactor = T(1,idx);
                        idx = tabIdx(T,'Kx');
                        influenceFactor = [influenceFactor, T(1,idx)];
                end
                
            case 3 % step
                b = [T.Kas(1)~=0, T.Kxs (1)~=0, T.KfS(1)~=0, T.KrS(1)~=0];
                b = bindec(b);
                switch b
                    case 1
                        sheetName = 'ROCOF step';
                        idx = tabIdx(T,'KrS');
                        influenceFactor = T(1,idx);
                    case 2
                        sheetName = 'frequency step';
                        idx = tabIdx(T,'KfS');
                        influenceFactor = T(1,idx);
                    case 4
                        sheetName = 'amplitude step';
                        idx = tabIdx(T,'Kxs');
                        influenceFactor = T(1,idx);
                    case 8
                        sheetName = 'phase step';
                        idx = tabIdx(T,'Kas');
                        influenceFactor = T(1,idx);
                    otherwise
                        sheetName = 'combined step';
                        idx = tabIdx(T,'Kas');
                        influenceFactor = T(1,idx);
                        idx = tabIdx(T,'Kxs');
                        influenceFactor = [influenceFactor, T(1,idx)];
                        idx = tabIdx(T,'KfS');
                        influenceFactor = [influenceFactor, T(1,idx)];
                        idx = tabIdx(T,'KrS');
                        influenceFactor = [influenceFactor, T(1,idx)];
                end
                
            otherwise
                error('Unrecognized test type: %d',T.eTestType)
        end
    end
    
    % Create a formatted plot of the Voltage TVE
    function fig = plotVoltageTVE(self,T,limit)
        
        lstVoltage = ["MaxTVE_VA" "MaxTVE_VB" "MaxTVE_VC" "MaxTVE_Vp"];

        % Create figure
        fig = figure;
        set(fig,'Position',self.figPos);
        
        % X and Y data
        X1 = T{:,1};
        YMatrix1 = zeros(4,size(T,1));
        for ii = 1:numel(lstVoltage)
            YMatrix1(ii,:) = T{:,lstVoltage(ii)};
        end

        
        % Create axes
        axes1 = axes('Parent',fig,...
            'Position',self.axPos, 'YGrid', 'on');        
        hold(axes1,'on');
        
        % Create multiple lines using matrix input to plot
        plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
        set(plot1(1),'DisplayName','MaxTVE\_VA');
        set(plot1(2),'DisplayName','MaxTVE\_VB');
        set(plot1(3),'DisplayName','MaxTVE\_VC');
        set(plot1(4),'DisplayName','MaxTVE\_V+');
        
        % Draw the limit line
        yline(limit(1),'-r','linewidth',2,'DisplayName','TVE Limit')
        
        title('Voltage TVE')
        % Create ylabel
        ylabel('TVE (%)');        
        % Create xlabel
        xlabel('Input Frequency (Hz)');
        
        ylim(axes1,[0 1.2]);
        hold(axes1,'off');
        % Set the remaining axes properties
        set(axes1,'FontSize',6, 'YGrid', 'on');
        % Create legend
        lgd = legend(axes1,'show');
        set(lgd,'Location','eastoutside');
        set (lgd,'Position',self.lgdPos);
    end
    
    % Create a formatted plot of the Current TVE
    function fig = plotCurrentTVE(self,T,limit)
        
        lstCurrent = ["MaxTVE_IA" "MaxTVE_IB" "MaxTVE_IC" "MaxTVE_Ip"];  

        % Create figure
        fig = figure;            
        set(fig,'Position',self.figPos);

        % X and Y date
        X1 = T{:,1};
        YMatrix1 = zeros(4,size(T,1));
        for ii = 1:numel(lstCurrent)
            YMatrix1(ii,:) = T{:,lstCurrent(ii)};
        end

        
        % Create axes
        axes1 = axes('Parent',fig,...
            'Position',self.axPos, 'YGrid', 'on');
        hold(axes1,'on');
       
        % Create multiple lines using matrix input to plot
        plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
        set(plot1(1),'DisplayName','MaxTVE\_IA');
        set(plot1(2),'DisplayName','MaxTVE\_IB');
        set(plot1(3),'DisplayName','MaxTVE\_IC');
        set(plot1(4),'DisplayName','MaxTVE\_I+');

        % Draw the limit line
        yline(limit(1),'-r','linewidth',2,'DisplayName','TVE Limit')
        
        title('Current TVE')
        
        % Create ylabel
        ylabel('TVE (%)');
        % Create xlabel
        xlabel('Input Frequency (Hz)');

        ylim(axes1,[0 1.2]);
        hold(axes1,'off');
        % Set the remaining axes properties
        set(axes1,'FontSize',6,'YGrid','on');
        % Create legend
        lgd = legend(axes1,'show');
        set(lgd,'Location','eastoutside');
        set (lgd,'Position',self.lgdPos);
    end
    
    % Create a formatted plot of the Frequency Error
    function fig = plotFrequencyError(self,T,limit)
        
        lstFE = ["MinFE", "MaxFE"];
        
        % Create figure
        fig = figure;
        set(fig,'Position',self.figPos);
        
        % X and Y data
        X1 = T{:,1};
        YMatrix1 = zeros(4,size(T,1));
        for ii = 1:numel(lstFE)
            YMatrix1(ii,:) = T{:,lstFE(ii)};
        end
        
        % Create axes
        axes1 = axes('Parent',fig,...
            'Position',self.axPos, 'YGrid', 'on');        
        hold(axes1,'on');
        
        % Create multiple lines using matrix input to plot
        plot1 = plot(X1,YMatrix1,'-b','LineWidth',2,'Parent',axes1);
        set(plot1(1),'DisplayName','Max\_FE');
        set(plot1(2),'DisplayName','Min\_FE');
       
        % Draw the limit line
        if isnumeric(limit)&& ~isinf(limit)
            limLine = yline(limit(1),'-r','linewidth',2,'DisplayName','FE Limit');
            yline(-limit(1),'-r','linewidth',2)
            lgd = legend([plot1(1),plot1(2), limLine]);
        else
            lgd = legend([plot1(1),plot1(2)]);
        end

        
        title('Frequency Error')
        % Create ylabel
        ylabel('FE (Hz)');        
        % Create xlabel
        xlabel('Input Frequency (Hz)');
        
        ylim(axes1,[-.006 .006]);
        hold(axes1,'off');
        % Set the remaining axes properties
        set(axes1,'FontSize',6, 'YGrid', 'on');
        % Create legend
        set(lgd,'Location','eastoutside');
        set (lgd,'Position',self.lgdPos);
    end    
    
        % Create a formatted plot of the ROCOF Error
    function fig = plotRocofError(self,T,limit)

        lstRFE = ["MinRFE", "MaxRFE"];
        
        % Create figure
        fig = figure;
        set(fig,'Position',self.figPos);
        
        % X and Y data
        X1 = T{:,1};
       YMatrix1 = zeros(4,size(T,1));
        for ii = 1:numel(lstRFE)
            YMatrix1(ii,:) = T{:,lstRFE(ii)};
        end
         
        % Create axes
        axes1 = axes('Parent',fig,...
            'Position',self.axPos, 'YGrid', 'on');        
        hold(axes1,'on');
        
        % Create multiple lines using matrix input to plot
        plot1 = plot(X1,YMatrix1,'-b','LineWidth',2,'Parent',axes1);
        set(plot1(1),'DisplayName','Max\_RFE');
        set(plot1(2),'DisplayName','Min\_RFE');
       
        % Draw the limit line
        if isnumeric(limit)&& ~isinf(limit)
            limLine = yline(limit(1),'-r','linewidth',2,'DisplayName','RFE Limit');
            yline(-limit(1),'-r','linewidth',2)
            lgd = legend([plot1(1),plot1(2), limLine]);
        else
            lgd = legend([plot1(1),plot1(2)]);
        end
       
        title('ROCOF Error')
        % Create ylabel
        ylabel('RFE (Hz/S)');        
        % Create xlabel
        xlabel('Input Frequency (Hz)');
        
        ylim(axes1,[-1.5 1.5]);
        hold(axes1,'off');
        % Set the remaining axes properties
        set(axes1,'FontSize',6, 'YGrid', 'on');
        % Create legend
        set(lgd,'Location','eastoutside');
        set (lgd,'Position',self.lgdPos);
    end    
    
    % from the sheet name, get the TVE, FE and RFE limits from the 
    % specifications property
    function Lim = getLimitsFromSheetName(self,sheetname)
        switch sheetname
            case 'frequency range'
                Lim = self.specifications.FreqRng;
                return
                
            case 'signal magnitude'
                Lim = self.specifications.MagRng;
                return
                
            case 'harmonic distortion'
                Lim = self.specifications.Harm;
                return
                
            case 'out of band interference'
                Lim = self.specifications.OOB;
                return
                
            case 'frequency ramp'
                Lim = self.specifications.RampPos;
                return
                
            case {'phase modulation', 'amplitude modulation', 'combined modulation'}
                Lim = self.specifications.AmplMod;
                return
                
            case {' amplitude step' 'phase step', 'combined step'}
                Lim = [self.specifications.PhaseStepRespTime; self.specifications.PhaseStepDelayTime; self.specifications.PhaseStepOverShoot];
                return
                
            otherwise
                warning ('There are no known limits for %s', sheetname)
        end
        
    end
    
    end 
end