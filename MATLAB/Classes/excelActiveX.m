% Class that encapsulates Excel activeX handlers for Matlab

classdef excelActiveX < handle
    
    properties
        reportPath  %path to location where reports will be written
        hExcel   % handle to an activeX Excel object
        hBook    % handle to an excel Book object
        hSheet   % handle to a sheet
    end
    
    
    methods
        
        % ask the user where to put the reports
        function getReportPath(self)
            self.reportPath = uigetdir('.','Path to report files');
        end
        
        %% Application handeling        
        % initialize the handle (hExcel) to the Excel object
        function Connect(self)
            
            % connect to an Excel COM server
            try
                self.hExcel = actxserver('Excel.Application');
            catch ME
                msg = [ME.message,'\nCannot start Excel server.'];
                ME  = MException(ME.identifier,msg);
                throw(ME);
            end
            % set default path
            self.hExcel.DefaultFilePath = pwd;
            % # sheets in new workbook
            self.hExcel.SheetsInNewWorkbook = 1;
        end
        
        % disconnect from the Excel server
        % (closes all open workbooks and does not save)
        function Disconnect(self)
            self.CloseAllBooks;         % release all open books
            self.hExcel.Quit;                    % delete Excel object
            self.hExcel.release;                 % stop process
        end
        
        function NewBook(self)
            try
                self.hBook = self.hExcel.Workbooks.Add;
            catch ME
                msg = '\nUnable to create a new Workbook. Verify that hExcel is valid';
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
        end
        
        %% Book Handling
        % closes all open books but does not save or prompt for save
        function CloseAllBooks(self)
            try
                nbBooks = self.hExcel.Workbooks.Count;
            catch ME
                msg = 'hExcel is not a valid handle to an Excel server.';
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
            for b = nbBooks:-1:1
                self.hBook = self.hExcel.Workbooks.Item(b);
                self.hBook.Close(false);
                self.hBook.release;
            end
        end
        
        function Visible(self,val)
            self.hExcel.Visible = val;           
        end
        
        %% Sheet handling
        
        % ListSheets
        % returns the number of sheets and a list of the sheet handles
        function [nbSheets,shList] = ListSheets(self)
            try
                nbSheets = self.hBook.Sheets.Count;
            catch ME
                msg('\nCannot get the sheet list, verify that hBook is valid');
                ME = MException(ME.identifier,[ME.message,msg]);
                throw(ME)
            end
            shList = cell(nbSheets,3);
            for s = 1:nbSheets
                shList{s,1} = self.hBook.Sheets.Item(s).Index;
                shList{s,2} = self.hBook.Sheets.Item(s).Name;
                shList{s,3} = self.hBook.Sheets.Item(s).CodeN;                
            end
        end
        
        % GetSheet
        % get a sheet handle
        function GetSheet(self,sheet)
            try
                self.hSheet = get(self.hBook.sheets,'item',sheet);
            catch ME
                msg = ['Worksheet does not exist, or hBook ', ...
                    'is not a valid handle to a workbook.'];
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end         
        end
        
        % NewSheet
        % creates a new sheet after the sheet in self.hSheet
        % optionally, a string argument may be included to name the sheet
        function NewSheet(self,varargin)                 
            try
                self.hSheet = self.hBook.Sheets.Add([],self.hSheet);
            catch ME
                msg = ['Unable to create a new sheet. ', ...
                    'either hBook or hSheet is invalid'];
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
            
            if ~isempty(varargin) 
                if isstring(varargin{1})
                    self.NameSheet(varargin{1}); 
                end
            end
       end
            
        % NameSheet
        % renames the sheet in hSheet
        function NameSheet(self,strName)
            try
                self.hSheet.Name = strName;
            catch ME
                msg = ['Unable to name a sheet. ', ...
                    'either hSheet or the name string is invalid'];
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
        end
                
            
        %% Writing data or plots
        
        % data is a matrix, array, table, or  timeseries object. It
        % contains the data that will be written to the sheet. Can also pass a
        % handle to a figure:the figure is copied to the worksheet.
        %
        % addr is a vector of two numbers, indicating the top-left corner where the
        % data are to be written. addr can be given in 'A1'-format (i.e. 'B3'
        % instead of [3,2]). Default is [1 1] or 'A1'.
        %
        % The function returns the address where data was written to in
        % 'A1'-format.
        
        function addr = WriteData(self,data,varargin)
            % is the data a handle to a figure?
            try
                isFigure = (ishandle(data) & strcmp(get(data,'type'),'figure'));
            catch
                isFigure = false;
            end
            disp(isFigure);
            
            % It's a figure.
            if isFigure
                % addr argument
                if isempty(varargin)
                    addr = 'A1';
                elseif isnumeric(varargin{1})
                    addr = XRangeAddress([1,1],varargin{1});
                end
                % copy content of figure to clipboard using hgexport
                hgexport(data,'-clipboard');
                % paste clipboard into hSheet
                try
                    self.hSheet.Paste(self.hSheet.Range(addr));
                catch ME
                    msg = 'hSheet is not a valid handle to an Excel worksheet.';
                    ME  = MException(ME.identifier,[ME.message,msg]);
                    throw(ME);
                end
                % It's a table.
            elseif isa(data,'table')
            end
            
            
        end

    end
    
end