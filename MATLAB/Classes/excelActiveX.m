% Class that encapsulates Excel activeX handlers for Matlab

% terminology taken from the Microsoft Office VBA Reference Object Model

classdef excelActiveX < handle
    
    properties
        hExcel   % activeX Excel object handle
        hBooks   % Workbooks collection object handle
        hBook    % Workbook (single) object handle
        hSheets  % Sheets colection object handle
        hSheet   % Active sheet object handle
        rangeTable = table();   % A table that holds multiple Range object handles
        filePath  %path to location where files will be written
        
    end
    
    
    methods

        %% Excel Com Object       
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
            % Get the Workbooks Collection object
            try
                self.hBooks = self.hExcel.WorkBooks;
            catch ME
                msg = [ME.message,'\nUnable to get Worrkbooks Collection Object'];
                ME  = MException(ME.identifier,msg);
                throw(ME);
            end                
            % set default path
            self.hExcel.DefaultFilePath = pwd;
        end
        
        function Visible(self,val)
            self.hExcel.Visible = val;           
        end        
        
        % disconnect from the Excel server
        % (closes all open workbooks and does not save)
        function Disconnect(self)
            self.CloseAllBooks;         % release all open books
            self.hExcel.Quit;                    % delete Excel object
            self.hExcel.release;                 % stop process
        end
        
 
        %% Workbooks Object
        
        % create a new book
        function AddBook(self)
            try
                self.hBook = self.hExcel.Workbooks.Add;
                self.hSheets = self.hBook.Sheets;
                self.hSheet = self.hSheets.Item(1);
            catch ME
                msg = '\nUnable to create a new Workbook. Verify that hExcel is valid';
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end   
        end
        
        % Activate a book
        function ActivateBook(self,book)
            try
                self.hBooks.Item(book).Activate
                self.hBook = self.hExcel.ActiveWorkbook;
                self.hSheets = self.hBook.Sheets;
            catch ME
                if isnumeric(book)
                    msg = sprintf('\nUnable to activate book %d\n', book);
                else
                    msg = sprintf('\nUnable to activate book %s\n', book);
                end
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
                
            end
        end
        
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
                self.hBooks = self.hExcel.Workbooks.Item(b);
                self.hBooks.Close(false);
                self.hBooks.release;
            end
        end
        
        function SaveAs(self,FileName)
            % first, check that the folder exists and create it if not
            [path,~,~] = fileparts(FileName);
            if ~exist(path,'dir')
                try
                    mkdir(path);
                catch ME
                    msg = '\nUnable to create save folder\n';
                    ME  = MException(ME.identifier,[ME.message,msg]);
                    throw(ME);
                end                   
            end
            
            % now try to save the file
            try
                self.hBook = self.hExcel.ActiveWorkBook;
                self.hBook.SaveAs(FileName)
            catch ME
                msg = '\nUnable to save workbook\n';
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
        end  
        
        function Save(self)
            try
                self.hBook.Save;
            catch ME
                msg = '\nUnable to save workbook\n';
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end                
        end
        
        %% Sheets Object
        
        % ListSheets
        % returns the number of sheets and a list of the sheet handles
        function [nbSheets,shList] = ListSheets(self)
            try
                nbSheets = self.hSheets.Count;
            catch ME
                msg('\nCannot get the sheet list, verify that hBook is valid');
                ME = MException(ME.identifier,[ME.message,msg]);
                throw(ME)
            end
            shList = cell(nbSheets,3);
            for s = 1:nbSheets
                shList{s,1} = self.hSheets.Item(s).Index;
                shList{s,2} = self.hSheets.Item(s).Name;
                shList{s,3} = self.hSheets.Item(s).CodeN;                
            end
        end
        
        % ActivateSheet
        function ActivateSheet(self,sheet)
            self.hSheets.Item(sheet).Activate;
            self.hSheet = self.hExcel.ActiveSheet;
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
        % creates a new sheet after the active sheet
        % optionally, a string argument may be included to name the sheet
        function NewSheet(self,varargin)                 
            try
                self.hSheet = self.hSheets.Add([],self.hSheet);
            catch ME
                msg = ['Unable to create a new sheet. ', ...
                    'either hBook or hSheets is invalid'];
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
            
            % Name the sheet
            if ~isempty(varargin)
                if ischar(varargin{1}) || isstring(varargin{1})
                    self.NameSheet(varargin{1}); 
                else
                    warning('NewSheet argument is not a string or charactor array. New sheet will use default name')
                end
            end
       end
            
        % NameSheet
        % renames the sheet in hSheets
        function NameSheet(self,strName)
            try
                self.hSheet.Name = strName;
            catch ME
                msg = ['Unable to name a sheet. ', ...
                    'either hSheets or the name string is invalid'];
                ME  = MException(ME.identifier,[ME.message,msg]);
                throw(ME);
            end
        end
        
        %% Range Objects
        
        % self.hRangeTableis a table of range object handles.  They can be
        % accessed by their range name. 
        
        % create a new range on the active sheet.
        %   AddRange('name', [arg, val])
        %   Range Name is required, the following are optional arguments:
        %   ['Cell', 'A1]' a single cell range
        %   ['Cells', [1,1] ] a single cell range in "cells" format
        %   ['Row', 1] a single row on the worksheet
        %   ['Rows', []] all rows on the worksheet (the second argument
        %                will be ignored)
        %   ['Column', 'A'] a single column on the worksheet (Can replace 'A' with 1)
        %   ['Columns' []) all columns on the worksheet (second argument will be ignored)
        %
        %   Examples:
        %       AddRange ('name') adds a empty range object to the rangeTable called 'name
        %       AddRange ('name' 'Cell', 'B5') adds a single cell range to the rangeTable (can also use [5,2])
        %       AddRange ('name' 'Start' 'B5', End' 'D9') adds the range from B5 to D9 to rangeTable       
        %  
        function AddRange(self, name, varargin)
            
            % check the input formatting
            if ~(ischar(name) || isstring(name))
                error('AddRange name must be a string or a charactor array')
            end
            nArgs = numel(varargin);
            if mod(nArgs,2)~=0
                error('AddRange improperly formatted arguments.  optional arguments must be in pairs [arg, value]')
            end
            
            switch varargin{1}
                case {'Cell' 'cell' 'Cells' 'cells'}
                    self.rangeTable.(name) = get(self.hSheet,'Range',varargin{2});
                    return
                    
                case {'Row' 'row' 'Rows' 'rows'}
                    self.rangeTable.(name) = get(self.hSheet,'Rows',varargin{2});
                    return
                    
                case {'Column' 'column' 'Columns' 'columns'}
                    self.rangeTable.(name) = get(self.hSheet,'Columns',varargin{2});
                    return                                    
                    
                otherwise
                    error('AddRange unrecognized argument ''%s''',varargin{1})
            end
        end
        
    
        % remove a range from the range table
        function RemoveRange (self, name)
            self.rangeTable.(name)=[];
            %TODO I could not figure out how to destroy the range object in hSheet.  
            % This is a memory leak.            
        end
        
        % write values to a named range
        function WriteRange (self, name, value) 
                % if the values are strings, they must be in a cell array of strings
                if isstring(value); value=cellstr(value); end

                self.rangeTable.(name).Value =  value;                
        end
        
        % deletes the cells in a named range
        function DeleteRange (self, name)
            get(self.rangeTable.(name),'Delete');  
            % note that once cells are deleted, the range object no longer points to anything
            self.RemoveRange(name);
        end
        
        % clears the cells in a named range
        function ClearRange (self, name)
            get(self.rangeTable.(name),'Clear');
        end
        
        % inserts cells at a named range
        function InsertRange (self, name)
            get(self.rangeTable.(name),'Insert')
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
        
        function WriteData(self,data,varargin)
            % is the data a handle to a figure?
            try
                isFigure = (ishandle(data) & strcmp(get(data,'type'),'figure'));
            catch
                isFigure = false;
            end
            
            % It's a figure.
            if isFigure
                % addr argument
                if isempty(varargin)
                    addr = 'A1';
                elseif ischar(varargin{1})
                    %addr = XRangeAddress([1,1],varargin{1});
                    addr = varargin{1};
                end
                % copy content of figure to clipboard using hgexport
                % This will export the figure using the size shown on the screen
                style = hgexport('factorystyle'); %g et the style
                style.Resolution = 0;   % Resolution = 0 meanst use the screen resolution
                hgexport(data,'-clipboard',style);
                % paste clipboard into hSheets
                try
                    self.hSheet.Paste(self.hSheet.Range(addr));
                catch ME
                    msg = 'hSheet is not a valid handle to an Excel worksheet.';
                    ME  = MException(ME.identifier,[ME.message,msg]);
                    throw(ME);
                end
                % It's a table.
            elseif isa(data,'table')
                self.hSheet.Value(self.hSheet.Range(addr)) = data;
            end
            
            
        end
        
        % Convert a number into its albhabetic base-26 representaton
        function lets = num2letters(self,nums)
            lets = arrayfun(@(n)num2char(n),nums,'UniformOutput',0);
            function s = num2char(d)
                b = 26;
                n = max(1,round(log2(d+1)/log2(b)));
                while (b^n <= d)
                    n = n + 1;
                end
                s(n) = rem(d,b);
                while n > 1
                    n = n - 1;
                    d = floor(d/b);
                    s(n) = rem(d,b);
                end
                n = length(s);
                while (n > 1)
                    if (s(n) <= 0)
                        s(n) = b + s(n);
                        s(n-1) = s(n-1) - 1;
                    end
                    n = n - 1;
                end
                s(s<=0) = [];
                symbols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
                s = reshape(symbols(s),size(s));
            end
        end
        
        function nums = letters2nums(self,lets)
            nums = cellfun(@(x) (sum((double(x)-64).*26.^(length(x)-1:-1:0))),lets);            
        end
        
        
        % ask the user where to put the reports
        function getReportPath(self)
            self.filePath = uigetdir('.','Path to report files');
        end
                

    end
    
end