classdef testExcelActiveX < matlab.unittest.TestCase
    % to run the tests:
    %    testCase = testExcelActiveX;
    %    res = run(testCase);
    
    %   You only need to run the testCase = testExcelActiveX once until you clear
    %   all, after editing and saving any file, you only need to run
    %   res=run(testCase) to run the unit tests.
    
    %% Properties
    properties
        Excel       % Handle for the excelActiveX object
    end
    
    %% Test Methods
    % Any uncommented test functions will be run in series. These functions are
    % found below in the Public Methods in the order they appear in the list.
    methods(Test)
        function regressionTests (testCase);
            %testOpenCloseApplication(testCase);
            %testNewSheet(testCase)
            testWriteFigure(testCase);
            
        end
        
    end
    
    %% Public Methods
    methods(Access = public)
        
        %%Test simple connect, make visible and disconnect
        function testOpenCloseApplication(testCase)
            testCase.Excel = excelActiveX;
            testCase.Excel.Connect;  % connect to the Excel Application Server
            testCase.Excel.Visible(1);  % make it visible (it is not by dfault)
            testCase.Excel.NewBook;     % Add a workbook
            uiwait(msgbox('You should see Excel. Next it will close'));
            testCase.Excel.Disconnect;
        end
        
        %%Test the addition of sheets, sheets with names and sheets that
        %follow other sheets with and without names.
        function testNewSheet(testCase)
            % open a new book
            testCase.Excel = excelActiveX;
            testCase.Excel.Connect;  % connect to the Excel Application Server
            testCase.Excel.Visible(1);  % make it visible (it is not by dfault)
            testCase.Excel.NewBook;     % Add a workbook
            
            % add a sheet with no name
            [nbSheets, shList] = testCase.Excel.ListSheets; % get the list of sheets
            testCase.Excel.GetSheet(shList{nbSheets,2});
            testCase.Excel.NewSheet;
            testCase.Excel.NewSheet("LastSheet");
                      
            
            % Now wait for the user and close out
            uiwait(msgbox('You should see Sheet 2 after Sheet1 and LastSheet at the end'));
            testCase.Excel.Disconnect;            
        end
        
        function testWriteFigure(testCase)
            testCase.Excel = excelActiveX;
           testCase.Excel = excelActiveX;
            testCase.Excel.Connect;  % connect to the Excel Application Server
            testCase.Excel.Visible(1);  % make it visible (it is not by dfault)
            testCase.Excel.NewBook;     % Add a workbook
            
            % add a sheet with no name
            [nbSheets, shList] = testCase.Excel.ListSheets; % get the list of sheets
            testCase.Excel.GetSheet(shList{nbSheets,2});            

            %create a figure to be written
            fig = figure(1);
            hFigure = gcf;
            addr = 'B3';
            testCase.Excel.WriteData(hFigure,addr);
            
            
            
        end
        
    end
end