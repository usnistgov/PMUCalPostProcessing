function self = makeNewReportFile(self)
% Creates a new Excel file 

self.hExcel = excelActiveX;      % consruct an excel object
self.hExcel.Connect;
self.hExcel.Visible(1);
self.hExcel.AddBook;            % Add a new workbook
self.hExcel.SaveAs(self.ReportFile);
end