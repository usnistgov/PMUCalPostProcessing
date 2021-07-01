function [nextLine] = writeResultsToExcel(self,nextLine,i,new,influenceFactor)
% for non-ste tests, calculate and write the single line of results from the data
[Hdr,Vals] = calcTableLine (self.dataFiles(i));
Vals = [influenceFactor{1,1:numel(influenceFactor)}, Vals];

% if the sheet is new, write the Header to the new sheet
if new == true
    Hdr = [influenceFactor.Properties.VariableNames{1,:},Hdr];   % concat
    self.writeExcelHeader(Hdr);
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
    Vals(i,3) = min(cell2mat(C(2:end,i)));
    Vals(i,4) = max(cell2mat(C(2:end,i)));
end
Vals = reshape(Vals',[1,4*i]);
Hdr = reshape(Hdr',[1,4*i]);
end
