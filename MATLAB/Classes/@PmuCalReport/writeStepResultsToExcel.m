function [i] = writeStepResultsToExcel(self,i,influenceFactor)
% For step tests, Create a table of ETS data.

% First, find out how many files have the same influence factor
k = i;
iFactor = influenceFactor.Properties.VariableNames{1,1};
iF = iFactor;
while (iF == iFactor)
    k = k + 1;
    if k <= numel(self.paramFiles)
        [~, influenceFactor] = self.makeSheetName(self.paramFiles{1,k});
        iF = influenceFactor.Properties.VariableNames{1,1};
    else
        k = k-1;
        break
    end
end
n = k - i;      % this is how many files have this influence factor

% Fill a matrix with all the data
j = i;          % Start from the first file with this iFactor
C = readcell(cell2mat(self.dataFiles(j)));
M = zeros(length(C)-1,size(C,2),n);
h = 1;
while (h <= n)
    M(:,:,h) = cell2mat(C(2:end,:));
    j = j + 1;
    h = h + 1;
    C = readcell(cell2mat(self.dataFiles(j)));
end

% Interleave the matrix into a single array
A = permute(M,[2,3,1]);
A = reshape(A,size(A,1),[])';

% correct the time stamps, beginning with the first
t = A(1,1);     %First time of the test
dt = mean(diff(M(:,1,1)))/10;
for j = 0:length(A)-1
    A(j+1,1) = t+(j*dt);
end

% Write the cell array to the active sheet
hdr = string(C(1,:));
self.writeExcelHeader(hdr);
nCol = self.hExcel.num2letters(size(A,2));
rng = strcat('A2:',nCol{1},string(size(A,1)));
self.hExcel.AddRange('Rng','Cells',rng{1});
self.hExcel.WriteRange('Rng',A);
i = i + k;      % Update i for the next set of files

end