
function [sheetName, influenceFactor] = FlukeParams(self,T)
% Get the sheet names and influcene factors from Fluke formatted
% parameter files

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
                sheetName ='amplitude step pos';
                if T.Kx < 0; sheetName ='amplitude step neg'; end;
                idx = tabIdx(T,'Kx');
                influenceFactor = T(1,idx);
            case 2
                sheetName ='phase step pos';
                if T.Kx < 0; sheetName ='phase step neg'; end;
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