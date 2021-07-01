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
        if b == 0
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
        
    case 1 % modulation
        b = [T.Kx(1)~=0, T.Ka(1)~=0];
        b = bindec(b);
        switch b
            case 1
                sheetName = 'phase modulation';
                idx = tabIdx(T,'Fa');
                influenceFactor = T(1,idx);
            case 2
                sheetName = 'amplitude modulation';
                idx = tabIdx(T,'Fx');
                influenceFactor = T(1,idx);
            otherwise
                sheetName = 'combined modulation';
                idx = tabIdx(T,'Fa');
                influenceFactor = T(1,idx);
                idx = tabIdx(T,'Fx');
                influenceFactor = [influenceFactor, T(1,idx)];
        end
        
    case 2 %Ramp
        sheetName = 'frequency ramp';
        idx = tabIdx(T,'Fin');
        influenceFactor = T(1,idx);
        idx = tabIdx(T,'dF');
        influenceFactor = [influenceFactor, T(1,idx)];
        
    case 3 % step
        b = [T.KaS(1)~=0, T.KxS(1)~=0, T.KfS(1)~=0, T.KrS(1)~=0];
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
                sheetName = 'amplitude step pos';
                if T.KxS < 0; sheetName = 'amplitude step neg'; end
                idx = tabIdx(T,'KxS');
                influenceFactor = T(1,idx);
            case 8
                sheetName = 'phase step pos';
                if T.Kas < 0; sheetName = 'phase step neg'; end
                idx = tabIdx(T,'KaS');
                influenceFactor = T(1,idx);
            otherwise
                sheetName = 'combined step';
                idx = tabIdx(T,'KaS');
                influenceFactor = T(1,idx);
                idx = tabIdx(T,'KxS');
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