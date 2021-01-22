function [configOut] = uiConfig(configIn)
% A UI to verify the configuration of the results file to be analysed

configOut = configIn;

% build the ui

% uiFig = uifigure('units','pixels',...
%                   'position',[800,500,200,50],...
%                   'name','Verify the PMU configuration of the files to be analysed',...
%                   'WindowStyle','normal'...
%                   );
S.fh = uifigure;
S.fh.Name = 'Verify the PMU configuration of the files to be analysed';
S.fh.Units = 'pixels';
S.fh.Position = [1700, 1700, 520, 200];



S.pb(1) = uibutton(S.fh);
S.pb(1).Position = [220 20 100 40];
S.pb(1).Text = 'Continue';

S.TF = false


set(S.pb(:),'ButtonPushedFcn',{@pb_call,S})    % set calbacks


if S.TF
    close(S.Fh);
end
    
    function[] = pb_call(varargin)
    %callback for the buttons.
    if varargin{1}==S.pb(1)
        S.TF = true;
    end
    end
    

end

