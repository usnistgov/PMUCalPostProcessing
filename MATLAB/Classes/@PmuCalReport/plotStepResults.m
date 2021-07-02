function plotStepResults(self,T,Lim,xLabel)
% Voltage Response time plots and values
[respTime, fig] = plotStepVoltageResponseTime(self,T,Lim(1),xLabel(1));
self.hExcel.WriteData(fig,'B4')   % write the plot
self.hExcel.WriteData(respTime,'B1')  % write the values
close(fig)

% Current response time plots and values
[respTime, fig] = plotStepCurrentResponseTime(self,T,Lim(1),xLabel(1));
self.hExcel.WriteData(fig,'J4')   % write the plot
self.hExcel.WriteData(respTime,'J1')  % write the value
close(fig)

% Frequency Response Time plot and values
[respTime, fig] = plotStepFreqResponseTime(self,T,Lim(2),xLabel(1));
self.hExcel.WriteData(fig,'Q4')   % write the plot
self.hExcel.WriteData(respTime,'Q1')  % write the value
close(fig)

% ROCOF Response Time plot and values
[respTime, fig] = plotStepROCOFResponseTime(self,T,Lim(2),xLabel(1));
self.hExcel.WriteData(fig,'X4')   % write the plot
self.hExcel.WriteData(respTime,'X1')  % write the value
close(fig)

% Voltage Over and Undershoot
[Results, fig] = plotStepVoltageOvershoot(self,T,Lim(4),xLabel(1));
Results{:,'phase'} = {'VA';'VB';'VC';'V+'};
self.hExcel.WriteData(fig,'B24')   % write the plot
%self.hExcel.WriteData(Results,'B12')  % write the value ARG: cannot write the strings
self.hExcel.WriteData(Results(:,2:end),'B18')  % write the value
close(fig)

end


%% -------------------------------------------------------------------------
function [respTime,fig] = plotStepVoltageResponseTime(self,T,limit,xLabel)
% Plot the Step Test Response Time

lstVoltage = ["TVE_VA" "TVE_VB" "TVE_VC" "TVE_Vp"];

% Create Figure
fig = figure;
set(fig,'Position',self.figPos)

%X and ////y Data
dT = mean(diff(T{:,1}));
X1 = (0:size(T,1)-1)*dT;

YMatrix1 = zeros(4,size(T,1));
for ii = 1:numel(lstVoltage)
    YMatrix1(ii,:) = T{:,lstVoltage(ii)};
end

% for each phase, find the point where the TVE first crosses 1%
nPhases = (size(YMatrix1,1));
iResp = zeros(nPhases,2);    % will hold indesex of 1% TVE crossing points
respTime = zeros(1,nPhases);
for i = 1:nPhases
    idx = find(YMatrix1(i,:)> 1);
    iResp(i,:) = [idx(1),idx(end)];
    respTime(1,i) = X1(iResp(i,end)) - X1(iResp(i,1));
end

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','TVE\_VA');
set(plot1(2),'DisplayName','TVE\_VB');
set(plot1(3),'DisplayName','TVE\_VC');
set(plot1(4),'DisplayName','TVE\_V+');

% Draw the limit line
yline(1,'-g','linewidth',2,'DisplayName','TVE Limit')

% Draw the response time lines on the longest resp time
[~,idx] = max(respTime);    % index of the maximum repsones time
respTime = array2table(respTime,'VariableNames',{'RT_VA','RT_VB','RT_VC','RT_V+'});
xline(X1(iResp(idx,1)),'b','linewidth',2,'DisplayName','TVE > 1%')
xline(X1(iResp(idx,1))+limit,'r','linewidth',2,'DisplayName','RT Limit%')

% set the x limits to 10 cycles from each xline
xl = [X1(iResp(idx,1))-10/self.F0, X1(iResp(idx,2))+10/self.F0];
xlim(xl);

title('Voltage Response Time')
% Create ylabel
ylabel('TVE (%)');
% Create xlabel
xlabel(xLabel);

hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);

end

%% -------------------------------------------------------------------------
function [respTime,fig] = plotStepCurrentResponseTime(self,T,limit,xLabel)
% Plot the Step Test Response Time

lstCurrent = ["TVE_IA" "TVE_IB" "TVE_IC" "TVE_Ip"];

% Create Figure
fig = figure;
set(fig,'Position',self.figPos)

%X and y Data
dT = mean(diff(T{:,1}));
X1 = (0:size(T,1)-1)*dT;

YMatrix1 = zeros(4,size(T,1));
for ii = 1:numel(lstCurrent)
    YMatrix1(ii,:) = T{:,lstCurrent(ii)};
end

% for each phase, find the point where the TVE first crosses 1%
nPhases = (size(YMatrix1,1));
iResp = zeros(nPhases,2);    % will hold indesex of 1% TVE crossing points
respTime = zeros(1,nPhases);
for i = 1:nPhases
    idx = find(YMatrix1(i,:)> 1);
    iResp(i,:) = [idx(1),idx(end)];
    respTime(1,i) = X1(iResp(i,end)) - X1(iResp(i,1));
end

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','TVE\_IA');
set(plot1(2),'DisplayName','TVE\_IB');
set(plot1(3),'DisplayName','TVE\_IC');
set(plot1(4),'DisplayName','TVE\_I+');

% Draw the limit line
yline(1,'-g','linewidth',2,'DisplayName','TVE Limit')

% Draw the response time lines on the longest resp time
[~,idx] = max(respTime);    % index of the maximum repsones time
respTime = array2table(respTime,'VariableNames',{'RT_IA','RT_IB','RT_IC','RT_I+'});
xline(X1(iResp(idx,1)),'b','linewidth',2,'DisplayName','TVE > 1%')
xline(X1(iResp(idx,1))+limit,'r','linewidth',2,'DisplayName','RT Limit%')

% set the x limits to 10 cycles from each xline
xl = [X1(iResp(idx,1))-10/self.F0, X1(iResp(idx,2))+10/self.F0];
xlim(xl);

title('Current Response Time')
% Create ylabel
ylabel('TVE (%)');
% Create xlabel
xlabel(xLabel);

hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);

end

%% -------------------------------------------------------------------------
function [respTime, fig] = plotStepFreqResponseTime(self,T,limit,xLabel)

% create figure
fig = figure;
set(fig,'Position',self.figPos)

%X and Y data
%X and y Data
dT = mean(diff(T{:,1}));
X1 = (0:size(T,1)-1)*dT;

Y = abs(T{:,'FE'});

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create |FE| line
plot1 = plot(X1,Y,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','|FE|');

yline(0.005,'-g','linewidth',2,'DisplayName','|FE| Limit')

% Find the response time start and end points
idx = find(Y > 0.005);
respTime = 0;
if ~isempty(idx)
    iResp = [idx(1),idx(end)];
    respTime = X1(iResp(end)-iResp(1));
    xline(X1(iResp(1)),'b','linewidth',2,'DisplayName','|FE| > 0.005 Hz')
    xline(X1(iResp(1))+limit,'r','linewidth',2,'DisplayName','RT Limit%')
    % set the x limits to 10 cycles from each xline
    xl = [X1(iResp(idx,1))-10/self.F0, X1(iResp(idx,2))+10/self.F0];
    xlim(xl);
end
respTime = array2table(respTime,'VariableNames',{'RT_|FE|'});

title('Frequency Response Time')
% Create ylabel
ylabel('|FE| (Hz)');
% Create xlabel
xlabel(xLabel);

hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);

end

%% -------------------------------------------------------------------------
function [respTime, fig] = plotStepROCOFResponseTime(self,T,limit,xLabel)

% create figure
fig = figure;
set(fig,'Position',self.figPos)

%X and Y data
%X and y Data
dT = mean(diff(T{:,1}));
X1 = (0:size(T,1)-1)*dT;

Y = abs(T{:,'RFE'});

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create |FE| line
plot1 = plot(X1,Y,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','|RFE|');

yline(0.1,'-g','linewidth',2,'DisplayName','|FE| Limit')

% Find the response time start and end points
idx = find(Y > 0.1);
respTime = 0;
if ~isempty(idx)
    iResp = [idx(1),idx(end)];
    respTime = X1(iResp(end)-iResp(1));
    xline(X1(iResp(1)),'b','linewidth',2,'DisplayName','|RFE| > 0.005 Hz/sec')
    xline(X1(iResp(1))+limit,'r','linewidth',2,'DisplayName','RT Limit%')
    % set the x limits to 10 cycles from each xline
    xl = [X1(iResp(idx,1))-10/self.F0, X1(iResp(idx,2))+10/self.F0];
    xlim(xl);
end
respTime = array2table(respTime,'VariableNames',{'RT_|RFE|'});

title('ROCOF Response Time')
% Create ylabel
ylabel('|RFE| (Hz)');
% Create xlabel
xlabel(xLabel);

hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);

end
    
%% ------------------------------------------------------------------------
function [Results, fig] = plotStepVoltageOvershoot(self,T,limit,xLabel)

% Plot the Voltage magnitude overshoot
lstVoltageReal = ["PMU_VA_real" "PMU_VB_real" "PMU_VC_real" "PMU_Vp_real"];
lstVoltageImag = ["PMU_VA_imag" "PMU_VB_imag" "PMU_VC_imag" "PMU_Vp_imag"];

% Create Figure
fig = figure;
set(fig,'Position',self.figPos)

%X and ////y Data
dT = mean(diff(T{:,1}));
X1 = (0:size(T,1)-1)*dT;

% Get the PMU values
PMUData = complex(zeros(4,size(T,1)));
for ii = 1:numel(lstVoltageReal)
    PMUData(ii,:) = abs(complex(T{:,lstVoltageReal(ii)},T{:,lstVoltageImag(ii)}));
end

lstVoltageReal = ["REF_VA_real" "REF_VB_real" "RE_ VC_real" "REF_Vp_real"];
lstVoltageImag = ["REF_VA_imag" "REF_VB_imag" "REF_VC_imag" "PMU_Vp_imag"];

% Get the Reference values
RefData = complex(zeros(4,size(T,1)));
for ii = 1:numel(lstVoltageReal)
    RefData(ii,:) = abs(complex(T{:,lstVoltageReal(ii)},T{:,lstVoltageImag(ii)}));
end

displayNames = {'PMU\_VA','PMU\_VB','PMU\_VC','PMU\_V+'};


[Results] = self.plotStepOsUs(X1,PMUData,RefData,displayNames,fig,limit);

end


