function plotResults(self,T,Lim,xLabel)

fig = plotVoltageTVE(self,T,Lim(1),xLabel(1));
self.hExcel.WriteData(fig,'B2')
close(fig)
fig = plotCurrentTVE(self,T,Lim(1),xLabel(2));
self.hExcel.WriteData(fig,'H2')
close(fig)
fig = plotFrequencyError(self,T,Lim(2),xLabel(2));
self.hExcel.WriteData(fig,'B16')
close(fig)
fig = plotRocofError(self,T,Lim(3),xLabel(2));
self.hExcel.WriteData(fig,'H16')
close(fig)

end

%% -------------------------------------------------------------------------
function fig = plotVoltageTVE(self,T,limit,xLabel)

lstVoltage = ["MaxTVE_VA" "MaxTVE_VB" "MaxTVE_VC" "MaxTVE_Vp"];

% Create figure
fig = figure;
set(fig,'Position',self.figPos);

% X and Y data
X1 = T{:,1};
YMatrix1 = zeros(4,size(T,1));
for ii = 1:numel(lstVoltage)
    YMatrix1(ii,:) = T{:,lstVoltage(ii)};
end


% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','MaxTVE\_VA');
set(plot1(2),'DisplayName','MaxTVE\_VB');
set(plot1(3),'DisplayName','MaxTVE\_VC');
set(plot1(4),'DisplayName','MaxTVE\_V+');

% Draw the limit line
yline(limit(1),'-r','linewidth',2,'DisplayName','TVE Limit')

title('Voltage TVE')
% Create ylabel
ylabel('TVE (%)');
% Create xlabel
xlabel(xLabel);

ylimit = max([max(max(YMatrix1')),(1.1*limit(1))]);
ylim(axes1,[0 ylimit]);
hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);
end

%% -------------------------------------------------------------------------
function fig = plotCurrentTVE(self,T,limit,xLabel)
% Create a formatted plot of the Current TVE

lstCurrent = ["MaxTVE_IA" "MaxTVE_IB" "MaxTVE_IC" "MaxTVE_Ip"];

% Create figure
fig = figure;
set(fig,'Position',self.figPos);

% X and Y date
X1 = T{:,1};
YMatrix1 = zeros(4,size(T,1));
for ii = 1:numel(lstCurrent)
    YMatrix1(ii,:) = T{:,lstCurrent(ii)};
end


% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','MaxTVE\_IA');
set(plot1(2),'DisplayName','MaxTVE\_IB');
set(plot1(3),'DisplayName','MaxTVE\_IC');
set(plot1(4),'DisplayName','MaxTVE\_I+');

% Draw the limit line
yline(limit(1),'-r','linewidth',2,'DisplayName','TVE Limit')

title('Current TVE')

% Create ylabel
ylabel('TVE (%)');
% Create xlabel
xlabel(xLabel);

ylimit = max([max(max(YMatrix1')),(1.1*limit(1))]);
ylim(axes1,[0 ylimit]);
hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6,'YGrid','on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);
end

%% ------------------------------------------------------------------------
% Create a formatted plot of the Frequency Error
function fig = plotFrequencyError(self,T,limit,xLabel)

lstFE = ["MinFE", "MaxFE"];

% Create figure
fig = figure;
set(fig,'Position',self.figPos);

% X and Y data
X1 = T{:,1};
YMatrix1 = zeros(2,size(T,1));
for ii = 1:numel(lstFE)
    YMatrix1(ii,:) = T{:,lstFE(ii)};
end

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'-b','LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','Max\_FE');
set(plot1(2),'DisplayName','Min\_FE');

% Draw the limit line
if isnumeric(limit)&& ~isinf(limit)
    limLine = yline(limit(1),'-r','linewidth',2,'DisplayName','FE Limit');
    yline(-limit(1),'-r','linewidth',2)
    lgd = legend([plot1(1),plot1(2), limLine]);
else
    lgd = legend([plot1(1),plot1(2)]);
end


title('Frequency Error')
% Create ylabel
ylabel('FE (Hz)');
% Create xlabel
xlabel(xLabel);

ylimit = max([max(max(abs(YMatrix1)')),(1.1*limit(1))]);
ylim(axes1,[-ylimit ylimit]);
hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);
end

%% ------------------------------------------------------------------------
% Create a formatted plot of the ROCOF Error
function fig = plotRocofError(self,T,limit,xLabel)

lstRFE = ["MinRFE", "MaxRFE"];

% Create figure
fig = figure;
set(fig,'Position',self.figPos);

% X and Y data
X1 = T{:,1};
YMatrix1 = zeros(2,size(T,1));
for ii = 1:numel(lstRFE)
    YMatrix1(ii,:) = T{:,lstRFE(ii)};
end

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

% Create multiple lines using matrix input to plot
plot1 = plot(X1,YMatrix1,'-b','LineWidth',2,'Parent',axes1);
set(plot1(1),'DisplayName','Max\_RFE');
set(plot1(2),'DisplayName','Min\_RFE');

% Draw the limit line
if isnumeric(limit)&& ~isinf(limit)
    limLine = yline(limit(1),'-r','linewidth',2,'DisplayName','RFE Limit');
    yline(-limit(1),'-r','linewidth',2)
    lgd = legend([plot1(1),plot1(2), limLine]);
else
    lgd = legend([plot1(1),plot1(2)]);
end

title('ROCOF Error')
% Create ylabel
ylabel('RFE (Hz/S)');
% Create xlabel
xlabel(xLabel);

ylimit = max([max(max(abs(YMatrix1)')),(1.1*limit(1))]);
ylim(axes1,[-ylimit ylimit]);
hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);
end

