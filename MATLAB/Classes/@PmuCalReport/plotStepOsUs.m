function [Results] = plotStepOsUs(self,X1,PMUData,RefData,displayNames,fig,limit)

% over/undershoot limit
if self.PmuClass == "M"; ouLimit = 0.1; else ouLimit = 0.05; end

nPhases = size(PMUData,1);

% Pre-allocate a results table
VarNames = {'phase','pre-over (%)','pre-under (%)','post-over (%)','post-under (%)','delay time (s)'};
VarTypes = {'string','double','double','double','double','double',};
Results = table('Size',[nPhases,numel(VarNames)],'VariableTypes',VarTypes);
Results.Properties.VariableNames = VarNames;

% Determination of overshoot and undershoot
% See IEC/IEEE 60255-118-1 Clause 5.2.4
% Step 1, use a histogram to determine the initial and final state levels

states = zeros(nPhases,2);
boundaries = zeros(nPhases,4);
preIdx = zeros(nPhases,1);
postIdx = zeros(nPhases,1);
for ii = 1:nPhases
    
    [N,~,bin] = histcounts(PMUData(ii,:),20);  % Histogram data without actually plotting
    %------
    % find out if it is a positive step or a negative step
    pos = true;    % assume positive step
    init = mean(bin(1:floor(length(bin)/2)));
    final = mean(bin(ceil(length(bin)/2):end));
    if init > final; pos = false; end
    
   [~,idx1] = max(N);   % the index of the largest bin
    state1 = bin==idx1; %
    state1 = mean(PMUData(ii,state1));
    
    N(idx1) = 0; % zero out the largest N
    [~,idx2] = max(N);  % index of the second largest bin
    state2 = bin==idx2;
    state2 = mean(PMUData(ii,state2));
    
   state = [state1, state2];
    
    % find out which state was highest
    high1 = true;
    if idx1 < idx2; high1 = false; end
    if (high1 && pos) || (~high1 && ~pos)
        state = flip(state);    % flip the state if needed;
    end  
    states(ii,:) = state; % states contains the initial and final state values
    %-----
     
    % next, get the state boundaries
    stepAmpl = abs(diff(state));  % the size of the step
    boundaries(ii,:) = [state(1)-ouLimit*stepAmpl, ...
                        state(1)+ouLimit*stepAmpl, ...
                        state(2)-ouLimit*stepAmpl, ...
                        state(2)+ouLimit*stepAmpl];
                    
    % over and undershoot are evaluated in the pre-transition and post-transition regions
    
    % pre-transiton ends when the initial state passes the boundary for the last time
    if pos % positive step
        %preBound = boundaries(2); 
        %inBound = YMatrix1(ii,:) < boundaries(ii,2);
        preIdx(ii) = find(PMUData(ii,:) < boundaries(ii,2),1,'last');
    else  % negative step
        %preBound = boundaries(1); 
        %inBound = YMatrix1(ii,:) > boundaries(ii,1);
        preIdx(ii) = find(PMUData(ii,:) > boundaries(ii,2),1,'last');
    end
    
    % post-transition begins when the final state enters the boundary for the first time
    if pos % positive step
        postIdx(ii) = find(PMUData(ii,:) > boundaries(ii,3),1,'first');
    else % negative step
        postIdx(ii) = find(PMUData(ii,:) < boundries(ii,4),1,'first');
    end    
     
    % difference in % of step size between the value an the state levels
    % pre-transition
    preDiff = ((PMUData(ii,1:preIdx(ii))-states(ii,1))/stepAmpl)*100;
    postDiff = ((PMUData(ii,postIdx(ii):end)-states(ii,2))/stepAmpl)*100;
   
    % Over/undershoot Results
    Results{ii,'pre-over (%)'} = max(preDiff);
    Results{ii,'pre-under (%)'} = abs(min(preDiff));
    Results{ii,'post-over (%)'} = max(postDiff);
    Results{ii,'post-under (%)'} = abs(min(postDiff));
    
    % Delay Time is the difference between the Refeence transition time and the 50%
    % point in the rise
    if pos
        refRiseIdx = find(RefData(ii,:) > states(ii,1)+0.5*stepAmpl,1,'first');
        pmuRiseIdx = find(PMUData(ii,:) > states(ii,1)+0.5*stepAmpl,1,'first');
    else
        refRiseIdx = find(RefData(ii,:) < states(ii,1)+0.5*stepAmpl,1,'first');
        pmuRiseIdx = find(PMUData(ii,:) < states(ii,1)+0.5*stepAmpl,1,'first');
    end
    % Interpolate the position of the 50% crossing
    interp = (PMUData(ii,pmuRiseIdx) - PMUData(ii,pmuRiseIdx-1))...
              /(states(ii,1)+stepAmpl - PMUData(ii,pmuRiseIdx-1));
    pmuCross = X1(pmuRiseIdx-1)+(interp*(X1(pmuRiseIdx)-X1(pmuRiseIdx-1)));
    Results{ii,'delay time (s)'} = pmuCross - X1(refRiseIdx);
    

end     
    
% Prepare a plot

% Normalize the data
YNorm = zeros(size(PMUData));
for ii = 1:nPhases
    YNorm(ii,:) = (PMUData(ii,:)-states(ii,1))/(states(ii,2)-states(ii,1))*100;
end   

% Normalize the time scale
XNorm = X1 - X1(refRiseIdx);

% Create axes
axes1 = axes('Parent',fig,...
    'Position',self.axPos, 'YGrid', 'on');
hold(axes1,'on');

plot1 = plot(XNorm,YNorm);
set(plot1(1),'DisplayName',displayNames{1});
set(plot1(2),'DisplayName',displayNames{2});
set(plot1(3),'DisplayName',displayNames{3});
set(plot1(4),'DisplayName',displayNames{4});

% create a set of boundary limit lines
preX = XNorm(1:preIdx(1));
postX = XNorm(postIdx(1):end);
preBounds = [zeros(1,length(preX))-ouLimit.*100;zeros(1,length(preX))+ouLimit.*100]+100*~pos;
postBounds = [zeros(1,length(postX))-ouLimit.*100;zeros(1,length(postX))+ouLimit.*100]+100*pos;
plot2 = plot(preX,preBounds,'-r');
set(plot2(1),'DisplayName','O/U boundary');
set(plot2(2),'HandleVisibility','off');
plot3 = plot(postX,postBounds,'-r');
set(plot3(1),'HandleVisibility','off'); 
set(plot3(2),'HandleVisibility','off'); 

% draw vertical lines at the step time and the delay time limit
plot4(1) = line([0,0],[20,80]);
set(plot4(1),'DisplayName','Step time')
plot4(2) = line([limit,limit],[20,80]);
set(plot4(2),'DisplayName','Delay time limit')
set(plot4(2),'Color','magenta')

hold(axes1,'off');
% Set the remaining axes properties
set(axes1,'FontSize',6, 'YGrid', 'on');
% Create legend
lgd = legend(axes1,'show');
set(lgd,'Location','eastoutside');
set (lgd,'Position',self.lgdPos);

xlim([-7/self.F0,7/self.F0])


end
    