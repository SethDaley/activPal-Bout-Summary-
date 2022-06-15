%% ActivPal Breakdown V2.5
% May 2022
% Version history:
    % 2.0 - No longer uses the epoch file to calculate bout length. Instead
    % the associated events file is loaded and the events associated with
    % the identified days are isolated and catagorized.
    % 2.1 - Option added to ignore the data from a final "half day" of data
    % collected before the activePAL was turned in
    % 2.2 - Progress bar added during Julian time conversion
    % 2.2.1 - Fixed an error where the sleep-time was not correctly
    % organized for display when a half day is not present.
    % 2.3 - Adjusted the MET threshold from 19 to 23 to account for the
    % potential high initial met level before the start of day 1
    % 2.4 - Option to reselect day start and end times without restarting
    % the program.
    % 2.5 - Added check to allow the user to examine the data before
    % entering the number of days in the collection period
% W Seth Daley
% ActivPal output epoch and events files required for this analysis
% Imported data is left in original form in a raw. variable. 
% Organizational variables are stored in an org. structured variable
% Processed data is stored at various milestones in the data. structured
% variable
% Outputs are stored as tables in Output, as well as exported to a
% specified excel file.

% Copyright 2021, 2022 W Seth Daley
% 
% Licensed under the Apache License, Version 2.0 (the "License");
% you may not use this file except in compliance with the License.
% You may obtain a copy of the License at
% 
%    http://www.apache.org/licenses/LICENSE-2.0
% 
% Unless required by applicable law or agreed to in writing, software
% distributed under the License is distributed on an "AS IS" BASIS,
% WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
% See the License for the specific language governing permissions and
% limitations under the License. 

%% Import ActivPal epoch datafile

clc
clear

% User selects the file to be analyzed
[org.epochFileName,org.fileLoc] = uigetfile({'*.csv',...
    'CSV Files (*.csv)';'*.*','All Files'},...
    'Select the epoch to be analyzed');

% Create a temp variable to store the epoch file location and name
tempEpochName = strcat(org.fileLoc,org.epochFileName);

% Load the data file parameters
opts = detectImportOptions(tempEpochName);

% Load the epoch data file
raw.epochData = readmatrix(tempEpochName,opts);

% Ask the user if they need to check the MET data 
inq = ['Do you need to check the MET file to determine the'...
    ' number of days?'];
resp = questdlg(inq,'Day check','Yes','No','No');

switch resp
    case 'Yes'
        dayCheck = 1;
        close
    case 'No'
        dayCheck = 0;
        close
end

if dayCheck == 1
    scr_siz = get(0,'ScreenSize');
    
    g = figure;
    g.Position = [100 100 scr_siz(3)*.9 scr_siz(4)*.7];
    plot(raw.epochData(:,3))
    title('Check for the number of days');
end

clear g scr_siz dayCheck inq resp

% Gather trial data from user

x = inputdlg({'Number of full days wearing ActivPal',...
    'Does the data include a final half day? (y/n)'},...
    'Input',[1 50],{'5','y'});

close

% Store the inputted trial data
org.numDays = str2double(x(1));

if x{2} == 'y'
    org.halfDay = true;
    org.numDays = org.numDays + 1;
elseif x{2} == 'n'
    org.halfDay = false;
else
    msg = ['Half-day answer must be in the form of y/n.'...
        ' Please re-run the program'];
    warningMessage = sprintf(msg);
    uiwait(msgbox(warningMessage));
    return
end

% Modify the filename to produce the default name for the events file
org.rootFileName = erase(org.epochFileName,' by 15s epoch.csv');
org.eventFileName = strcat(org.rootFileName,' Events.csv');

% Create a temp variable to store the events file location and name
tempEventName = strcat(org.fileLoc,org.eventFileName);

% Check if the events file exists
if isfile(tempEventName)
    % If the events file exists
    % Load the data file parameters
    opts = detectImportOptions(tempEventName);

    % Load the data file
    raw.eventData = readmatrix(tempEventName,opts);
else
     % File does not exist.
     warningMessage = sprintf('Warning: Event file does not exist:\n%s',...
         tempEventName);
     uiwait(msgbox(warningMessage));
     return
end



clear temp* ans opts x msg
%% Create variables for Julian date in raw files

% Loop over the epoch file daya
wt = 'Please wait while Julian time is convereted from the Epoch file';
f = waitbar(0,wt);
n = length(raw.epochData);
x = 0;
for i=1:length(raw.epochData)
    
    x=x+1;
    step = x/n;
    waitbar(step);
    % Convert and store
    raw.epochDateData(i) = datetime(raw.epochData(i,1),...
        'ConvertFrom','excel');
end
close(f)
clear i n x f step wt

% Loop over the events file data
wt = 'Please wait while Julian time is convereted from the Events file';
f = waitbar(0,wt);
n = length(raw.eventData);
x = 0;
for i=1:length(raw.eventData)
    x=x+1;
    step = x/n;
    waitbar(step);
    raw.eventDateData(i) = datetime(raw.eventData(i,1),...
        'ConvertFrom','excel');
end
close(f)
clear i n x f step wt
%% Separate datafile and time markers into days


check = 0;

while check == 0

    % Set the number of locations as twice the number of days
    n = org.numDays*2;
    
    % Determine the size of the user's screen
    scr_siz = get(0,'ScreenSize');
    
    % Plot the activity waveform in a large figure
    f = figure;
    f.Position = [100 100 scr_siz(3)*.9 scr_siz(4)*.7];
    plot(raw.epochDateData,raw.epochData(:,3))
    tit = ['Select the start and end of each day. Space between'...
        ' vertical selection line and start of day is acceptable. Treat'...
        ' any half-day as a full day for this step.'];
    title(tit);
    ax = gca;
    % Store the user-selected timepoints
    [x,~] = ginput(n);
    
    for i=1:length(x)
        xdate(i) = num2ruler(x(i),ax.XAxis); 
        [~,org.dayTimes(i)] = min(abs(raw.epochDateData - xdate(i)));
    end
    
    % Close the figure
    close(f);
    
    % Clear temporary variables
    clear f y n scr_siz ax x xdate tit
    
    % Store the day time points in a temporary variable
    tempTimes = org.dayTimes;
    
    % Separating out each days data based on user inputs
    % Loop over the number of days
    for i=1:org.numDays
        
        % Store the name of the day in a temp variable
        day = ['Day_',num2str(i,'%01.0f')];
        
        % Takes the first time marked as the start of the day
        dayStart = round(tempTimes(1));
        
        % Takes the second time marked as the end
        dayEnd = round(tempTimes(2));
        
        % Cuts the day's worth of data from the raw file
        tempData = raw.epochData(dayStart:dayEnd,:);
        
        % Stores the Day's data in a structured variable
        data.raw.(day) = tempData;
        
        % Chops the used time markers
        tempTimes(1:2) = [];
        
        % Clear Temp variables used within the loop
        clear day* tempData
    end
    
    
    % Organizing the Day times
    % Loop over number of days
    for i=1:org.numDays
        
        % Store the name of the day in a temp variable
        day = ['Day_',num2str(i,'%01.0f')];
        
        % Create temporary variables for the start and end points of each
        % day
        tempStart = round(org.dayTimes(2*i-1));
        tempStop = round(org.dayTimes(2*i));
        
        % Store the stand and end points in structured variable
        data.Times.DayStart.raw.(day) = tempStart;
        data.Times.DayEnd.raw.(day) = tempStop;
        
        clear day tempS*
    end
    
    % Clear temp variables
    clear i tempTimes
    
    
    
    for i=1:org.numDays
        
        % Set the MET activity threshold
        m = 23;
        
        % Store the name of the day in a temp variable
        day = ['Day_',num2str(i,'%01.0f')];
        
        % Store the untrimmed data in a temp variable
        tempdata = data.raw.(day);
        
        % Find the first point when the MET exceeds threshold and store
        cutFromFront = find(tempdata(:,3)>m,1);
        
        % Find the last point when the met exceeds threshold and store
        cutFromEnd = find(tempdata(:,3)>m,1,'last');
        
        % Use the new time locations to trim the day data and store
        tempTrimmedData = tempdata(cutFromFront:cutFromEnd,:);
        
        % Assign the trimmed data to a structured variable
        data.trimmed.(day) = tempTrimmedData;
        
        % Modify the Day start and end times by the amount trimmed 
        data.Times.DayStart.trimmed.(day) =...
            data.Times.DayStart.raw.(day) + cutFromFront;
        data.Times.DayEnd.trimmed.(day) =...
            data.Times.DayEnd.raw.(day) - (length(tempdata)-cutFromEnd);
        
        % Create a variable with the modified day start and end times for
        % user conformation
        
        org.newtimes((2*i)-1) = org.dayTimes((2*i)-1) + cutFromFront;
        org.newtimes(2*i) = org.dayTimes(2*i) -...
            (length(tempdata)-cutFromEnd);
        
        clear temp* cut* day
    end
    
    
    % Determine the size of the user's screen
    scr_siz = get(0,'ScreenSize');
    
    % Create vector to color the lines 
    % *** If the data is > 12 days long the code will not work and the 
    % length of this variable should be adjusted***
    if (org.numDays > 12)
        er = ['The code does not currently support data files longer'...
            ' than 12 days'];
        error(er)
    else
        tempColour = ['y','y','m','m','c','c','r','r','g','g','b','b',...
            'y','y','m','m','c','c','r','r','g','g','b','b'];
    end

    % Plot the activity waveform with the adjusted day marks for user input
    f = figure;
    f.Position = [100 100 scr_siz(3)*.9 scr_siz(4)*.7];
    plot(raw.epochDateData,raw.epochData(:,3))
    title('Day marks - Adjusted');
    
    hold on
    
    % Loop over the number of day markers
    for i=1:length(org.newtimes)
        
        % Draw a line on the figure to indicate each new position
        xline(raw.epochDateData(org.newtimes(i)),tempColour(i),...
            'LineWidth',2);
    end
    hold off
    tempName = strcat(org.fileLoc,'dayCheck.png');
    exportgraphics(f,tempName,'Resolution',900)

    quest = ['Check that start and end times are acceptable either on the' ...
        ' displayed figure or on the high-res saved version.'];
    answer = questdlg(quest,'Day Confirmation','Correct','Re-identify',...
        'Correct');
    
    switch answer
        case 'Correct'
            uiwait(msgbox("Day markers accepted, continuing analysis"));
            check = 1;
            close
        case 'Re-identify'
            uiwait(msgbox("Re-select start and end times"));
            check = 0;
            close
            clear i m f scr_siz tempColour h
    end

end

clear i m f scr_siz temp* h quest answer check er


%% Calculate day and sleep length

% Loop over the number of days
for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Assign the start location for the day to a temporary variable
    tempStart = data.Times.DayStart.trimmed.(day);
    
    % Assign the end location for the day to a temporary variable
    tempEnd = data.Times.DayEnd.trimmed.(day);
    
    % Find the difference between the start and end, and assign
    tempLength = tempEnd - tempStart;
    
    % Find the day length by multiplying the number of samples long the day
    % is by 15 seconds/sample and convert to time
    tempDay = duration(0,0,(tempLength * 15));
    
    % Assign the time to a structured variable
    data.Times.DayLength.(day) = tempDay;
    
    clear temp* day
end

clear i

for i=1:org.numDays-1
    
    % Store the name of the first day in a temp variable
    dayA = ['Day_',num2str(i,'%01.0f')];
    
    % Store the name of the second day in a temp variable
    dayB = ['Day_',num2str(i+1,'%01.0f')];
    
    % Assign the end of the first day to a temp variable
    tempStart = data.Times.DayEnd.trimmed.(dayA);
    
    % Assign the start of the second day to a temp variable
    tempEnd = data.Times.DayStart.trimmed.(dayB);
    
    % Find the difference between the two points
    tempLength = tempEnd - tempStart;
    
    % Convert the difference into time
    tempNight = duration(0,0,(tempLength * 15));
    
    % Store the time in a structured variable
    data.Times.SleepLength.(dayA) = tempNight;
    
    clear temp* day*
end

clear i

%% Seperate events data into days

for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Load the start and end point of the day into temp variables
    tempStart = data.Times.DayStart.trimmed.(day);
    tempEnd = data.Times.DayEnd.trimmed.(day);
    
    % Use the start and end points to find the associated time codes
    tempStartTime = raw.epochData(tempStart,1);
    tempEndTime = raw.epochData(tempEnd,1);
    
    % Use the time codes to identify the cooorespinding points in the
    % events file
    tempEventStart = find(raw.eventData(:,1) > tempStartTime,1);
    tempEventEnd = find(raw.eventData(:,1) > tempEndTime,1)-1;
    
    % Use the events times to trim the data for the day
    data.events.(day) = raw.eventData(tempEventStart:tempEventEnd,:);
    
    clear temp* day
end
clear i

%% Isolating behaiviour

for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')];   
    
    tempData = data.events.(day);
    tempSedBout = [];
    tempStandBout = [];
    tempStepBout = [];
    
    for j=1:length(tempData)
        if tempData(j,4) == 0
            tempSedBout = [tempSedBout tempData(j,3)];
        elseif tempData(j,4) == 1
            tempStandBout = [tempStandBout tempData(j,3)]; 
        elseif tempData(j,4) == 2
            tempStepBout = [tempStepBout tempData(j,3)]; 
        end
    end
    clear j
    
    
    data.totals.sedentary.(day) = tempSedBout;
    data.totals.standing.(day) = tempStandBout;
    data.totals.stepping.(day) = tempStepBout;
    
    clear temp* day
end
clear i

%% Categorize Bouts


for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')];   
    
    % Load Sedentary bouts
    boutSedentary = data.totals.sedentary.(day);
    
    % Initialize structured variable to store categorized bouts
    tempBouts.less10min = [];
    tempBouts.less15min = [];
    tempBouts.btwn10_20min = [];
    tempBouts.btwn20_30min = [];
    tempBouts.btwn30_40min = [];
    tempBouts.btwn40_50min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn1_2hr = [];
    tempBouts.btwn2_3hr = [];
    tempBouts.btwn3_4hr = [];
    tempBouts.grtr4hr = [];
    tempBouts.less30min = [];
    tempBouts.grtr30min = [];
    
    % Loop over the length of the unadjusted sedentary bouts
    for l=1:length(boutSedentary)
        
        % Load individual bout
        tempSample = boutSedentary(l);
        
        % If bout is < 900 seconds, categorize as < 15 minutes
        if tempSample < 900
            
            % Append the bout to the < 15 min variable
            tempBouts.less15min = [tempBouts.less15min tempSample];
        end
        
        % If bout is <= 1800 seconds, categorize as < 30 minutes 
        if tempSample <= 1800
            
            % Append the bout to the < 30 min variable
            tempBouts.less30min = [tempBouts.less30min tempSample];
        end
        
        % If bout is > 1800 seconds, categorize as > 30 minutes
        if tempSample > 1800
            
            % Append the bout to the > 30 min variable
            tempBouts.grtr30min = [tempBouts.grtr30min tempSample];
        end
        
        % If the bout is less than 600 seconds
        if tempSample < 600
            % Append the bout to the < 10min variable
            tempBouts.less10min = [tempBouts.less10min tempSample];
        
        % If the bout is between 600 (inclusive) and 1200 seconds
        elseif (tempSample >= 600) && (tempSample < 1200)
            % Append the bout to the between 10 and 20min variable
            tempBouts.btwn10_20min = [tempBouts.btwn10_20min tempSample];
        
        % If the bout is between 1200 (inclusive) and 1800 seconds 
        elseif (tempSample >= 1200) && (tempSample < 1800)
            % Append the bout to the between 20 and 30min variable
            tempBouts.btwn20_30min = [tempBouts.btwn20_30min tempSample];
        
        % If the bout is between 1800 (inclusive) and 2400 seconds
        elseif (tempSample >= 1800) && (tempSample < 2400)
            % Append the bout to the between 30 and 40min variable
            tempBouts.btwn30_40min = [tempBouts.btwn30_40min tempSample];
        
        % If the bout is between 2400 (inclusive) and 3000 seconds
        elseif (tempSample >= 2400) && (tempSample < 3000)
            % Append the bout to the between 40 and 50min variable
            tempBouts.btwn40_50min = [tempBouts.btwn40_50min tempSample];
        
        % If the bout is between 3000 (inclusive) and 3600 seconds
        elseif (tempSample >= 3000) && (tempSample < 3600)
            % Append the bout to the between 50 and 60min variable
            tempBouts.btwn50_60min = [tempBouts.btwn50_60min tempSample];
        
        % If the bout is between 3600 (inclusive) and 7200 seconds
        elseif (tempSample >= 3600) && (tempSample < 7200)
            % Append the bout to the between 1 and 2hour variable
            tempBouts.btwn1_2hr = [tempBouts.btwn1_2hr tempSample];
        
        % If the bout is between 7200 (inclusive) and 14400 seconds
        elseif (tempSample >= 7200) && (tempSample < 14400)
            % Append the bout to the between 2 and 3hour variable
            tempBouts.btwn2_3hr = [tempBouts.btwn2_3hr tempSample];
        
        % If the bout is between 14400 (inclusive) and 21600 seconds
        elseif (tempSample >= 14400) && (tempSample < 21600)
            % Append the bout to the between 3 and 4hour variable
            tempBouts.btwn3_4hr = [tempBouts.btwn3_4hr tempSample];
        
        % If the bout is greater than 21600 seconds(inclusive)
        elseif tempSample >= 21600
            % Append the bout to the > 4hour variable
            tempBouts.grtr4hr = [tempBouts.grtr4hr tempSample];
        end
        
        clear tempSample

    end
    
    % Store the sorted bouts
    data.bouts.(day).sedentaryBouts = tempBouts;
    clear l temp* boutSedentary
   
end

for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')]; 
    
    
    % Standing bouts
    % Reinitialize structured variable to store categorized bouts
    
    tempBouts.less10min = [];
    tempBouts.less15min = [];
    tempBouts.btwn10_20min = [];
    tempBouts.btwn20_30min = [];
    tempBouts.btwn30_40min = [];
    tempBouts.btwn40_50min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn1_2hr = [];
    tempBouts.btwn2_3hr = [];
    tempBouts.btwn3_4hr = [];
    tempBouts.grtr4hr = [];
    tempBouts.less30min = [];
    tempBouts.grtr30min = [];
    
    boutStanding = data.totals.standing.(day);
    
    % Unadjusted Bouts
    for l=1:length(boutStanding)
        
        tempSample = boutStanding(l);
        
        if tempSample < 900
            tempBouts.less15min = [tempBouts.less15min tempSample];
        end
        if tempSample <= 1800
            tempBouts.less30min = [tempBouts.less30min tempSample];
        end
        if tempSample > 1800
            tempBouts.grtr30min = [tempBouts.grtr30min tempSample];
        end
        
        if tempSample < 600
            tempBouts.less10min = [tempBouts.less10min tempSample];
        elseif (tempSample >= 600) && (tempSample < 1200)
            tempBouts.btwn10_20min = [tempBouts.btwn10_20min tempSample];
        elseif (tempSample >= 1200) && (tempSample < 1800)
            tempBouts.btwn20_30min = [tempBouts.btwn20_30min tempSample];
        elseif (tempSample >= 1800) && (tempSample < 2400)
            tempBouts.btwn30_40min = [tempBouts.btwn30_40min tempSample];
        elseif (tempSample >= 2400) && (tempSample < 3000)
            tempBouts.btwn40_50min = [tempBouts.btwn40_50min tempSample];
        elseif (tempSample >= 3000) && (tempSample < 3600)
            tempBouts.btwn50_60min = [tempBouts.btwn50_60min tempSample];
        elseif (tempSample >= 3600) && (tempSample < 7200)
            tempBouts.btwn1_2hr = [tempBouts.btwn1_2hr tempSample];
        elseif (tempSample >= 7200) && (tempSample < 14400)
            tempBouts.btwn2_3hr = [tempBouts.btwn2_3hr tempSample];
        elseif (tempSample >= 14400) && (tempSample < 21600)
            tempBouts.btwn3_4hr = [tempBouts.btwn3_4hr tempSample];
        elseif tempSample >= 21600
            tempBouts.grtr4hr = [tempBouts.grtr4hr tempSample];
        end
        
        clear tempSample
    end
    
    data.bouts.(day).standingBouts = tempBouts;
    clear temp* l boutStanding
    
end

for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')]; 
    
    
    % Stepping bouts
    
    % Reinitialize structured variable to store categorized bouts
    
    tempBouts.less10min = [];
    tempBouts.less15min = [];
    tempBouts.btwn10_20min = [];
    tempBouts.btwn20_30min = [];
    tempBouts.btwn30_40min = [];
    tempBouts.btwn40_50min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn50_60min = [];
    tempBouts.btwn1_2hr = [];
    tempBouts.btwn2_3hr = [];
    tempBouts.btwn3_4hr = [];
    tempBouts.grtr4hr = [];
    tempBouts.less30min = [];
    tempBouts.grtr30min = [];
    
    boutStepTime = data.totals.stepping.(day);
    
    for l=1:length(boutStepTime)
        
        tempSample = boutStepTime(l);
        
        if tempSample < 900
            tempBouts.less15min = [tempBouts.less15min tempSample];
        end
        if tempSample <= 1800
            tempBouts.less30min = [tempBouts.less30min tempSample];
        end
        if tempSample > 1800
            tempBouts.grtr30min = [tempBouts.grtr30min tempSample];
        end
        
        if tempSample < 600
            tempBouts.less10min = [tempBouts.less10min tempSample];
        elseif (tempSample >= 600) && (tempSample < 1200)
            tempBouts.btwn10_20min = [tempBouts.btwn10_20min tempSample];
        elseif (tempSample >= 1200) && (tempSample < 1800)
            tempBouts.btwn20_30min = [tempBouts.btwn20_30min tempSample];
        elseif (tempSample >= 1800) && (tempSample < 2400)
            tempBouts.btwn30_40min = [tempBouts.btwn30_40min tempSample];
        elseif (tempSample >= 2400) && (tempSample < 3000)
            tempBouts.btwn40_50min = [tempBouts.btwn40_50min tempSample];
        elseif (tempSample >= 3000) && (tempSample < 3600)
            tempBouts.btwn50_60min = [tempBouts.btwn50_60min tempSample];
        elseif (tempSample >= 3600) && (tempSample < 7200)
            tempBouts.btwn1_2hr = [tempBouts.btwn1_2hr tempSample];
        elseif (tempSample >= 7200) && (tempSample < 14400)
            tempBouts.btwn2_3hr = [tempBouts.btwn2_3hr tempSample];
        elseif (tempSample >= 14400) && (tempSample < 21600)
            tempBouts.btwn3_4hr = [tempBouts.btwn3_4hr tempSample];
        elseif tempSample >= 21600
            tempBouts.grtr4hr = [tempBouts.grtr4hr tempSample];
        end
        

    end
    data.bouts.(day).StepTimeBouts = tempBouts;
    clear temp* l boutStepTime

end
clear temp* i index* day

% total counts of sedentary to upright
for i=1:org.numDays
    
    % Store the name of the day in a temp variable
    day = ['Day_',num2str(i,'%01.0f')];

    temp = data.trimmed.(day)(:,7);
    
    data.totals.sed2up.(day) = sum(temp);
    
    clear day temp
end
clear i


%% Creating output file for event file data

%overall file

% Create a variable storing row names
org.rowNames = [];


if org.halfDay
    org.displayDays = org.numDays - 1;
else
    org.displayDays = org.numDays;
end

for i=1:org.displayDays
    
    day = convertCharsToStrings(['Day ',num2str(i,'%01.0f')]);
    
    org.rowNames = [org.rowNames day];
    
    clear day
end
clear i

% Initialize temporary variable
temp = zeros(org.displayDays,10);

for i=1:org.displayDays
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Column 1: Day length in minutes
    temp(i,1) = minutes(data.Times.DayLength.(day));
    % Column 2: Day length in hours
    temp(i,2) = hours(data.Times.DayLength.(day));
    
    % Column 5: Step count
    temp(i,5) = sum(data.trimmed.(day)(:,2));
    % Column 6: Step time
    temp(i,6) = round(sum(data.trimmed.(day)(:,6))/60,2);
    % Column 7: Standing time
    temp(i,7) = round(sum(data.trimmed.(day)(:,5))/60,2);
    % Column 8: Sedentary time
    temp(i,8) = round(sum(data.trimmed.(day)(:,4))/60,2);
    % Column 9: Sedentary to upright time
    temp(i,9) = round(sum(data.trimmed.(day)(:,7))/60,2);
    % Column 10: Upright to sedentary time
    temp(i,10) = round(sum(data.trimmed.(day)(:,8))/60,2);
    
    % Clear temporary variable
    clear day
end

% Clear indexing variable
clear i

if org.halfDay == 0
    x = 1;
else
    x = 0;
end

for i=1:org.displayDays-x
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Column 3: Sleep length in minutes
    temp(i,3) = minutes(data.Times.SleepLength.(day));
    % Column 4: Sleep length in hours
    temp(i,4) = hours(data.Times.SleepLength.(day));
    
    % Clear temporary variable
    clear day
end

% Clear indexing variable
clear i x

% Format Summary output table
Output.SummaryOutput = table(temp(:,1),temp(:,2),temp(:,3),temp(:,4),...
    temp(:,5),temp(:,6),temp(:,7),temp(:,8),temp(:,9),temp(:,10),...
    'VariableNames',{'Time Awake(m)','Time Awake(h)','Time Asleep (m)',...
    'Time Asleep (h)','Step Count','Time Stepping (m)',...
    'Time Standing (m)','Time Sedentary (m)','Time Sed. to Up (m)',...
    'Time Up to Sed. (m)'},'RowNames',org.rowNames);

% Clear temporary variable
clear temp



% Overall output

standing = [];
stepping = [];
sedentary = [];
sed2up = [];
count = 0;
total = 0;

for i=1:org.displayDays
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    standing = [standing data.totals.standing.(day)];
    stepping = [stepping data.totals.stepping.(day)];
    sedentary = [sedentary data.totals.sedentary.(day)];
    
    tempHour = hours(data.Times.DayLength.(day));
    tempSed2upcount = data.totals.sed2up.(day);
    tempSed2up = tempSed2upcount/tempHour;
    
    sed2up = [sed2up tempSed2up];
    
    count = count + length(data.totals.sedentary.(day));
    total = total + sum(data.totals.sedentary.(day));
    
    clear day temp*
end
clear i

temp = zeros(2,4);

temp(1,1) = mean(standing/60);
temp(1,2) = mean(stepping/60);
temp(1,3) = mean(sedentary/60);
temp(1,4) = mean(sed2up);

temp(2,1) = std(standing/60);
temp(2,2) = std(stepping/60);
temp(2,3) = std(sedentary/60);
temp(2,4) = std(sed2up);

clear count total standing stepping sedentary sed2up

Output.OverallOutput = table(temp(:,1),temp(:,2),temp(:,3),temp(:,4),...
    'VariableNames',{'Standing (min)','Stepping (min)',...
    'Sedentary (min)','Times Break up'},'RowNames',{'Mean','Std. Dev.'});

clear temp





% Sedentary Bout Count

% Initialize temporary variable
temp = zeros(org.displayDays+2,13);

% Loop over the number of days
for i=1:org.displayDays
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Populate the columns with data based on bout size
    % Column 1: < 10 min
    temp(i,1) = length(data.bouts.(day).sedentaryBouts.less10min);
    % Column 2: < 15 min
    temp(i,2) = length(data.bouts.(day).sedentaryBouts.less15min);
    % Column 3: 10-20 min
    temp(i,3) = length(data.bouts.(day).sedentaryBouts.btwn10_20min);
    % Column 4: 20-30 min
    temp(i,4) = length(data.bouts.(day).sedentaryBouts.btwn20_30min);
    % Column 5: 30-40 min
    temp(i,5) = length(data.bouts.(day).sedentaryBouts.btwn30_40min);
    % Column 6: 40-50 min
    temp(i,6) = length(data.bouts.(day).sedentaryBouts.btwn40_50min);
    % Column 7: 50-60 min
    temp(i,7) = length(data.bouts.(day).sedentaryBouts.btwn50_60min);
    % Column 8: 1-2 hours
    temp(i,8) = length(data.bouts.(day).sedentaryBouts.btwn1_2hr);
    % Column 9: 2-3 hours
    temp(i,9) = length(data.bouts.(day).sedentaryBouts.btwn2_3hr);
    % Column 10: 3-4 hours
    temp(i,10) = length(data.bouts.(day).sedentaryBouts.btwn3_4hr);
    % Column 11: 4+ hours
    temp(i,11) = length(data.bouts.(day).sedentaryBouts.grtr4hr);
    % Column 12: < 30 min
    temp(i,12) = length(data.bouts.(day).sedentaryBouts.less30min);
    % Column 13: > 30 min
    temp(i,13) = length(data.bouts.(day).sedentaryBouts.grtr30min);
    
    % Clear temporary variable
    clear day
end

% Clear indexing variable
clear i

% Loop over number of Rows
for i=1:13
    
    % Mean for each column
    temp(org.displayDays+1,i)=mean(temp(1:org.displayDays,i));
    % St. Dev. for each colunm
    temp(org.displayDays+2,i)=std(temp(1:org.displayDays,i));
    
end

% Clear indexing variable
clear i

tempRows = [org.rowNames "Mean" "St. Dev."];

% Format Sedentary bout count table
Output.SedBoutCountOutput = table(temp(:,1),temp(:,2),temp(:,3),...
    temp(:,4),temp(:,5),temp(:,6),temp(:,7),temp(:,8),temp(:,9),...
    temp(:,10),temp(:,11),temp(:,12),temp(:,13),'VariableNames',...
    {'< 10 mins','< 15 mins','10-20 mins','20-30 mins','30-40 mins',...
    '40-50 mins','50-60 mins','1-2 hours','2-3 hours','3-4 hours',...
    '> 4 hours','< 30 mins','> 30 mins'},'RowNames',tempRows);

% Clear temporary variable
clear temp*






% Stepping Bout Count

% Initialize temporary variable
temp = zeros(org.displayDays+2,11);

% Loop over the number of days
for i=1:org.displayDays
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Populate the columns with data based on bout size
    % Column 1: < 10 min
    temp(i,1) = length(data.bouts.(day).StepTimeBouts.less10min);
    % Column 2: < 15 min
    temp(i,2) = length(data.bouts.(day).StepTimeBouts.less15min);
    % Column 3: 10-20 min
    temp(i,3) = length(data.bouts.(day).StepTimeBouts.btwn10_20min);
    % Column 4: 20-30 min
    temp(i,4) = length(data.bouts.(day).StepTimeBouts.btwn20_30min);
    % Column 5: 30-40 min
    temp(i,5) = length(data.bouts.(day).StepTimeBouts.btwn30_40min);
    % Column 6: 40-50 min
    temp(i,6) = length(data.bouts.(day).StepTimeBouts.btwn40_50min);
    % Column 7: 50-60 min
    temp(i,7) = length(data.bouts.(day).StepTimeBouts.btwn50_60min);
    % Column 8: 1-2 hours
    temp(i,8) = length(data.bouts.(day).StepTimeBouts.btwn1_2hr);
    % Column 9: 2-3 hours
    temp(i,9) = length(data.bouts.(day).StepTimeBouts.btwn2_3hr);
    % Column 10: 3-4 hours
    temp(i,10) = length(data.bouts.(day).StepTimeBouts.btwn3_4hr);
    % Column 11: 4+ hours
    temp(i,11) = length(data.bouts.(day).StepTimeBouts.grtr4hr);
    
    % Unused data
    % < 30 min
    %temp(i,12) = length(data.bouts.(day).StepTimeBouts.less30min);
    % > 30 min
    %temp(i,13) = length(data.bouts.(day).StepTimeBouts.grtr30min);
    
    % Clear temporary variable
    clear day
end

% Clear indexing variable
clear i

% Loop over number of Rows
for i=1:11
    
    % Mean for each column
    temp(org.displayDays+1,i)=mean(temp(1:org.displayDays,i));
    % St. Dev. for each colunm
    temp(org.displayDays+2,i)=std(temp(1:org.displayDays,i));
    
end

% Clear indexing variable
clear i

tempRows = [org.rowNames "Mean" "St. Dev."];

% Format Stepping bout count table
Output.StepBoutCountOutput = table(temp(:,1),temp(:,2),temp(:,3),...
    temp(:,4),temp(:,5),temp(:,6),temp(:,7),temp(:,8),temp(:,9),...
    temp(:,10),temp(:,11),'VariableNames',{'< 10 mins','< 15 mins',...
    '10-20 mins','20-30 mins','30-40 mins','40-50 mins','50-60 mins',...
    '1-2 hours','2-3 hours','3-4 hours','> 4 hours'},'RowNames',tempRows);

% Clear temporary variable
clear temp





% Standing Bout Count

% Initialize temporary variable
temp = zeros(org.displayDays+2,11);

% Loop over the number of days
for i=1:org.displayDays
    
    % Load a temporary variable for the name of the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Populate the columns with data based on bout size
    % Column 1: < 10 min
    temp(i,1) = length(data.bouts.(day).standingBouts.less10min);
    % Column 2: < 15 min
    temp(i,2) = length(data.bouts.(day).standingBouts.less15min);
    % Column 3: 10-20 min
    temp(i,3) = length(data.bouts.(day).standingBouts.btwn10_20min);
    % Column 4: 20-30 min
    temp(i,4) = length(data.bouts.(day).standingBouts.btwn20_30min);
    % Column 5: 30-40 min
    temp(i,5) = length(data.bouts.(day).standingBouts.btwn30_40min);
    % Column 6: 40-50 min
    temp(i,6) = length(data.bouts.(day).standingBouts.btwn40_50min);
    % Column 7: 50-60 min
    temp(i,7) = length(data.bouts.(day).standingBouts.btwn50_60min);
    % Column 8: 1-2 hours
    temp(i,8) = length(data.bouts.(day).standingBouts.btwn1_2hr);
    % Column 9: 2-3 hours
    temp(i,9) = length(data.bouts.(day).standingBouts.btwn2_3hr);
    % Column 10: 3-4 hours
    temp(i,10) = length(data.bouts.(day).standingBouts.btwn3_4hr);
    % Column 11: 4+ hours
    temp(i,11) = length(data.bouts.(day).standingBouts.grtr4hr);
    
    % Unused data
    % < 30 min
    %temp(i,12) = length(data.bouts.(day).standingBouts.less30min);
    % > 30 min
    %temp(i,13) = length(data.bouts.(day).standingBouts.grtr30min);
    
    % Clear temporary variable
    clear day
end

% Clear indexing variable
clear i

% Loop over number of Rows
for i=1:11
    
    % Mean for each column
    temp(org.displayDays+1,i)=mean(temp(1:org.displayDays,i));
    % St. Dev. for each colunm
    temp(org.displayDays+2,i)=std(temp(1:org.displayDays,i));
    
end

% Clear indexing variable
clear i

tempRows = [org.rowNames "Mean" "St. Dev."];

% Format Standing bout count table
Output.StandBoutCountOutput = table(temp(:,1),temp(:,2),temp(:,3),...
    temp(:,4),temp(:,5),temp(:,6),temp(:,7),temp(:,8),temp(:,9),...
    temp(:,10),temp(:,11),'VariableNames',{'< 10 mins','< 15 mins',...
    '10-20 mins','20-30 mins','30-40 mins','40-50 mins','50-60 mins',...
    '1-2 hours','2-3 hours','3-4 hours','> 4 hours'},'RowNames',tempRows);

% Clear temporary variable
clear temp



% Sedentary output Table

% Initialize temporary variable
temp = zeros(org.displayDays*4,13);

% Loop over days to populate table
for i=1:org.displayDays
    
    % load temporary variable for the day
    day = ['Day_',num2str(i,'%01.0f')];
    
    % Column 1: < 10 minutes
    temp(4*i-3,1) = length(data.bouts.(day).sedentaryBouts.less10min);
    temp(4*i-2,1) = sum(data.bouts.(day).sedentaryBouts.less10min)/60;
    temp(4*i-1,1) = mean(data.bouts.(day).sedentaryBouts.less10min/60);
    temp(4*i,1) = std(data.bouts.(day).sedentaryBouts.less10min/60);
    
    % Column 2: < 15 minutes
    temp(4*i-3,2) = length(data.bouts.(day).sedentaryBouts.less15min);
    temp(4*i-2,2) = sum(data.bouts.(day).sedentaryBouts.less15min)/60;
    temp(4*i-1,2) = mean(data.bouts.(day).sedentaryBouts.less15min/60);
    temp(4*i,2) = std(data.bouts.(day).sedentaryBouts.less15min/60);
    
    % Column 3: Between 10 & 20 minutes
    temp(4*i-3,3) = length(data.bouts.(day).sedentaryBouts.btwn10_20min);
    temp(4*i-2,3) = sum(data.bouts.(day).sedentaryBouts.btwn10_20min)/60;
    temp(4*i-1,3) = mean(data.bouts.(day).sedentaryBouts.btwn10_20min/60);
    temp(4*i,3) = std(data.bouts.(day).sedentaryBouts.btwn10_20min/60);
    
    % Column 4: Between 20 & 30 minutes
    temp(4*i-3,4) = length(data.bouts.(day).sedentaryBouts.btwn20_30min);
    temp(4*i-2,4) = sum(data.bouts.(day).sedentaryBouts.btwn20_30min)/60;
    temp(4*i-1,4) = mean(data.bouts.(day).sedentaryBouts.btwn20_30min/60);
    temp(4*i,4) = std(data.bouts.(day).sedentaryBouts.btwn20_30min/60);
    
    % Column 5: Between 30 & 40 minutes
    temp(4*i-3,5) = length(data.bouts.(day).sedentaryBouts.btwn30_40min);
    temp(4*i-2,5) = sum(data.bouts.(day).sedentaryBouts.btwn30_40min)/60;
    temp(4*i-1,5) = mean(data.bouts.(day).sedentaryBouts.btwn30_40min/60);
    temp(4*i,5) = std(data.bouts.(day).sedentaryBouts.btwn30_40min/60);
    
    % Column 6: Between 40 & 50 minutes
    temp(4*i-3,6) = length(data.bouts.(day).sedentaryBouts.btwn40_50min);
    temp(4*i-2,6) = sum(data.bouts.(day).sedentaryBouts.btwn40_50min)/60;
    temp(4*i-1,6) = mean(data.bouts.(day).sedentaryBouts.btwn40_50min/60);
    temp(4*i,6) = std(data.bouts.(day).sedentaryBouts.btwn40_50min/60);
    
    % Column 7: Between 50 & 60 minutes
    temp(4*i-3,7) = length(data.bouts.(day).sedentaryBouts.btwn50_60min);
    temp(4*i-2,7) = sum(data.bouts.(day).sedentaryBouts.btwn50_60min)/60;
    temp(4*i-1,7) = mean(data.bouts.(day).sedentaryBouts.btwn50_60min/60);
    temp(4*i,7) = std(data.bouts.(day).sedentaryBouts.btwn50_60min/60);
    
    % Column 8: Between 1 & 2 hours
    temp(4*i-3,8) = length(data.bouts.(day).sedentaryBouts.btwn1_2hr);
    temp(4*i-2,8) = sum(data.bouts.(day).sedentaryBouts.btwn1_2hr)/60;
    temp(4*i-1,8) = mean(data.bouts.(day).sedentaryBouts.btwn1_2hr/60);
    temp(4*i,8) = std(data.bouts.(day).sedentaryBouts.btwn1_2hr/60);
    
    % Column 9: Between 2 & 3 hours
    temp(4*i-3,9) = length(data.bouts.(day).sedentaryBouts.btwn2_3hr);
    temp(4*i-2,9) = sum(data.bouts.(day).sedentaryBouts.btwn2_3hr)/60;
    temp(4*i-1,9) = mean(data.bouts.(day).sedentaryBouts.btwn2_3hr/60);
    temp(4*i,9) = std(data.bouts.(day).sedentaryBouts.btwn2_3hr/60);
    
    % Column 10: Between 3 & 4 hours
    temp(4*i-3,10) = length(data.bouts.(day).sedentaryBouts.btwn3_4hr);
    temp(4*i-2,10) = sum(data.bouts.(day).sedentaryBouts.btwn3_4hr)/60;
    temp(4*i-1,10) = mean(data.bouts.(day).sedentaryBouts.btwn3_4hr/60);
    temp(4*i,10) = std(data.bouts.(day).sedentaryBouts.btwn3_4hr/60);
    
    % Column 11: Greater than 4 hours
    temp(4*i-3,11) = length(data.bouts.(day).sedentaryBouts.grtr4hr);
    temp(4*i-2,11) = sum(data.bouts.(day).sedentaryBouts.grtr4hr)/60;
    temp(4*i-1,11) = mean(data.bouts.(day).sedentaryBouts.grtr4hr/60);
    temp(4*i,11) = std(data.bouts.(day).sedentaryBouts.grtr4hr/60);
    
    % Column 12: Less than 30 minutes
    temp(4*i-3,12) = length(data.bouts.(day).sedentaryBouts.less30min);
    temp(4*i-2,12) = sum(data.bouts.(day).sedentaryBouts.less30min)/60;
    temp(4*i-1,12) = mean(data.bouts.(day).sedentaryBouts.less30min/60);
    temp(4*i,12) = std(data.bouts.(day).sedentaryBouts.less30min/60);
    
    % Column 13: Greater than 30 minutes
    temp(4*i-3,13) = length(data.bouts.(day).sedentaryBouts.grtr30min);
    temp(4*i-2,13) = sum(data.bouts.(day).sedentaryBouts.grtr30min)/60;
    temp(4*i-1,13) = mean(data.bouts.(day).sedentaryBouts.grtr30min/60);
    temp(4*i,13) = std(data.bouts.(day).sedentaryBouts.grtr30min/60);
    
    % Clear temporary variable
    clear day
end
% Clear indexing variable
clear i

tempDayRows = [];
for i=1:org.displayDays
    tempDayRows = [tempDayRows strcat(org.rowNames(i),' events (#)')...
        strcat(org.rowNames(i),' Total time (min)')...
        strcat(org.rowNames(i),' Mean event time (min)')...
        strcat(org.rowNames(i),' SD event time (min)')];
end

% Format Sedentary Output table
Output.SedentaryOutput = table(temp(:,1),temp(:,2),temp(:,3),temp(:,4),...
    temp(:,5),temp(:,6),temp(:,7),temp(:,8),temp(:,9),temp(:,10),...
    temp(:,11),temp(:,12),temp(:,13),'VariableNames',{'< 10 mins',...
    '< 15 mins','10-20 mins','20-30 mins','30-40 mins','40-50 mins',...
    '50-60 mins','1-2 hours','2-3 hours','3-4 hours','> 4 hours',...
    '< 30 mins','> 30 mins'},'RowNames',tempDayRows);
clear temp*

%temp = zeros(org.numDays,2);

for i=1:org.displayDays
    
    day = ['Day_',num2str(i,'%01.0f')];
    
    temp(i,1) = raw.epochDateData(data.Times.DayStart.trimmed.(day));
    temp(i,2) = raw.epochDateData(data.Times.DayEnd.trimmed.(day));
    
    clear day
end
clear i


Output.dayTimes = table(temp(:,1),temp(:,2),'VariableNames',...
    {'Day Start','Day End'},'RowNames',org.rowNames);
clear temp

%% Export data

% Request file save location and name from user, default based on original
% file
[file,path] = uiputfile({'*.xlsx'},'Save output as:',...
    strcat(erase(org.epochFileName,' by 15s epoch.csv'),'.xlsx'));


% Write each output table as a sheet in the indicated file
writetable(Output.OverallOutput,strcat(path,file),...
    'Sheet','Summary Output','WriteRowNames',true);
writetable(Output.SummaryOutput,strcat(path,file),...
    'Sheet','Summary Output','Range','A5','WriteRowNames',true);
writetable(Output.dayTimes,strcat(path,file),...
    'Sheet','Day times','WriteRowNames',true);
writetable(Output.SedentaryOutput,strcat(path,file),...
    'Sheet','Sedentary Output','WriteRowNames',true);
writetable(Output.SedBoutCountOutput,strcat(path,file),...
    'Sheet','Sedentary Bout Count','WriteRowNames',true);
writetable(Output.StandBoutCountOutput,strcat(path,file),...
    'Sheet','Standing Bout Count','WriteRowNames',true);
writetable(Output.StepBoutCountOutput,strcat(path,file),...
    'Sheet','Stepping Bout Count','WriteRowNames',true);

clear file path choice
