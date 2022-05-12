%%%%  AIS localization,intensity profile and button information from z-series images
% user guider
% set the byeye==1 if you need to mearsure by eye(default=1)
% set the pixel(default = 0.1439)
% set the zstep(default = 0.5)
% set the threshold(default = 0.33)
% !!!must input the excel produced by imagej macro "ais_point.ijm"

function [] = ais_point()
%     Reads the specific suffix filename at the specified path
xlsfile = dir('*');
xlsfile = regexpi({xlsfile.name},'.*(.xlsx)$','match');
xlsfile = [xlsfile{:}];
for i = 1: length(xlsfile)
    getresults(xlsfile{i});
end
end


function [] = getresults(filename)
[~, Sheets]=xlsfinfo(filename);
% set the threshold and other parameters of the start and end points of ais
pixel = 0.1439;
zstep = 0.5;
threshold = 0.33; %%[threshold: AIS marker (AnkG)]
 %%%sets on of pixels each side, 2 * d * pixel ≈ 5    i.e. for d = 20, width of sliding window is 41
d = round(2.5/pixel);
for sheeti = 1:length(Sheets)
    [data_num,txt] = xlsread(filename, Sheets{sheeti});
    button = 0; % if there are no synapses on AIS
    txt_title = txt(2,:);
    for i = 1:length(txt_title)  
        if isequal(txt_title{i},'nearest_point')% find the nearest_point column
            button = i;
            break
        end
    end
    ais = data_num(1:end,1:button);
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %% Plots AnkG image.
    Chais = mat2gray(ais);
    f = figure(1);
    imagesc(Chais)                                         
    colormap(gray)
    daspect([1 zstep/pixel 1])
    hold on
    
    %% selet the ais middle line
    xy = [];
    n = 0;  %%%% Loop, picking up the points.
    points = [];
    while 1
        [xi,yi,but] = ginput(1);	%%%% getting co-ordinates from each mouse click
        n = n+1;
        xy(:,n) = [xi;yi];
        if but == 1
            points(n) = plot(xi,yi,'r*');
        end
        if but ~= 1
            if but == 3
                points(n) = plot(xi,yi,'r*');
                break      %%%% right click to finish
            else
                delete(points(n-1));
                xy(:,n-1:n) = []; %%%% press other key to repick the points
                n = n-2;
            end
        end
    end
    
    % make sure the start point on the first line
    xy(2,1) = 1 ;
    
    %%%% interpolate points with a spline curve and finer spacing.
    t = 1:n;
    ts = 1: 0.01: n;    %%%% fine sampling, with double deletion below, ensures pixel-by-pixel line section
    xysm = makima(t,xy,ts);
    for u = 2:(length(xysm))
        ax(u) = sqrt ( ((xysm(1,u) - xysm(1,u-1)) * zstep)^2 + ((xysm(2,u)-xysm(2,u-1))* pixel)^2 );
    end
    ax = cumsum(ax);    %%%% so distance (in pixels) along spline from start of axon
    ax_um = ax;   %%% distance along spline in um
    xys = round(xysm); %%% rounding spline to whole-number pixel co-ordinates only
    plot(xys(1,:),xys(2,:),'-b')
    plot(xysm(1,:),xysm(2,:),'-y')
    hold off
    pause(0.1);%wait for process
    close(f)
    
    %% get the ais path unique coordinate
    double_id = [];
    for i = 1:length(xys(1,:))
        for j = (i+1):length(xys(1,:))
            if xys(1,i) == xys(1,j)
                if xys(2,i) == xys(2,j)
                    double_id = [double_id; j]; %%%% so finding double co-ordinates
                end
            end
        end
    end
    double_id = unique(double_id);
    alln = 1:length(xys(1,:));
    m = setdiff(alln,double_id);    %%% so m is index of all unique xy points in line section
    xs = xys(1,:); ys = xys(2,:);
    x_pix = xs(m);	%%%% so x_pix is unique array of axon x co-ordinates
    y_pix = ys(m);	%%%% y_pix is unique array of axon y co-ordinates
    x_ax = zeros(1,length(x_pix));
    y_ax = zeros(1,length(y_pix));
    
    
    %% get axon path length
    mindi = zeros(1,length(x_pix));
    for g = 1:length(x_pix) %%%% for each pixel, finding nearest location on spline axon
        dis{g} = [];
        for h = 1:length(xysm)
            dis{g} = [dis{g}; sqrt(((x_pix(g)-xysm(1,h)) * zstep)^2+((y_pix(g)-xysm(2,h))* pixel)^2)];
        end
        [~,mindi(g)] = min(dis{g});
        %%%% for each pixel, finding nearest coordinate on spline axon
        x_ax(g) = xysm(1,mindi(g));
        y_ax(g) = xysm(2,mindi(g));
    end
    saxon_um = ax_um(mindi);   %%% so is array of distances of each pixel along axon - more accurate version of axon_um
    saxon_um = saxon_um - saxon_um(1); %%% start with 0;
    
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    

%% get axon profile
X_profile = saxon_um;
Y_profile = zeros(1,length(saxon_um));
for p = 1:length(saxon_um)
    % ais intensity path
    Y_profile(1,p) = ais(y_pix(p),x_pix(p));
end

 %% Identify the start and end points of ais by threshold

    %% smooth the curve
    Y_smooth = Y_profile;
    distance = X_profile;
    for y = 1:length(Y_smooth)
        if y<(d+1)  %%% at very start of axon profile , not allowing full window
            Y_smooth(y) = mean(Y_profile(1:y+d)); %% sliding mean with lv_smooth values
        elseif y>(length(distance)-(d+1))  %%% at very end of axon profile, not allowing full window
            Y_smooth(y) = mean(Y_profile(y-d:end));
        else   %%% in middle of axon profile, allowing full window
            Y_smooth(y) = mean(Y_profile(y-d:y+d));
        end
    end
    Y_normal = (Y_smooth - min(Y_smooth)) ./ (max(Y_smooth)-min(Y_smooth));

    %%  find the max intensity point of ais
    ais_normal = Y_normal;
    [~,ais_max_index] =  max(ais_normal);
    ais_max_length = distance(ais_max_index);
    %% get the end of ais
    %%%% point index along axon past max where fluorescence intensity falls to f of its peak
    ais_end_index = find( (distance> ais_max_length) & (ais_normal < threshold) );
    if ~isempty(ais_end_index)
        Y_end = distance(ais_end_index(1));
    else
        Y_end = distance(end);
    end
    %% get the start of ais
    %%%% point index along axon pre max where fluorescence intensity falls to f of its peak
    ais_start_index = find( (distance< ais_max_length) & (ais_normal < threshold) );
    if ~isempty(ais_start_index)
        Y_start = distance(ais_start_index(end));
    else
        Y_start = distance(1);
    end    
    %% get ais length 
    Y_length = Y_end - Y_start;
    
    %% get ais length mean intensity
    %%get ais part
    end_distance = find(distance == Y_end,1);
    start_distance = find(distance == Y_start,1);
    distance_ais = distance(start_distance:end_distance);
    data_ais = Y_profile(start_distance:end_distance);    
    mean_intensity =trapz(distance_ais,data_ais,2)./Y_length;   
   %% get button location
    if button ~= 0
        button_info = data_num(:,button:button+1);
        button_info = button_info(find(button_info(:,end)),:);%%delete zeros
        x_location = button_info(:,2);
        y_location = button_info(:,1);
        location3d = zeros(length(x_location),1);%%to save the distances of synapses to soma 
        type = zeros(length(x_location),1);%%to save the whether the synapses is on the soma
        for i = 1:length(x_location)
            point_dis = [];
            for h = 1:length(xysm)
                point_dis = [point_dis; sqrt(((x_location(i)-xysm(1,h)) * zstep)^2+((y_location(i)-xysm(2,h))* pixel)^2)];
            end
            [~,nearest] = min(point_dis);
            location3d(i) = ax_um(nearest)-ax_um(mindi(1));  
           % judge whether the button is on the ais
            if (location3d(i) > Y_end) || (location3d(i) < Y_start)
               type(i) = 0;
            else
               type(i) = 1;
            end
        end
        reloca = location3d./Y_end;
    end
    %% aggregate information.
    length_info = [Y_start, Y_end, Y_length, mean_intensity, data_num(1,button+2:end)];
    profile_info = [X_profile; Y_profile; Y_smooth]';   
    
    %% output
    title = [{'start position of AIS'} {'end position of AIS'} {'length of AIS'} {'Mean intensity of AnkG'} ...
         {'width'} {'x'} {'y'}];
    title_profile =[{'distance to soma'} {'profile of AnkG intensity'} {'smoothed profile of AnkG intensity'}];

    
    xlswrite([filename '_AIS.xlsx'],title,Sheets{sheeti},'A1');
    xlswrite([filename '_AIS.xlsx'],length_info,Sheets{sheeti},'A2');
    xlswrite([filename '_Profile.xlsx'],title_profile,Sheets{sheeti},'A1');
    xlswrite([filename '_Profile.xlsx'],profile_info,Sheets{sheeti},'A2');
    
    if button ~= 0
        title_button = [{'location3d'}, {'is on ais'}, {'relative location of synapses'}];
        button_info_sum = [location3d, type, reloca];
        xlswrite([filename '_AIS.xlsx'],title_button,Sheets{sheeti},'A4');
        xlswrite([filename '_AIS.xlsx'],button_info_sum,Sheets{sheeti},'A5');
    end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    
    
end
end
