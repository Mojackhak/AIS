%%%%  AIS localization and intensity profile from z-series images
% user guider
% set the pixel(default=0.1439)
% set the zstep(default=1)
% set the threshold(default=[0.33,0.33])
% !!!must input the excel produced by imagej macro "ais_line.ijm"

function [] = ais_line()
%     Reads the specific suffix filename at the specified path
xlsfile = dir('*');
xlsfile = regexpi({xlsfile.name},'.*(.xlsx)$','match');
xlsfile = [xlsfile{:}];
for i = 1: length(xlsfile)
    getresults(xlsfile{i});
end
end


function getresults(filename)
[~, Sheets]=xlsfinfo(filename);
% set the threshold and other parameters of the start and end points of ais
pixel = 0.1439;
zstep = 1;
threshold = [0.33,0.33]; %%[threshold: protein, AIS marker (AnkG)]
%%%sets on of pixels each side, 2 * d * pixel ≈ 5    i.e. for d = 20, width of sliding window is 41
d = round(2.5/pixel); 
for sheeti = 1:length(Sheets)
    [data_num,txt] = xlsread(filename, Sheets{sheeti});
    txt_title = txt(2,:);
    for i = 1:length(txt_title)  
        if isequal(txt_title{i},'width')% find the nearest_point column
            width = i;
            break
        end
    end
    ais = data_num(1:end/2,1:width-1);
    protein = data_num(end/2+1:end,1:width-1);
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
Y_profile = zeros(2,length(saxon_um));
for p = 1:length(saxon_um)
    %         ais intensity path
    Y_profile(2,p) = ais(y_pix(p),x_pix(p));
    %         protein intensity  path
    Y_profile(1,p) = protein(y_pix(p),x_pix(p));
end

 %% Identify the start and end points of protein/ais by threshold

    %% smooth the curve
    Y_smooth = Y_profile;
    distance = X_profile;
    Y_length = [0,0];
    Y_start = [0,0];
    Y_end = [0,0];
    for y = 1:length(Y_smooth)
        if y<(d+1)  %%% at very start of axon profile , not allowing full window
            Y_smooth(:,y) = mean(Y_profile(:,1:y+d),2); %% sliding mean with lv_smooth values
        elseif y>(length(distance)-(d+1))  %%% at very end of axon profile, not allowing full window
            Y_smooth(:,y) = mean(Y_profile(:,y-d:end),2);
        else   %%% in middle of axon profile, allowing full window
            Y_smooth(:,y) = mean(Y_profile(:,y-d:y+d),2);
        end
    end
    Y_normal = (Y_smooth - min(Y_smooth,[],2)) ./ (max(Y_smooth,[],2)-min(Y_smooth,[],2));
    
    for ch = 1:2
        %%  find the max intensity point of protein/ais
        ais_normal = Y_normal(ch,:);
        [~,ais_max_index] =  max(ais_normal);
        ais_max_length = distance(ais_max_index);
        %% get the end of protein/ais
        %%%% point index along axon past max where fluorescence intensity falls to f of its peak
        ais_end_index = find( (distance> ais_max_length) & (ais_normal < threshold(ch)) );
        if ~isempty(ais_end_index)
            Y_end(ch) = distance(ais_end_index(1));
        else
            Y_end(ch) = distance(end);
        end
        %% get the start of protein/ais
        %%%% point index along axon pre max where fluorescence intensity falls to f of its peak
        ais_start_index = find( (distance< ais_max_length) & (ais_normal < threshold(ch)) );
        if ~isempty(ais_start_index)
            Y_start(ch) = distance(ais_start_index(end));
        else
            Y_start(ch) = distance(1);
        end
        %% get protein/ais length
        Y_length(ch) = Y_end(ch) - Y_start(ch);
    end
    
    %% get protein/ais length mean intensity
    end_distance = find(distance == Y_end(2),1);
    start_distance = find(distance == Y_start(2),1);
    distance_ais = distance(start_distance:end_distance);
    data_ais = Y_profile(:,start_distance:end_distance);    
    mean_intensity =trapz(distance_ais,data_ais,2)./Y_length(2);   
    %% aggregate information.
    length_info = [Y_start, Y_end, Y_length, mean_intensity'];
    profile_info = [X_profile; Y_profile;Y_smooth]';

    %% output
    title = [{'start position of protein'} {'start position of AIS'} ... 
             {'end position of protein'} {'end position of AIS'}  ... 
             {'length of protein'} {'length of AnkG'} ... 
             {'Mean intensity of protein'} {'Mean intensity of AnkG'}];
    title_profile =[{'distance to soma'} {'profile of protein intensity'} {'profile of AnkG intensity'} ... 
                    {'smoothed profile of protein intensity'} {'smoothed profile of AnkG intensity'}];
    
    xlswrite([filename '_AIS.xlsx'],title,Sheets{sheeti},'A1');
    xlswrite([filename '_AIS.xlsx'],length_info,Sheets{sheeti},'A2');
    xlswrite([filename '_Profile.xlsx'],title_profile,Sheets{sheeti},'A1');
    xlswrite([filename '_Profile.xlsx'],profile_info,Sheets{sheeti},'A2');


end
end
