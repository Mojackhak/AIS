% user guider
% set the pixel(default=294.6278/2048)
% set the zstep(default=1)
% set the the angle between two rectangular coordinate systems 
% set the the original coordinate system
% set the the z coordinate 
% !!!must input the excel produced by imagej macro "AIS_.ijm"
%this is the simplified version of aisprocess which will only outputs 3D
%information drawed by hand.

function [] = ais()
%     Reads the specific suffix filename at the specified path
xlsfile = dir('*');
aisfile = regexpi({xlsfile.name},'.*AIS.xlsx$','match');
aisfile = [aisfile{:}];
buttonfile = regexpi({xlsfile.name},'.*button.xlsx$','match');
buttonfile = [buttonfile{:}];
for i = 1: length(aisfile)
    getresults(aisfile{i},buttonfile{i});
end
end


function [] = getresults(aisfile,buttonfile)
[~, Sheets]=xlsfinfo(aisfile);
% set the threshold and other parameters of the start and end points of ais
pixel = 294.6278/2048;
zstep = 0.5;
d = round(2.5/pixel); %%%sets on of pixels each side, 2 * d * pixel ≈ 5    i.e. for d = 20, width of sliding window is 41
%bg = round(2.6/pixel);
threshold = 0.33; %%[threshold:protein ais(AnkG)]

for sheeti = 1:length(Sheets)
    data_num = xlsread(aisfile, Sheets{sheeti});
    ais = data_num(1:end/2,9:end);
    morphology = data_num(1:end/2,7:8); 
    [button, txt]= xlsread(buttonfile, Sheets{sheeti});
    txt = txt(2,1:end-3);
    %% produce the image
    Chais = mat2gray(ais);
    f=figure(1);
    set(gcf,'position',get(0,'ScreenSize'));               %max the window
    imagesc(Chais)                                         %Plots selected ais figure.
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
    pause(0.1);%in case of right click on matlab
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
    %         ais intensity path
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
    ais_max_index = find(ais_normal == max(ais_normal));
    ais_max_length = distance(ais_max_index);
    %% get the end of ais
    %%%% point index along axon past max where fluorescence intensity falls to f of its peak
    ais_end_index = find( (distance> ais_max_length) & (ais_normal < threshold) );
    if length(ais_end_index)>0
        Y_end = distance(ais_end_index(1));
    else
        Y_end = distance(end);
    end
    %% get the start of protein/ais
    %%%% point index along axon pre max where fluorescence intensity falls to f of its peak
    ais_start_index = find( (distance< ais_max_length) & (ais_normal < threshold) );
    if length(ais_start_index)>0
        Y_start = distance(ais_start_index(end));
    else
        Y_start = distance(1);
    end    
    %% get ais length 
    Y_length = Y_end - Y_start;
    
    %% get ais length mean intensity
    %%get ais part
    end_distance = find(distance == Y_end);
    start_distance = find(distance == Y_start);
    distance_ais = distance(start_distance:end_distance);
    data_ais = Y_profile(start_distance:end_distance);    
    mean_intensity =trapz(distance_ais,data_ais,2)./Y_length;   
    %% get morphology from soma to ais terminal
    morphology3dx = x_ax';
    morphology3dy = y_ax';
    morphology2d = zeros(length(morphology3dx),2);
    for m = 1:length(morphology2d)
        morphology2d(m,:) = morphology(y_pix(m),:);
    end
    
   %% get button location
    if sum(button(:,3)) ~= 0
        button_info = button(:,1:end-3);
        y_location = button_info(:,1);
        for p = 1:length(button_info(:,1))
            dis_cali = data_num(1:end/2,6); 
%             find(dis_cali==y_location(p))
            y_location(p) = find(dis_cali==y_location(p));
        end
        location3d = zeros(length(y_location),1);
        type = zeros(length(y_location),1);
        for i = 1:length(y_location)
            point_dis = [];
            for h = 1:length(xysm)
                point_dis = [point_dis;(y_location(i)-xysm(2,h))^2];
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
    length_info = [Y_start, Y_end, Y_length, mean_intensity, button(1,end-2:end)];
    morphology_info = [morphology2d morphology3dx morphology3dy];
    profile_info = [X_profile; Y_profile; X_profile; Y_smooth]';
    point_info = [start_distance(1), end_distance(1)];
    
    
    %% output
    title = [{'start 647'} {'end 647'} {'length 647'} {'Mean intensity of 647'} ...
         {'width'} {'x'}  {'y'}];
    title_profile =[{'ais_distance'} {'647 profile intensity'} {'ais_distance'} {'647 profile smooth'} {'647 start'} {'647 end'}];
    title_morphology =[{'x 2d'} {'y 2d'} {'x 3d'} {'y 3d'} {'647 start'} {'647 end'}];

    
    xlswrite([aisfile '_length.xlsx'],title,Sheets{sheeti},'A1');
    xlswrite([aisfile '_length.xlsx'],length_info,Sheets{sheeti},'A2');
    xlswrite([aisfile '_morphology.xlsx'],title_morphology,Sheets{sheeti},'A1');
    xlswrite([aisfile '_morphology.xlsx'],morphology_info,Sheets{sheeti},'A2');
    xlswrite([aisfile '_morphology.xlsx'],point_info,Sheets{sheeti},'E2');
    xlswrite([aisfile '_Profile.xlsx'],title_profile,Sheets{sheeti},'A1');
    xlswrite([aisfile '_Profile.xlsx'],profile_info,Sheets{sheeti},'A2');
    xlswrite([aisfile '_Profile.xlsx'],point_info,Sheets{sheeti},'E2');
    xlswrite([aisfile '_Profile.xlsx'],point_info,Sheets{sheeti},'E2');
    
    if sum(button(:,3)) ~= 0
        title_button = [txt, {'location3d'}, {'is on ais'}, {'relative location'}];
        button_info_sum = [button_info, location3d, type, reloca];
        xlswrite([aisfile '_length.xlsx'],title_button,Sheets{sheeti},'A4');
        xlswrite([aisfile '_length.xlsx'],button_info_sum,Sheets{sheeti},'A5');
    end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%    
    
end
end
