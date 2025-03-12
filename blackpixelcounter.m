function [ black,total] = blackpixelcounter(I,startx,starty)
    % Funtion: Blackpixelcounter
    % This funtion returns number of black pixel and total pixel of a
    % circle of the omr
    % Inputs:
    %   I = gray image file of omr
    %   startx = x ordiante value of center of the circle
    %   starty = y coordinate value of center of the circle
    % Outputs:
    %   black = number of black pixel in the circle
    %   total = number of total pixel in the circle
    radius = 21;
    diameter = 43;
    sx = startx;
    sxnew = sx - radius;
    newwidy = 12;
    totpix = 0;
    n = 0;
    t=200;
  
    for i = 1:diameter

        synew = starty - newwidy/2;
        for j =1:newwidy
            a = I(synew+j-1,sxnew+i-1);
            if a<=t
                n = n + 1;
            end
            totpix = totpix + 1;
        end
        if i<=4
            newwidy = newwidy + 4;
        elseif (i>4 && i<=11) || i == 15 
            newwidy = newwidy + 2;
        elseif (i>32 && i<40) || i==27 || i==29
            newwidy = newwidy - 2;
        elseif i>=40
            newwidy = newwidy - 4;
        end
    end
    black = n;
    total = totpix;
    
end

