function oka =is_img_ok(I)
    %   This function evaluates if an image of an omr sheet can be evaluated or not
    % Inputs: 
    %   I = grayscale image of an omr sheet
    % Outputs:
    %   oka = return 1 if can be evaluated , 0 otherwise
    t =200;
    blackblock = 0;

    xwid = 22;
    ywid = 22;
    for l = 1:4
        if l == 1
            startx= 128;
            starty = 261; 
        elseif l==2
            startx = 2324;
            starty = 264;
        elseif l==3
            startx = 130;
            starty = 3214;
        elseif l==4
            startx = 2325;
            starty = 3215;
        end

        synew = starty;
        sxnew = startx;
        n = 0;total =0;
        for i = 1:xwid
            for j = 1:ywid
                a = I(synew+j-1,sxnew+i-1);
                if a<=t
                    n = n + 1;
                end
                total = total + 1;
            end
        end
        if n>=.4*total
            blackblock = blackblock +1;
        end
    end
    if blackblock ==4
        oka =1;
    else
        oka =0;
    end
    
    
end
