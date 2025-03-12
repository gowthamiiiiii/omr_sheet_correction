function [mark,remarks,Roll,TD]=finalver7(I,solution,num)
    % Function: finalver7
    % Written By-
    %   Md. Mehedi Hossen Limon
    %   Fariba Islam
    %   Department of Electrical & Electronic Engineering
    %   Bangladesh University of Engineering and Technology
    %
    % Inputs :
    %   I = image file
    %   solution = a xlsx solution file
    %   num = number of question
    % Outputs:
    %   mark = Total Correct answer
    %   remarks = Comment on the Result
    %   Roll = Roll of the examinee
    %   TD = Test ID
    
    diffcol = 57;
    diffrow = 116;
    diffsection = 462;
    solution = solution(1:num);
    if length(solution)==0
        msgbox(sprintf('Error!! Please Restart!'),'Error','Error');
        return
    end
    p = .6;
    str = 'ABCD';

    c  = zeros(num,5);
    c(:,1) = 1:num;
    mark = 0;
    R = zeros(10);
    Roll = 0;
    T = zeros(10,3);
    TD = 0;

    % for MCQ
    for k = 1:num

        count = (k<=15) + 2*(k>=16&&k<=30)+3*(k>=31&&k<=45)+ 4*(k>45);
        startx = 487; starty = 1488;
        startx = startx+(count-1)*diffsection;
        sy = starty + ((k-((count-1)*15)-1))*diffrow;
        for l = 1:4
            sx = startx + (l-1)*diffcol;
            [n,total] = blackpixelcounter(I,sx,sy);
            if n>= p*total
                c(k,l+1) = 0;
                Result= str(l);
                % check right answer;
                if Result== solution(k)
                    mark = mark + 1;
                    
                end
                
            else
                c(k,l+1) = 1;
            end
                            % check double ;
            if l == 4
                dob = 0;markdouble = 0;
                for a = 1:l
                    if c(k,a+1) == 0
                        dob = dob+1;
                        if Result == solution(k)
                            markdouble = markdouble +1;
                        end
                    end
                end
                if dob>1 && markdouble ~=0
                    mark = mark -1;
                    disp('double')
                end
            end

        end
    end

    % for roll
    for k = 1:10

        startx = 256;starty = 737;
        diffrow = 57;
        sy = starty +(k-1)*diffrow;
        for l = 1:10
            sx = startx + (l-1)*diffcol;
            [n,total] = blackpixelcounter(I,sx,sy);
            if n>= p*total
                R(k,l) = 0;
                if k ~=10
                    Roll = Roll + k*10^(9-(l-1));
                end
            else
                R(k,l) = 1;
            end
        end
    end
    
    % for test id
    for k = 1:10

        startx = 949;
        starty = 737;diffrow = 57;
        sy = starty +(k-1)*diffrow;

        for l = 1:3
            sx = startx + (l-1)*diffcol;
            [n,total] = blackpixelcounter(I,sx,sy);
            if n>= p*total
                if k ~= 10
                    TD = TD + k*10^(2-(l-1));
                end
            else
                T(k,l) = 1;
            end
        end
    end
    if(mark>=floor(.33*num))
        status='Great!   Passed!';
    else
        status='Sorry!  Failed! Need Improvement';
    end
    remarks = status;
    
end
