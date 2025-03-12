function varargout = finalproject(varargin)
clc
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @finalproject_OpeningFcn, ...
                   'gui_OutputFcn',  @finalproject_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end


function finalproject_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

guidata(hObject, handles);

function varargout = finalproject_OutputFcn(hObject, eventdata, handles) 
varargout{1} = handles.output;
axes(handles.axes1);
imshow('C:\Users\Admin\Desktop\myproject\OMR-Sheet-Evaluation-using-Matlab\OMR IMAGE\restartimage.png');


function load_image_Callback(hObject, eventdata, handles)

[path,user_cance]=imgetfile();
if user_cance
    return
end
im=rgb2gray(imread(path));
siz = size(im);

if siz(1)==0 || siz(2)==0 || siz(1)<3000 || siz(2)<2400
    msgbox(sprintf('Error!! You have Selected Wrong image!'),'Error','Error');
    return
end
ok = is_img_ok(im);
if ~ok
    msgbox(sprintf('Error!This OMR can not be evaluated! Please Restart!'),'Error','Error');
    return;
end

handles.im = im;

axes = handles.axes1;
imshow(handles.im);
guidata(hObject,handles);
function qno_Callback(hObject, eventdata, handles)

function qno_CreateFcn(hObject, eventdata, handles)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function evaluation_Callback(hObject, eventdata, handles)

qu_no = str2num(get(handles.qno,'String'));
[marks,Remark,Roll,Test] = finalver7(handles.im,handles.solution,qu_no);
handles.marks = marks;
set(handles.mark,'String',num2str(marks))
handles.Roll = Roll;
set(handles.remarks,'String',Remark)
handles.Remark= Remark;
set(handles.roll,'String',num2str(Roll))
handles.Test = Test;
set(handles.testid,'String',num2str(Test))
guidata(hObject,handles);


function about_Callback(hObject, eventdata, handles)
h = msgbox ({'Welcome to our OMR Sheet Evaluation Toolbox','','This is a mini project done under the guidance of prof.SAI JYOTHI ','','','', ...
    'We are:','','INDUMATHI','','BHAVISHYA','','BHAGYA SRI','','GOWTHAMI','','MOKSHITHA','','', ...
    'Department of Electronics and Communication Engineering','','ECE-3','','E3-SEM 1','','RGUKT-ONGOLE','','Thank You!', ...
    ''},'About');


function Help_Callback(hObject, eventdata, handles)

h = msgbox({'Here is our Help Window','Please read to solve your problem or to know what you want!','','Load OMR','Click the Load Image button', ...
	'a file open dialog box will open','Go to OMR Image folder in which the scanned image of OMR sheets are stored', ... 
    'Select an image and click open.' ...
    '','','In the input box give the number of question you want to evaluate(maximum 60)','','Load Solution','','Click the Load Image button', ...
	'a file open dialog box will open.',' Go to OMR Solution folder in which the solution excel file of OMR sheets are stored.', ...
    'Select a file and click open.','','IF YOU CAN NOT SEE ANY FILE,THEN CHANGE THE FILE TYPE TO *all files ','AND SELECT YOUR SOLUTION FILE.', ...
    'NOTE THAT THIS FILES ARE EXCEL FILES.,','','click Evaluate OMR button and wait to see roll,test id,marks,remarks in the specified field.', ...
    'The Roll, marks,Test ID and Remarks are shown In the Specified Field', ...
    '','','Click restart if any problem arises or you want to evaluate freshly.','','', ...
    'Note: After first omr sheet evaluation', 'you can change the omr or the solution according to your wish until you press restart button.', ...
    'Hope, Your query is Satisfied!'},'HELP');


function load_solution_Callback(hObject, eventdata, handles)

[path,user_cance]=imgetfile();
if user_cance
    return
end
[~,~,sol] = xlsread(path,'A1:A60');
solution = upper(char (sol));
handles.solution = solution;
guidata(hObject,handles);



function restart_Callback(hObject, eventdata, handles)

axes(handles.axes1);
imshow('C:\Users\Admin\Desktop\myproject\OMR-Sheet-Evaluation-using-Matlab\OMR IMAGE\restartimage.png');

set(handles.qno,'String','');
set(handles.mark,'String','')
set(handles.remarks,'String','')
set(handles.roll,'String','')
set(handles.testid,'String','')
guidata(hObject,handles);

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
    
  

function [mark,remarks,Roll,TD]=finalver7(I,solution,num)
   
    
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



function okay =is_img_ok(I)
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
     if blackblock ==6
        okay =1;
     else
        okay =0;
     end
   
    

