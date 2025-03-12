
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
h = msgbox ({'Welcome to our OMR Sheet Evaluation Toolbox','','This is a project for Numerical Technique Laboratory','','Course No. EEE 212','', ...
    'We are:','','Md.Mehedi Hossen Limon','','Fariba Islam','','', ...
    'Department of Electrical & Electronic Engineering','','Bangladesh University of Engineering & Technology','','Thank You!', ...
    '','For Any Query or Advice','Feel free to email at','mhlimon1995@gmail.com','ezqhaz@gmail.com'},'About');


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
