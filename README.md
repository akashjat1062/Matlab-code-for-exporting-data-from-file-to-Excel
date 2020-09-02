# Matlab-code-for-exporting-data-from-file-to-Excel
We were given an assignment on Analyzing strong motion data, which includes downloading data from KNet seismograph network, data from KNET network come as a single zipped file which containes zipped files for all the recording sites, which in turn contains the files for 3 components NS, UD, EW.(North-South, Up-Down, East-West).
The format of these files were EW, ND, UD, which cannot be open by any windows program.

So for this project I had to extract all th files from a zipped folder and place them in a single folder, which is a relatively simple task on windows.
Now I had to take any one component, So i choose UD, and converted all the UD files to .csv format which can be done by using a single line command " ren *.UD *.csv " on windows command prompt. 
 Now from these csv files I had to extract some specific information, this information is located at the same place in every .csv files 
 The information I needed was latitude and longitude of epicentre, depth of earthquake, magnitude of earthquake, lat and long of station, and duration  for envelope functions.
 numerically i required 7 things to extract from the data.
 
 So what I did---
 I made a script file on MATLAB.
 the coding starts from here ....
 files = dir('*.csv');   # files is a structure containing information about all the files with .csv format in a given folder, suppose if some folder contains 6 .csv files then it # will be of 6x1 struct. it contains attributes like name, size and others attributes..
 files = files'
 nos = length(files)   # nos represents the number of stations which is equal to number of .csv files
 C = zeros(nos,7);     # the output cell to store the information
 E = zeros(nos,1);     # the output vector which will store the epicentral distance of each recording station using haversine function.
 H = zeros(nos,1);     # the output vector which contains the hpocentral sitance obtained by ((E)2+(Depth)2)^(0.5).
 for i = 1:nos
     filename = file(i).name   # taking the file name of each station and stroing it in filename
     [~,~,raw]= xlsread(filename);  # this is very important step, it saves the entire excel file in the form of a matrix , and the specific cells of an excel file can be accessed 
     # via mentioning them for example B2 can be accessed by raw{2,2}
    # Now the information i wanted was all stored in first column of the excel file but mingled with texts and symbols , therefore I used regex expressions to extract and store #them.
    C(i,:)= [double(regexp(string(raw{2,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{3,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{4,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{5,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{7,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{8,1}),'(\d+\.?+\d+)','match')),double(regexp(string(raw{12,1}),'(\d+\.?+\d+)','match'))]; # I know it looks a bit bulky expression, but all i did was design a first expression and copied it 7
    # 7 times to extract the specific information , if you have any idea about how this expression can be reduced, please feel free to suggest me.
    E(i,1)= haversine([C(i,1),C(i,2)],[C(i,5),C(i,6)]); #  to calculate the distance using the lat long difference of epicentre and station
    H(i,1)=(E(i,1)^2+C(i,4)^2)^(0.5);
     end
     
     hold on                         # plotting for further regression analysis 
X=C(:,7);
X=X-((0.0015)*(10)^(0.5*4.7));
X=log10(X);
Y=log10(H);
I=find(X>1.7700 & X<1.7800);
for j = I
    X(j)=[];
    Y(j)=[];
end

plot(Y,X,'ko')

# So overall this script takes all the .csv files in a folder, extract specific values from them, store them in a matrix, calculate epicentral and hypocentral distance and plot the graph between X and Y  defined by midorikawa's empirical relationship between duration and hypocentral distance.
 
