% code pour Maria 13/11/2024
% calculates VelocityCorr and MSD


% program startup
clear
% fetch the .xls folder
folder = uigetdir();
extension = '*.xls';
baseFileName = dir (extension) ;
name = {baseFileName.name};
size=length(baseFileName);

% loop to browse xls files
for i=1:size
    % get filename 
    filename = name(i);
    fichier = strcat(folder,filename);
    fichier= string(fichier);
    
    % used to specify the path the file must have the prefix 'DUP'
    fichier = insertBefore(fichier,'DUP','\');
    [pathstr,suj,ext] = fileparts(fichier);

    % transforms the .xls file into a text file
    matrix=readmatrix(fichier,"FileType","text");
    
    % transform in .csv
    newFile = strrep(fichier, '.xls','.csv');
    fid = fopen(newFile,'w');
    % write data into the csv file
    fprintf(fid,'%.9f,%.9f,%.9f,\n',matrix');
    fclose(fid);
    
    % transforms data into MSDanalyzer-readable data
    baseTable = readtable(newFile);
    baseTable2cell = table2cell(baseTable);
  
    % concatenates data to obtain a cell with all data
    stock = baseTable2cell;
    tracks = repmat({cell(360,3)},1,1);
    stock=cell2mat(stock);
    tracks{1,1}= stock;
    
    % initialize msdanalyser
    ma = msdanalyzer(2,'µm','s');
    ma = ma.addAll(tracks);
    
    % creation of the first figure (core tracking according to x and y positions)
    
    f1 = figure(1);
    tracking = ma.plotTracks;
    
    ma = ma.computeMSD;
    display(ma.msd);
    
    MSD=ma.msd;
   
    
    % creation of the second figure (MSD)
    f2 = figure(2);
    plotMSD = ma.plotMSD();
    
    % creation of the third figure (VCorr)
    v = ma.getVelocities;
    V = vertcat( v{:} );
    compVCorr = ma.computeVCorr;
    VCorr = compVCorr.vcorr;
    VCorr = cell2mat(VCorr);
    VCorr = array2table(VCorr);
    

    cHeader = {'time' 'mean'}; % header
    commaHeader = [cHeader;repmat({','},0,numel(cHeader))]; %insert commas
    commaHeader = commaHeader(:);
    textHeader = cell2mat(commaHeader);

    figure(3)
    plotMeanVCorr = ma.plotMeanVCorr;
    
    % sauvegarde les résultats ( le mkdir ResultFolder est utile une seule
    % fois )
    mkdir ResultFolder
    varFolder = '\ResultFolder';
    ResultFiles = strcat(folder,varFolder);
   
    name(i) = strrep(name (i),'.xls',' ');
    ResultFiles = fullfile(ResultFiles);
   
    cd(ResultFiles);
    
    MSD=cell2mat(MSD);
    MSD=array2table(MSD);
    
    slash = '\';
    finalFile = strcat (ResultFiles,slash,name(i));
    finalFile = char(finalFile);
    finalCSV = append(finalFile,'.csv');
    final = fopen(finalCSV,'w');
    
    % write MSD data to csv file using the specified format
    csvwrite(finalCSV,MSD{:,[1 2]});
    txt=fileread(finalCSV)
    fidMSD=fopen(finalCSV,'wt');
    fprintf(fidMSD,'time, mean_MSD\n');
    fprintf(fidMSD,'%s',txt);
    fclose(fidMSD);
   
    finalCorr = strcat (ResultFiles,slash,name(i),'_Corr');
    finalCorr = char(finalCorr);
    finalFileCorr = append(finalCorr,'.csv');
    finalC = fopen(finalFileCorr,'w');
    
    % write VCorr data to csv file using the specified format
    csvwrite(finalFileCorr,VCorr{:,[1 2]});
    txtVCorr=fileread(finalFileCorr)
    fidVCorr=fopen(finalFileCorr,'wt');
    fprintf(fidVCorr,'time, mean_Vcorr\n');
    fprintf(fidVCorr,'%s',txtVCorr);
    fclose(fidVCorr);
   

    saveas(tracking,suj+'_tracks.png');
    saveas(plotMSD, suj+'_MSD.png');
    saveas(plotMeanVCorr,suj+'_VCorr.png');
    
    clf(f1);
    clf(f2);
end
