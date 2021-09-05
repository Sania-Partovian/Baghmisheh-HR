clear all;

%% Reading data
T = readtable('MonthlyRep.csv');
tData = table2dataset(T);
tData = sortrows(tData,3);
Tname = readtable('BaghNames.xlsx');
tName = sortrows(table2dataset(Tname),3);

%% Variable and data structure initialization 
numerOfWorkingDays = 31;           
name = cell(length(tData.CardCode), 2);
name (:) = {'String'};
hozor = nan(length(tData.CardCode), 2);
dirkard = nan(length(tData.CardCode), 2);
mSaati = nan(length(tData.CardCode), 2);
mSaatiFinal = zeros(length(tData.CardCode), 1);
mRuz = zeros(length(tData.CardCode), 1);
ezTatil = nan(length(tData.CardCode), 2);
ezKol = nan(length(tData.CardCode), 2);
ezNew = nan(length(tData.CardCode), 2);
ezKolFinal = zeros(length(tData.CardCode), 1);
ayabZahab = nan(length(tData.CardCode), 1);
karkard = numerOfWorkingDays.*(ones(length(tData.CardCode),1));
dummyColumn = zeros(length(tData.CardCode), 1);

%% Fetching data and grooming
for i=1:length(tData.CardCode)
    % Saate ezafekari bedune mohasebeye dirkard
    tempEz = str2mat(tData.Parametr1(i));
    if tempEz ~= '-'
        ezKol(i,1) = str2double(tempEz(1:3));
        ezKol(i,2) = str2double(tempEz(5:6));
    else
        ezKol(i,1) = 0;
        ezKol(i,2) = 0;
    end
    
    % Mohasebeye dirkard be hamrahe 2 saat haghe dirkarde mojaz
    tempDir = str2mat(tData.DirKard(i));
    if (tempDir ~= '-' & str2double(tempDir(1:3))>1 & str2double(tempDir(5:6))>0)
        dirkard(i,1) = str2double(tempDir(1:3)) - 2;
        dirkard(i,2) = str2double(tempDir(5:6));
        
    else
        dirkard(i,1) = 0;
        dirkard(i,2) = 0;
    end
    
    % Mohasebeye kolle ezafekari ba ehtesabe dirkard
    ezNew(i,1) = ezKol(i,1) - dirkard(i,1);
    if (ezKol(i,2) >= dirkard(i,2))
        ezNew(i,2) = ezKol(i,2) - dirkard(i,2);
    else
        ezNew(i,1) = ezNew(i,1) - 1;
        ezNew(i,2) = ezKol(i,2) + 60 - dirkard(i,2);
    end
    
    % Gerd kardane saate ezafekari
    ezMin = ezNew(i,2);
    ezMutiplier = 1;
    if ezNew(i,1) < 0
        ezMutiplier = -1;
    end
    if (ezMin < 11)
        ezKolFinal(i) = ezNew(i,1);
    elseif (ezMin < 36)
        ezKolFinal(i) = ezNew(i,1) + (ezMutiplier*0.5);
    else
        ezKolFinal(i) = ezNew(i,1) + (ezMutiplier*1);
    end
        
    
    % Saate morakhasi saati
    tempMosaa = str2mat(tData.Mo_Saati(i));
    if tempMosaa ~= '-'
        mSaati(i,1) = str2double(tempMosaa(1:3));
        mSaati(i,2) = str2double(tempMosaa(5:6));
    else
        mSaati(i,1) = 0;
        mSaati(i,2) = 0;
    end
    
    % Gerd kardane morakhasi saati
    morMin = mSaati(i,2);
    if (morMin < 15)
        mSaatiFinal(i) = mSaati(i,1);
    elseif (morMin < 45)
        mSaatiFinal(i) = mSaati(i,1) + 0.5;
    else
        mSaatiFinal(i) = mSaati(i,1) + 1;
    end

    
    % Ayab Zahab
    tempAyab = str2mat(tData.Parametr5(i));
    ayabZahab(i) = str2double(tempAyab);

    % Morakhasi ruzaneh
    tempMoruz = str2mat(tData.Morak(i));
    mRuz(i) = str2double(tempMoruz);

end



myTable2 = table(str2num(cell2mat(tData.CardCode)), tName.name, tName.fname, karkard, ayabZahab, ezNew, mSaati, mRuz, 'VariableNames',{'CardCode' 'Name' 'fName' 'Karkard' 'AyabZahab' 'EZkoll' 'M_Saati' 'M_Ruz'});
myTable4 = table(mRuz, mSaatiFinal, dummyColumn, dummyColumn, ezKolFinal, ayabZahab, karkard, tName.fname, tName.name, str2num(cell2mat(tData.CardCode)),...
                'VariableNames',{'M_Ruz' 'M_Saati' 'Gheybat' 'KasrKar' 'EZkoll' 'AyabZahab' 'Karkard' 'fName' 'Name' 'CardCode'});

filename = 'outputData.xlsx';
writetable(myTable2,filename,'Sheet',1);


tempNameCode = [tData.CardCode, name];
tempNameCode2 = cell(size(tName));
tempNameCode2 = table(tName.codePers, str2mat(tName.name), str2mat(tName.fname));
tempNameCode2 = sortrows(tempNameCode2);




%% Generate Sheets based on the designate name list

CodeMoaven = [281; 237; 614];
CodeModir = [242; 173];
CodeMali = [156; 174; 203; 217; 232; 260; 269; 613];
CodeForush = [282; 509; 191; 508; 620];
CodeEdari = [154; 184; 194; 287; 300; 314; 503; 605];
CodeHoghugh = [176; 317; 515];
CodeFani = [167; 244; 249; 283; 511; 611; 617];
CodeEjra = [163; 208; 245; 315; 316; 322; 615; 616];
CodeKhadamat = [323; 279; 512; 618; 619; 504; 329; 507];
CodeX = {'CodeMoaven'; 'CodeModir'; 'CodeMali'; 'CodeForush'; 'CodeEdari'; 'CodeHoghugh'; 'CodeFani'; 'CodeEjra'; 'CodeKhadamat'};


writetable(myTable2(ismember(myTable2.CardCode,CodeMoaven),:),filename,'Sheet',2);
writetable(myTable2(ismember(myTable2.CardCode,CodeModir),:),filename,'Sheet',3);
writetable(myTable2(ismember(myTable2.CardCode,CodeMali),:),filename,'Sheet',4);
writetable(myTable2(ismember(myTable2.CardCode,CodeForush),:),filename,'Sheet',5);
writetable(myTable2(ismember(myTable2.CardCode,CodeEdari),:),filename,'Sheet',6);
writetable(myTable2(ismember(myTable2.CardCode,CodeHoghugh),:),filename,'Sheet',7);
writetable(myTable2(ismember(myTable2.CardCode,CodeFani),:),filename,'Sheet',8);
writetable(myTable2(ismember(myTable2.CardCode,CodeEjra),:),filename,'Sheet',9);
writetable(myTable2(ismember(myTable2.CardCode,CodeKhadamat),:),filename,'Sheet',10);


filename2 = 'outputData2.xlsx';
writetable(myTable4,filename2,'Sheet',1);

for i=1:9
    tTempStrip = [];
    tTempCont = eval(CodeX{i});
    for j=1:length(tTempCont)
        tTempStrip = [tTempStrip; myTable4(myTable4.CardCode==tTempCont(j),:)];
        j=j+1;
    end
    writetable(tTempStrip,filename2,'Sheet',i+1);
    i=i+1;
end

