%Wind data simulation using Marginal distribution (Weibull)
%Writes to WPPSimulated4.xlsx 


clc;
clear;

datacount=8760;

timeFrame=zeros(datacount,1);
for i=1:datacount
    timeFrame(i)=i;
end


R=rand(datacount,1);

%Values from 2015-IJPES-GravSearch
k=2.2; c=11; %=15*((10/80)^0.143);
%c=15;alpha=0.143 for smooth grass covered terrain

%Simulate wind speed using Weibull distribution
fWeibull=c(1)*(-log (R)).^(1/k(1));

colRange1='B3:B8762';
colRange2='C3:C8762';    
colRange3='D3:E3';           
    
meanSD=[mean(fWeibull) std(fWeibull)];

disp('Now writing wind speed data...')
xlswrite('WPPSimulated4.csv', timeFrame, 'Simulated', colRange1);
xlswrite('WPPSimulated4.csv', fWeibull, 'Simulated', colRange2);
xlswrite('WPPSimulated4.csv', meanSD, 'Simulated', colRange3);










