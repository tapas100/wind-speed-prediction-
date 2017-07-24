%This estimates AR parameters for 1984-Time Series Model
%It also simulates and forecasts wind speed using fitted AR model

clc;
clear;

duration=24;T=31;%total no of days
strLocation='AR1984';

%%%%%%%%%%%Read Data%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Reading data from excel...')
windDataM=xlsread('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', strLocation, 'C6:C749' );
windData=windDataM(1:24);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Transformation%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
windDataT=zeros(duration,1);windDataMT=zeros(duration*T,1);
for i=1:24
    windDataT(i)=sqrt(windData(i));    
end
for i=1:duration*T    
    windDataMT(i)=sqrt(windDataM(i));
end
windDataMT(1:24);
% disp('Now writing square root transformed wind speed...')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windDataT, strLocation, 'D6:D29');

%%%%%%%%%%%%%%%%%%%%%%%%%%Hourly Mean%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
hourlySum=zeros(24,1);hourlyData=zeros(24,31);
hourlyMean=zeros(24,1);hourlySD=zeros(24,1);
for i=1:24
    for j=1:31
        hourlySum(i)=hourlySum(i)+windDataMT(24*(j-1)+i);
        hourlyData(i,j)=windDataMT(24*(j-1)+i);
    end  
    %hourlyMean(i)=hourlySum(i)/31;
    hourlyMean(i)=mean(hourlyData(i,:));
    hourlySD(i)=std(hourlyData(i,:));
end
% disp('Now writing hourly mean and standard deviation...')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', hourlyMean, strLocation, 'E6:E29');
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', hourlySD, strLocation, 'F6:F29');

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%Standardization%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
windDataTS=zeros(duration,1);
for i=1:24
    windDataTS(i)=windDataT(i)-hourlyMean(i);
end
% disp('Now writing standardized wind speed...')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windDataTS, strLocation, 'G6:G29');

%Calculation of BIC requires calculation of 10th order Autocorrelation
p=2;

%%%%%%%%%%%%%%%%%%%%%%%%%%Autocorrelations%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Eq. 11-12, P-3,1984-timeseries
c0=0;
for t=1:24
    c0=c0+windDataTS(t)*windDataTS(t);
end
 
c=zeros(p,1);
for k=1:p
    for t=(k+1):24
        c(k)=c(k)+windDataTS(t-k)*windDataTS(t);
    end
end
% disp('Now writing Autocovariance and Autocorrelation....')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', c0, strLocation, 'H6:H6');
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', c, strLocation, 'H7:H8');

%sample correlation coefficients required for Yule-Walker recursion
r(1)=c(1)/c0;
r(2)=c(2)/c0;
R=[1 r(1); r(1) 1];
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', c, strLocation, 'I7:I8');
%partial autocorrelation coefficients
pacf(1)=r(1);
pacf(2)=(r(2)-pacf(1)*r(1))/(1-pacf(1)*r(1));
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', pacf', strLocation, 'J7:J8');

%%%%%%%%%%%%%%%AR coefficients using Yule-Walker recursion%%%%%%%%%%%%%%%%%
ARCoeff=r/R;
wnVar=1;
for k=1:p
    wnVar=wnVar*(1-pacf(k)*pacf(k));
end
wnVar=(c0/(24*31-26))*wnVar;
ARCoeff=[ARCoeff wnVar];
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', ARCoeff', strLocation, 'K6:K8');


%%%%%%%%%%%%SIMULATION OF WIND SPEED%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
windSpeedSim=zeros(duration,1);
%normrnd(mean,variance,m,n) m x n matrix simualated data
%standard normal distribution has mean=0 and s.d.=1
%Simulate Nt x Ns standard normal errors Nt: time Ns: location
stdNormWS=normrnd(0,1,24,1);%Gaussian random numbers
% disp('Now writing Gaussian Random numbers...')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', stdNormWS, strLocation, 'L6:L29');

%Transformed Wind speed for the first hour
estVar=sqrt(c0/(24*(T-1)));
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', estVar, strLocation, 'M6:M6');
windSpeedSim(1)=hourlyMean(1)+estVar*stdNormWS(1);
%Transformed wind speed for second hour
windSpeedSim(2)=hourlyMean(2)+r(1)*(windSpeedSim(1)-hourlyMean(1))...
                + estVar*sqrt(1-r(1)^2)*stdNormWS(2);
%Transformed wind speed for t hour
for t=3:duration
    windSpeedSim(t)=hourlyMean(t)+ARCoeff(1)*(windSpeedSim(t-1)-hourlyMean(t-1))...
                + ARCoeff(2)*(windSpeedSim(t-2)-hourlyMean(t-2))...
                + estVar*stdNormWS(t);
end
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windSpeedSim, strLocation, 'N6:N29');

%Wind Speed
windSpeedSimF=zeros(duration,1);
for i=1:24
    windSpeedSimF(i)=windSpeedSim(i)^2;
end
% disp('Now writing standardized wind speed...')
% xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windSpeedSimF, strLocation, 'O6:O29');

%%%%%%%%%%%%%%%%%%%%%%%%%%%FORECASTING OF WIND SPEED%%%%%%%%%%%%%%%%%%%%%%%
forHorizon=6; %for 6 hours in advance
windSpeedFt=zeros(forHorizon+2,1);windSpeedF=zeros(forHorizon+2,1);
windSpeedFt(1)=windDataTS(23)+hourlyMean(23);
windSpeedFt(2)=windDataTS(24)+hourlyMean(24);
for h=3:8
    windSpeedFt(h)=ARCoeff(1)*windSpeedFt(h-1)+ARCoeff(2)*windSpeedFt(h-2)+hourlyMean(h);
    windSpeedF(h)=windSpeedFt(h)^2;
end
disp('Now writing forecasted wind speed...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windSpeedFt(3:8), strLocation, 'P32:P37');
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPTimeSeries.xlsx', windSpeedF(3:8), strLocation, 'Q32:Q37');








