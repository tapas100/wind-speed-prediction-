%This program fits ARMA(p,q) model for specified time series data

clc;
clear all;

location=input('Enter location(1-S1, 2-S2, 3-S0)...');
disp('Now reading time series data...')

switch(location)
    case 1
       windDataS=xlsread('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', 'houly', 'D3:D503');
       p=3;q=3;
       modelWS=arima(p,0,q);  
       colRangeAR='R3:R5';
       colRangeMA='S3:S5';
       colRangeCT='R6:R6';
       colRangeVR='S6:S6';
       colRangeVC='B4:I11';
       colRangeE='D3:D503';
       colRangeV='E3:E503';
    case 2
        windDataS=xlsread('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', 'houly', 'G3:G503');
        p=2;q=2;
        modelWS=arima(p,0,q);  
        colRangeAR='W3:W4';
        colRangeMA='X3:X4';
        colRangeCT='W5:W5';
        colRangeVR='X5:X5';
        colRangeVC='L4:Q9';
        colRangeE='G3:G503';
        colRangeV='H3:H503';
    case 3
        windDataS=xlsread('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', 'houly', 'J3:J503');
        p=2;q=3;
        modelWS=arima(p,0,q); 
        colRangeAR='AB3:AB4';
        colRangeMA='AC3:AC5';
        colRangeCT='AB5:AB5';
        colRangeVR='AC6:AC6';
        colRangeVC='V4:AB10';
        colRangeE='J3:J503';
        colRangeV='K3:K503';
end
disp('data reading over...Now performing ARMA fit')
fprintf('\n');
disp('ARMA fitting starts...it will take a while..have patience')

[fit,varCov,LogL,info] = estimate(modelWS,windDataS);
[E,V]=infer(fit,windDataS);%infers residual and  conditional variances
sizeVarCov=p+q+2;
% fit
% varCov
% LogL
% info

%Uncomment code below for writing ARMA parameters to excel

%AR coefficients separation
modelParamAR=cell2mat(fit.AR);
modelParamAR2=zeros(p,1);
for i=1:p
    modelParamAR2(i)=modelParamAR(p+1-i);
end
fprintf('\n\n\n')
disp('Now writing AR parameters...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', modelParamAR2, 'houly', colRangeAR);

%MA coefficients separation
modelParamMA=cell2mat(fit.MA);
modelParamMA2=zeros(q,1);
for i=1:q
    modelParamMA2(i)=modelParamMA(q+1-i);
end
disp('Now writing MA parameters...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', modelParamMA2, 'houly', colRangeMA);


%Constant
modelParamCon=fit.Constant;
disp('Now writing Constant parameter...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', modelParamCon, 'houly', colRangeCT);

%Variance
modelParamCon=fit.Variance;
disp('Now writing Variance parameter...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', modelParamCon, 'houly', colRangeVR);

disp('Now writing variance-covariance matrix')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', varCov, 'VarCov', colRangeVC);

%produces an upper triangular matrix R such that RRt=varCov
R=chol(varCov);
%lower triangular matrix
L=chol(varCov,'lower');
%orthogonal transformation matrix B
B=L;

disp('Now writing cholesky decomposition upper triangular matrix...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', R, 'UT', colRangeVC);

disp('Now writing cholesky decomposition lower triangular matrix...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', L, 'LT', colRangeVC);

disp('Now writing E and V...')
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', E, 'Residual', colRangeE);
xlswrite('C:\Users\Rajat\Documents\MATLAB\IOFiles\WPPSimulated2.xlsx', V, 'Residual', colRangeV);

disp('Thank you for your patience...')



