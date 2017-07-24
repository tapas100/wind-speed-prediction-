%Power Curve of Enercon Model E-33 Wind Turbine
function [power]=WRAturbinePClinear(v,Prated,Vcutin,Vrated,Vfurling)

%disp('Enercon E-33 model wind turbine')
%Prated=2300; Vcutin=3; Vrated=11; Vfurling=20; 

if(v<Vcutin)
    power=0;
elseif (v>=Vcutin && v<Vrated)
    power=Prated*((v-Vcutin)/(Vrated-Vcutin));
elseif (v>=Vrated && v<Vfurling)
    power=Prated;
elseif (v>=Vfurling) 
    power=0;
else
    disp('No such wind speed.')
end

