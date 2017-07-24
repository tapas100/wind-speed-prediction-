%Weibull plots for various values of k
clc;
clear all;

U=0:0.1:25;

%Uav=8; 

k=1.6:0.4:3.2; 
%c=zeros(3,1);
c=[10 10 10 10 10];

for i=1:5
    %c(i)=Uav/gamma(1+1/k(i));
    f=wblpdf(U,c(i),k(i));
    switch i
        case 1
        plot(U,f,'LineWidth',2);      
        case 2
        plot(U,f,'LineWidth',2);
        case 3
        plot(U,f,'LineWidth',2);
        case 4
        plot(U,f,'LineWidth',2);
        case 5
        plot(U,f,'LineWidth',2);
    end
    hold all;
end
%text(10,0.05,'k=1.6')%,' \leftarrow sin(\pi)','FontSize',18)
legend('1.6','2.0','2.4','2.8','3.2','Location','NorthEast')
xlabel('Wind Speed')
ylabel('Frequency')