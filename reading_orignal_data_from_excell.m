function [data,X,Y,M_H,M_G,save_adress_name,I]=reading_orignal_data_from_excell(fileadress,times,I)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

M_e0=readtable(strcat(fileadress,'e.txt'));%读取excell数据
M_e=table2array(M_e0);
r_points=str2double(regexp(M_e{26,1},'\d*\.?\d*','match'));
c_points=str2double(regexp(M_e{25,1},'\d*\.?\d*','match'));
N_points=r_points*c_points;
interval_c=str2double(regexp(M_e{25,3},'\d*\.?\d*','match'));
interval_r=interval_c;
depth=(c_points-1)*interval_c;
width=(r_points-1)*interval_r;

M_load_depth0= readtable(strcat(fileadress,'load-depth.xlsx'));%读取excell数据
M_load_depth=table2array(M_load_depth0);
[r0,~]=find(isnan(M_load_depth(:,1)));
r_intial=[1;r0(2:2:end)+1];
r_end=[r0(1:2:end)-1;size(M_load_depth,1)];
data{times,N_points}=0;

I=I+1;
n_fg=figure(I);
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,800,400]);%设置图创大小
t = tiledlayout(r_points,c_points,'TileSpacing','tight','Padding','tight');


for i=1:N_points
    data{times,i}=M_load_depth(r_intial(i):r_end(i),1:2);

    nexttile
    plot(data{times,i}(:,2),data{times,i}(:,1))% hold on
    xx=max(data{times,i}(:,2));
    yy=max(data{times,i}(:,1));
    axis([0,1.1*xx,0,1.1*yy]);

end

xlabel(t,'depth(\mum)','FontName','Times New Roman')
ylabel(t,'load(mN)','FontName','Times New Roman')

%% 加边框
ax2 = axes('Position',get(gca,'Position'),...
           'XAxisLocation','top',...
           'YAxisLocation','right',...
           'Color','none',...
           'XColor','k','YColor','k');
set(ax2,'YTick', []);
set(ax2,'XTick', []);

save_adress_name{I,1}=strcat(fileadress,'load depth');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)

M_H_R0=readtable(strcat(fileadress,'H-R.xlsx'));%读取excell数据
M_H_R1=table2array(M_H_R0);

M_H0{times}=M_H_R1(3:N_points+2,6);
M_G0{times}=M_H_R1(3:N_points+2,7);
M_H=reshape(M_H0{times},r_points,c_points);
M_G=reshape(M_G0{times},r_points,c_points);

M_plastic_work0{times}=M_H_R1(3:N_points+2,10);
M_elastic_work0{times}=M_H_R1(3:N_points+2,11);
M_plastic_work=reshape(M_plastic_work0{times},r_points,c_points);
M_elastic_work=reshape(M_elastic_work0{times},r_points,c_points);



y = linspace(0,width,r_points);
x = linspace(0,depth,c_points);
[X, Y] = meshgrid(x,y);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
I=I+1;
n_fg=figure(I);
contourf(X, Y, M_H,14);
colorbar
ylabel(colorbar,'hardness(GPa)','FontName','Times New Roman')
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('depth(\mum)','FontName','Times New Roman')
ylabel('width(\mum)','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'contour of hardness');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%-------------------------------------------------------------------------
I=I+1;
n_fg=figure(I);
nbins = 25;
histogram(M_H0{times},nbins);
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('hardness(GPa)','FontName','Times New Roman')
ylabel('frequency','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'statistical graph of hardness');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


I=I+1;
n_fg=figure(I);
contourf(X, Y, M_G,14);
colorbar
ylabel(colorbar,'Er(GPa)','FontName','Times New Roman')
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('depth(\mum)')
ylabel('width(\mum)')

save_adress_name{I,1}=strcat(fileadress,'contour of Er');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%-------------------------------------------------------------------------
I=I+1;
n_fg=figure(I);
nbins = 25;
histogram(M_G0{times},nbins);
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('Er(GPa)','FontName','Times New Roman')
ylabel('frequency','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'statistical graph of Er');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

I=I+1;
n_fg=figure(I);
contourf(X, Y, M_elastic_work,14);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
colorbar;
ylabel(colorbar,'elastic work(nJ)','FontName','Times New Roman')
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
xlabel('depth(\mum)','FontName','Times New Roman')
ylabel('width(\mum)','FontName','Times New Roman')
set(gca,'position',[0.07 0.17 0.76 0.81]);

save_adress_name{I,1}=strcat(fileadress,'contour of elastic work');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%-------------------------------------------------------------------------
I=I+1;
n_fg=figure(I);
nbins = 25;
histogram(M_elastic_work0{times},nbins);
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('elastic work(nJ)','FontName','Times New Roman')
ylabel('frequency','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'statistical graph of elastic work');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

I=I+1;
n_fg=figure(I);
contourf(X, Y, M_plastic_work,14);
colorbar
ylabel(colorbar,'plastic work(nJ)','FontName','Times New Roman')
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'position',[0.07 0.17 0.77 0.81]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('depth(\mum)','FontName','Times New Roman')
ylabel('width(\mum)','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'contour of plastic_work');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%-------------------------------------------------------------------------
I=I+1;
n_fg=figure(I);
nbins = 25;
histogram(M_plastic_work0{times},nbins);
set(0,'defaultfigurecolor','w');%设定背景颜色为白色
set(gca,'looseInset',[0 0 0 0]);
set(gcf,'position',[360,200,500,500/c_points*r_points]);%设置图创大小
xlabel('plastic work(nJ)','FontName','Times New Roman')
ylabel('frequency','FontName','Times New Roman')

save_adress_name{I,1}=strcat(fileadress,'statistical graph of plastic_work');
savefig(figure(I),strcat(save_adress_name{I,1},'.fig'));
saveas(figure(I),strcat(save_adress_name{I,1},'.bmp'));
delete(n_fg)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%







