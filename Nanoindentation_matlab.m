clc
clear
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%需修改的东西%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Nano_indent_address='D:\大论文\纳米压痕\2022-5-6浆体腐蚀\';%原始数据存储位置,储存方式为每种元素一个文件夹，以元素名字命名，里面图片为tif格式，文件名为元素名称后紧跟次数编号，如1，2，3，
filename='浸泡在10%硫酸钠溶液中水灰比0.58的净浆';%样本名称
Word_address='D:\大论文\纳米压痕\2022-5-6浆体腐蚀\';%Word储存的位置
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filespec_user = (strcat(Word_address,filename,'.doc'));% 设定测试Word文件名和路径
[Word,Document,Content,Selection]=word_active_and_open(filespec_user);%word得激活和创建
L=word_PageSetup(Document);%% 页面设置
Content.Start = 0;%设定光标位置
Paragraphformat = Selection.ParagraphFormat;

Num.figures=1;n_table=1;Num.equation=1;Num.Table=1;Num.refer=1;%用于记录图片、表格、公式、参考文献得个数
%% %%%%%%%%%%%%%%%%%%%%%%%%标题%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
title = strcat(filename,"纳米压痕数据处理报告");
Content.Text = title;L=set_format_title1(Content);%字体和段落格式设定
%%
n_title2=0;n_title3=1;n_title4=1;%用于记录二级，三级，四级标题
Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
Title1 = strcat(num2str(n_title2),'. 数据处理过程描述');n_title2=n_title2+1;
Selection.Text = Title1;L=set_format_title2(Selection);%格式设定

Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
[T1_content] =Title1_content();

for i=1:size(T1_content)%不同段落的写入
    Selection.Text = T1_content{i,1};L=set_format_for_text(Selection);
    Selection.Start = Content.end;Selection.TypeParagraph;
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


for times=1:7
    I=0;
    Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
    Title1 = strcat(num2str(n_title2),'. 第',num2str(times),'次纳米压痕测试');n_title2=n_title2+1;
    Selection.Text = Title1;L=set_format_title2(Selection);%格式设定



   fileadress= strcat(Nano_indent_address,'1-',num2str(times),'\');
   [NUM{times},X{times},Y{times},M_H{times},M_G{times},save_adress_name,I]=reading_orignal_data_from_excell(fileadress,times,I);
  
   Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
   n_rows=str2double('2');n_columns=str2double('1');
   figure_save_name{1,1}=' ';
   figure_title_word=strcat('第',num2str(times),'次纳米压痕测试荷载随深度的变化');
   save_adress_name1{1,1}=save_adress_name{1,1};
   [Document,Selection,n_table,Num]=tables_figures_and_name(Document,Selection,save_adress_name1,figure_save_name,figure_title_word,n_rows,n_columns,Num,n_table);



%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
   n_rows=str2double('3');n_columns=str2double('2');
   figure_save_name{1,1}='纳米压痕测试硬度等高线图';
   figure_save_name{2,1}='纳米压痕测试硬度概率分布图';
   figure_title_word=strcat('第',num2str(times),'次纳米压痕测试硬度');
   save_adress_name{1,1}=save_adress_name{2,1};save_adress_name{2,1}=save_adress_name{3,1};
   [Document,Selection,n_table,Num]=tables_figures_and_name(Document,Selection,save_adress_name,figure_save_name,figure_title_word,n_rows,n_columns,Num,n_table);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
   n_rows=str2double('3');n_columns=str2double('2');
   figure_save_name{1,1}='纳米压痕测试复合响应模量等高线图';
   figure_save_name{2,1}='纳米压痕测试复合响应模量概率分布图';
   figure_title_word=strcat('第',num2str(times),'次纳米压痕测试复合响应模量');
   save_adress_name{1,1}=save_adress_name{4,1};save_adress_name{2,1}=save_adress_name{5,1};
   [Document,Selection,n_table,Num]=tables_figures_and_name(Document,Selection,save_adress_name,figure_save_name,figure_title_word,n_rows,n_columns,Num,n_table);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
   n_rows=str2double('3');n_columns=str2double('2');
   figure_save_name{1,1}='纳米压痕测试弹性功等高线图';
   figure_save_name{2,1}='纳米压痕测试弹性功概率分布图';
   figure_title_word=strcat('第',num2str(times),'次纳米压痕测试弹性功');
   save_adress_name{1,1}=save_adress_name{6,1};save_adress_name{2,1}=save_adress_name{7,1};
   [Document,Selection,n_table,Num]=tables_figures_and_name(Document,Selection,save_adress_name,figure_save_name,figure_title_word,n_rows,n_columns,Num,n_table);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Selection.Start = Content.end;Selection.TypeParagraph;% 定义开始的位置为上一段结束的位置%换行插入内容并定义字体字号
   n_rows=str2double('3');n_columns=str2double('2');
   figure_save_name{1,1}='纳米压痕测试塑性功等高线图';
   figure_save_name{2,1}='纳米压痕测试塑性功概率分布图';
   figure_title_word=strcat('第',num2str(times),'次纳米压痕测试塑性功');
   save_adress_name{1,1}=save_adress_name{8,1};save_adress_name{2,1}=save_adress_name{9,1};
   [Document,Selection,n_table,Num]=tables_figures_and_name(Document,Selection,save_adress_name,figure_save_name,figure_title_word,n_rows,n_columns,Num,n_table);


end


























