
%% MATLAB - clear trackchanges
% can use directly in the file because MATLAB will save the output in a new

% Procedure before submission
% 1) Clean BibTex
% 2) Remove Trackchanges
% 3) Produce WORD Version
% 4) Check grammar in WORD
% 5) Repeat 2 and 3 with corrected version from 4)

function LaTex_functions

fclose('all')

% Settings ----------------------------
flag_clean_BibTex = 1; % Cleans BibText undesirable entries, 1-clean, 0-don't
force_clean = 0;

flag_trackchanges = 1; % Cleans 'trackchages': 1-clean, 0-don't

flag_Latex2Word = 0; % Generate Word file, 1-yes, 0-no
flag_Word2Latex = 0; % Generate Latex-file, 1-yes, 0-now

flag_import_Bib2DB_ACCESS = 0; % ACCESS DB: import, Citation_keys, Year, Title and Abstract
force_import = 0;

% ------------------------------------

if (flag_trackchanges == 1 | flag_Latex2Word == 1)
    defaultdir = GetDefaultdir(1);
    [filename,filedir,FilterIndex] = uigetfile('*.tex','Choose the Latex (.tex) file to convert',defaultdir)
end

%% %%%%%%%%% ONLY CHANGE inputs above %%%%%%%%%%%%%%%%%%%%%

% Cleans BibText undesirable entries
defaultdir = GetDefaultdir(2);

if (flag_trackchanges || flag_Latex2Word || flag_Word2Latex || flag_clean_BibTex || flag_import_Bib2DB_ACCESS)
    [filenameBibTex,filedirBiBTex,FilterIndex] = uigetfile('*library.bib','Choose the BibTex (.bib) file use',defaultdir)
end

if flag_clean_BibTex == 1
    CleanBibTex(filedirBiBTex,filenameBibTex,force_clean)
end

% Latex 2 WORD
if flag_trackchanges | flag_Latex2Word      
    outputfile = [filename(1:end-4),'_trackchangesCleaned.tex'];
    if flag_trackchanges
        % Clean 'trackchages' from .tex file: deals with \remove, \add and \change
        % (generates new file in the same directory of original tex file - the original tex file is not changed)
        CleanLaTex_trackchanges(filedir,filename,outputfile)
    end    
    if flag_Latex2Word
    % Generate "WordDoc" (generates new file) - NOTE: need to remove all
    % https://www.mathworks.com/matlabcentral/fileexchange/56283-activex-word-control-base-commands?focused=6592993&tab=example
    % indentions from .tex
    % https://msdn.microsoft.com/en-us/library/office/ff840967.aspx
        Gen_Latex2Word(filedir,outputfile,filedirBiBTex,filenameBibTex)
    end    
end

% WORD to Latex
if flag_Word2Latex
    defaultdir = GetDefaultdir(1);
    [filenameWORD,filedirWORD,FilterIndex] = uigetfile({'*.docx';'*.doc'},'Choose the WORD (.doc or *docx) file to convert',defaultdir)
    Gen_Word2Latex(filenameWORD,filedirWORD,filedirBiBTex)
end

if flag_import_Bib2DB_ACCESS
   % Import BibTex to ACCESS DB: import, Citation_keys, Year, Title and Abstract
   Import_BiBtex_2_DB_ACESS(filedirBiBTex,filenameBibTex,force_import) 
end


%% Cleans BibText undesirable entries,
function CleanBibTex(filedir,filename,forceclean)

fileadd = [filedir,filename];

logfile_nam = [filedir,'BiBtextCleaning_LastRunLog.txt'];

fileinfo = dir(fileadd);
datefilemodified = fileinfo.date;
datefilemodified_num = datenum(datefilemodified);

lastrun = fileread(logfile_nam);
lastrun_datenum = datenum(lastrun);

if lastrun_datenum < datefilemodified_num | forceclean

str = importdata(char(fileadd)); 
fields2rem = {'url','month'};

fields2check_inyear = [9990,9991,9992,9993]; % 9991 - submitted, 9992 - undr review, 3- in press
key2print = {'in preparation','submitted','under review','in press'};

linechar = '@Preamble{ " \newcommand{\noop}[1]{} " }';
if ~strcmp(str(1),linechar)
    forsize = ['%',num2str(numel(linechar)),'s\n'];           
    fileID = fopen(fileadd,'w');
    fprintf(fileID,forsize,linechar);
    fclose(fileID);     
end

h1 = waitbar(0,'Cleaning BibTex...');
for i = 1:numel(str)
    line = str(i);
     
    % check if the line has the undesirable entry      
    a = 0;       
    for j = 1:numel(fields2rem)
        fieldsize = numel(char(fields2rem(j))); 
        linechar = char(line);    
        if numel(linechar) ~= 0 && numel(linechar) > fieldsize && strcmp(linechar(1:fieldsize),char(fields2rem(j)))
           a = a + 1;
        end        
    end
    
    % check if the paper is year is: 9991 - submitted, 9992 - undr review, 3- in press
    b = 0;
    if numel(linechar) ~= 0 && numel(linechar) >= numel(char('year = {9999}')) && strcmp(linechar(1:4),'year')
        for j = 1:numel(fields2check_inyear)      
            yearstrg = char(['year = {',num2str(fields2check_inyear(j)),'}']);
            fieldsize = numel(yearstrg);   
            if numel(linechar) ~= 0 && numel(linechar) >= fieldsize && strcmp(linechar(1:fieldsize),char(yearstrg))
               yearstr = ['\noop{3001}',key2print{j}];
               b = b + 1;
            end        
        end
    end
    
    % print if not one of the entires to delete  
   if b == 1
       linechar = ['year = {',yearstr,'}'];
   end
   if a == 0
        forsize = ['%',num2str(numel(linechar)),'s\n'];           
        fileID = fopen(fileadd,'a');
        fprintf(fileID,forsize,linechar);
        fclose(fileID);
   end
          
    waitbar(i / numel(str),h1,sprintf('%24s %5i','Cleaning BibTex: line = ',i))
end
waitbar(i / numel(str),h1,sprintf('%24s %5i %11s','Cleaning BibTex: line = ',i,' (finished)'))
close(h1)

% log file
fileID2 = fopen(logfile_nam,'w');
timenow =datetime('now','Format','d-MMM-y HH:mm:ss');
timenow_str = datestr(timenow);
fprintf(fileID2,timenow_str,'%21s');
fclose(fileID2);

end



%% Cleans 'trackchages' from .tex file: deals with \remove, \add and \change
% (generates new file)
function CleanLaTex_trackchanges(filedir,filename,outputfile)

paper_file = [filedir,filename];
fileID_read = fopen(paper_file,'r');

outputfile = [filedir,outputfile];
fileID_write = fopen(outputfile,'w');

paper_line_i = fgetl (fileID_read);
file_end = char('%% LATEX_document_END'); % add this line to the end of my file

% go line by line to clean and print the new file
h2 = waitbar(0,'Cleaning trackchanges...');
y = 1;

while ~strcmp(paper_line_i,file_end)
    
    i_all = [];
    newstr = [];
    if  ~isempty(paper_line_i) && ~strcmp(paper_line_i(1),'%')

        i_remove = strfind(paper_line_i,'\remove[');
        i_add = strfind(paper_line_i,'\add[');
        i_change = strfind(paper_line_i,'\change[');

        % finding the type of operation, e.g. \remove, \add, \change (and
        % storing the info)
        i_all = [i_remove,i_add,i_change];
        
        if ~isempty(i_all)
        
            i_all_oper = [i_remove./i_remove i_add./i_add.*2 i_change./i_change.*3];
            [i_all_order index] = sort(i_all);
            flag_oper = i_all_oper(index);

            remove_text = []; 
            for i = 1:numel(i_all_order)

                Start_cut = i_all_order(i);
                if i ~= numel(i_all_order)
                     text_fraction = i_all_order(i+1)-1;
                else
                     text_fraction = numel(paper_line_i);
                end
                [End_cut ExtraCut] = FindEnd(Start_cut,text_fraction,paper_line_i,flag_oper(i));
                remove_text = [remove_text; flag_oper(i) Start_cut, End_cut, ExtraCut];       
            end
            
            newstr = paper_line_i(1:remove_text(1,2)-1);
            for j = 1:numel(remove_text(:,1))
                if remove_text(j,1) ~= 1
                   port_newst = paper_line_i(remove_text(j,3)+1:remove_text(j,4)-1);
                   newstr = [newstr,port_newst];
                end
                   if j < numel(remove_text(:,1))
                       if remove_text(j,1) ~= 1
                            port_newst = paper_line_i(remove_text(j,4)+1:remove_text(j+1,2)-1);
                       else
                            port_newst = paper_line_i(remove_text(j,3)+1:remove_text(j+1,2)-1);
                       end
                       newstr = [newstr,port_newst];   
                   end
            end
            if remove_text(j,1) ~= 1
                port_newst = paper_line_i(remove_text(j,4)+1:numel(paper_line_i));
            else
                 port_newst = paper_line_i(remove_text(j,3)+1:numel(paper_line_i));
            end
            newstr = [newstr,port_newst]; 
        else
            newstr = paper_line_i; 
        end
        
        else
        newstr = paper_line_i;  

    end

% Final Clean Up (remove to spaces, two .., etc)
newstr = CleanUp(newstr);


fprintf(fileID_write,'%s\n',newstr);

paper_line_i = fgetl (fileID_read);

waitbar(y/3000,h2,sprintf('%30s %5i','Cleaning trackchanges: line = ',y))
y = y + 1;
end

waitbar(3000/3000,h2,sprintf('%30s %5i %11s','Cleaning trackchanges: line = ',y,' (finished)'))

% print the last line
fprintf(fileID_write,'%s',file_end);

fclose(fileID_read);
fclose(fileID_write);
close(h2)

% find End_cut and ExtraCut
function  [End_cut ExtraCut] = FindEnd(Start_cut,text_fraction,paper_line_i,flag_oper);

% flag_oper -> 1-remove, 2-add, 3-change

    [brack_open brack_close] = findBrackets(Start_cut,text_fraction,paper_line_i,flag_oper);
    
    if flag_oper == 1 % \remove
            End_cut = Start_cut + brack_close - 1;
            ExtraCut = NaN(1);
    elseif flag_oper == 2 %\add
             End_cut = Start_cut + brack_open - 1;
             ExtraCut = Start_cut + brack_close - 1;
    else    % change
             End_cut = Start_cut + brack_open - 1;
             ExtraCut = Start_cut + brack_close - 1;
    end

    
% find appropriate brackets
function [brack_open brack_close] = findBrackets(Start_cut,text_fraction,paper_line_i,flag_oper);

% flag_oper -> 1-remove, 2-add, 3-change

    Left_brack = find(paper_line_i(Start_cut:text_fraction) == '{');
    Right_brac = find(paper_line_i(Start_cut:text_fraction) == '}');
    
    brac_all = [Left_brack Right_brac];
    brac_all_calc = [Left_brack./Left_brack  Right_brac./Right_brac.*-1];
    [brac_all_sort index] = sort(brac_all);
    brac_all_calc_sort = brac_all_calc(index);
    
    cumsum_brackcalc = cumsum(brac_all_calc_sort);
    
    cumsum_brackcalc_art = cumsum_brackcalc(1);
    for k=2:numel(cumsum_brackcalc) % needs this artifact because values 1 in cumsum_brackcalc_art may arise when there are inner brackets
        addval = cumsum_brackcalc(k);
        if cumsum_brackcalc(k-1)==2
          addval = 1.1;  
        end
        cumsum_brackcalc_art = [cumsum_brackcalc_art,addval];
    end
    
    OutsideStart = find(cumsum_brackcalc_art == 1);
    OutsideClose = find(cumsum_brackcalc_art == 0);
    
    if flag_oper ~= 3 % \remove and \add
        brack_open = brac_all_sort(OutsideStart(1));
        
        if isempty(OutsideClose)
            OutsideClose
        end
        brack_close = brac_all_sort(OutsideClose(1));
    else % \change
        if numel(OutsideStart) == 1
           msgbox(['Problem with citation key not found: "',char(paper_line_i),'"']); 
        end
        brack_open = brac_all_sort(OutsideStart(2));
        brack_close = brac_all_sort(OutsideClose(2));
    end
   
    
%% Generate "WordDoc" (generates new file)
function Gen_Latex2Word(filedir,filename,filedirBiBTex,filenameBibTex)

paper_file = [filedir,filename];

% Preparation of inexing: citations and labels/references
IndexingCitations(filedirBiBTex,filenameBibTex); % Prepare Citation indexes for rapid access
InedexingLabelsRef(filedir,paper_file)


% Start working on the document now
fileID_read = fopen(paper_file,'r');
file_outputname = [filedir,filename(1:end-3),'doc'];
paper_line_i = fgetl (fileID_read);
file_end = char('%% LATEX_document_END'); % add this line to the end of my file


% Generate Word Doc
FontName = 'Arial';

word = actxserver('Word.Application');      %start Word
word.Visible = 1;                            %make Word Visible
document=word.Documents.Add;                %create new Document
selection=word.Selection;                   %set Cursor
selection.Font.Name=FontName;          %set Font
selection.Font.Size=12;                      %set Size

selection.Pagesetup.RightMargin = 100;   %set right Margin to 1cm
selection.Pagesetup.LeftMargin = 100;    %set left Margin to 1cm
selection.Pagesetup.TopMargin = 110;     %set top Margin to 1cm
selection.Pagesetup.BottomMargin = 110;  %set bottom Margin to 1cm
                                            %1cm is circa 28.34646 points
selection.Paragraphs.LineUnitAfter = 0.2;    %sets the amount of spacing
                                            %between paragraphs(in gridlines)
selection.Paragraphs.LineSpacingRule = 0.5;
                                            
latexMainComands_1 = {'\title{','\section{','\subsection{','\subsubsection{','\paragraph{','\subparagaph'};                                 
comi_1 = numel(latexMainComands_1);
% go line by line to clean and print the new file
h3 = waitbar(0,'Generating Latex2Word...');
y = 1;
flag_print = 0;
title_fag = 0;
abstract_flag = [0 0];
DocStart = 0;

selection.ParagraphFormat.Alignment = 9;     %Fully jusified

eqFlag_centre = 0;

while ~strcmp(paper_line_i,file_end)
   
paper_line_i = fgetl (fileID_read);
blackspaces = isspace(paper_line_i);

% Remove initial spaces
if ~isempty(blackspaces) & blackspaces(1) == 1
    while blackspaces(1) == 1
        paper_line_i = paper_line_i(2:end);
        blackspaces = isspace(paper_line_i);
        if isempty(blackspaces)
            break
        end
    end
end

len_line = numel(paper_line_i);
flagBold = 0;


if  ~isempty(paper_line_i) && ~strcmp(paper_line_i(1),'%')
    if numel(char(paper_line_i)) >= 4 && ~strcmp(paper_line_i(1:4),'\hl{')
        % 1) find title (to print)
            if title_fag == 0
                title_find = strfind(paper_line_i,'\title{');
                if ~isempty(title_find)
                    title_fag = 1;
                    newstr = paper_line_i(8:len_line-1);
                    flag_print = 1;
                    flagBold = 1;
                end
            end

        % 2) find abstract (to print)
            if abstract_flag(1) == 1 & abstract_flag(2) == 0
                flag_print = 1;
                newstr = paper_line_i;
            end
            if abstract_flag(1) == 0
                comand_find = strfind(paper_line_i,'\begin{abstract}');
                if ~isempty(comand_find)
                    abstract_flag(1) = 1;
                    newstr = 'ABSTRACT';
                    flag_print = 1;
                    selection.TypeParagraph;        %linebreak
                    flagBold = 1;
                end
            end
            if abstract_flag(2) == 0
                comand_find = strfind(paper_line_i,'\end{abstract}');
                if ~isempty(comand_find)
                    abstract_flag(2) == 0;
                    flag_print = 0;
                    abstract_flag = [2,2];
                end
            end

         % 3) begin document
          if DocStart == 1
              newstr = paper_line_i;
              flag_print = 1;
               for coman_i = 1:comi_1 
                comand_find = strfind(paper_line_i,latexMainComands_1(coman_i));
                    if ~isempty(comand_find)
                        len = numel(char(latexMainComands_1(coman_i)));
                        newstr = paper_line_i(len+1:end-1);
                        selection.TypeParagraph;        %linebreak
                        flagBold = 1;
                    end
               end                     
          elseif DocStart == 0
                comand_find = strfind(paper_line_i,'\section{Introduction}'); % marks the start of the main document
                if ~isempty(comand_find)
                    DocStart = 1;
                    flag_print = 1;
                    newstr = paper_line_i(10:end-1);
                    selection.TypeParagraph;        %linebreak
                    flagBold = 1;
                end
          end 

          % check if there is an equation to write in a new paragraph
          if ~isempty(paper_line_i) && numel(paper_line_i) >= 16 && strcmp(paper_line_i(1:16),'\begin{equation}') % to activate equation mode for the next line
             eqFlag_centre = -1; % don't print line
          elseif numel(paper_line_i)>=14 && strcmp(paper_line_i(1:14),'\end{equation}')
             eqFlag_centre = -1; % don't print line
          end

    else
        flag_print = 0;
    end
else
   flag_print = 0; 
end
  
 
% Stop printing when reaching the figures and tables in the end
if ~isempty(paper_line_i) 
    if numel(char(paper_line_i)) >= 13 && strcmp(paper_line_i(1:12),'\begin{table')
        break
    elseif numel(char(paper_line_i)) >= 14 && strcmp(paper_line_i(1:13),'\begin{figure') 
        break
    end
end

% PRINT
if flag_print == 1
     newstr = CleanLatexCom(newstr); % Clean some commands (e.g. \% adn \par)
     newstr = InputCite(filedirBiBTex,newstr,y); % Add citations
     newstr = CleanUp(newstr); % clean document, e.g. '..', ' .'
     newstr = InputRef(filedir,newstr); % add refs from labels database
     
     eqFlag = 0;
     if ~isempty(newstr)
        symbol = strfind(newstr,'$'); % check if there are equations inside
        equatloc = [];
        if ~isempty(symbol)
            for j = 1:numel(symbol)
                if symbol(1) == 1 || ~strcmp(newstr(symbol(j)-1),'\')
                   eqFlag = 1; 
                   equatloc = [equatloc,symbol(j)];
                end
            end
        end
        
        if eqFlag_centre == 0
            if  ~eqFlag % Tex only; no equations
                if abstract_flag(2) == 2 & flagBold ~=1 % to add indention (not in abstract or titles of sections)
                    newstr = ['     ',newstr];
                end
                  selection.TypeText(newstr);     % print
                  if flagBold == 1;
                        selection.MoveLeft(1,numel(newstr),1);
                        selection.Font.Bold = 1;
                        selection.Font.Size = 16; 
                        selection.MoveRight(1,numel(newstr) + 2);                    %5=row mode
                        selection.Font.Size = 12; 
                        selection.Font.Bold = 0;
                        selection.TypeParagraph;        %linebreak
                        flagBold = 0;
                  end
                  selection.TypeParagraph;        %linebreak
            elseif eqFlag % inline text
                flageqi = 0;
                equatloc_ref = [0, equatloc, numel(newstr)];
                if equatloc(1) == 1
                    flagtxti = 1; % there an equation at the start
                end
                for j = 1: numel(equatloc)+1
                    if flageqi == 0 %text
                        txtprint = newstr(equatloc_ref(j)+1:equatloc_ref(j+1)-1);
                        selection.TypeText(txtprint);     % print
                        flageqi = 1;
                    else
                        eqprint = newstr(equatloc_ref(j)+1:equatloc_ref(j+1)-1); 
                        eqprint = ConvertMaths(eqprint); % convert syntax to word
                        document.OMaths.Add(word.selection.Range);
                        eqn = document.OMaths.Item(1);
                        word.selection.TypeText(eqprint);          
                        eqn.BuildUp()
                        selection.MoveRight(1,1);

                        flageqi = 0;
                    end
                end
                 selection.TypeParagraph;        %linebreak
            end
        elseif eqFlag_centre == -1
            selection.TypeParagraph;        %linebreak
        else    
            eqprint = ConvertMaths(newstr); 
            eqprint = PrepEquWord(eqprint); % convert syntax to word
            document.OMaths.Add(word.selection.Range);
            eqn = document.OMaths.Item(1);
            word.selection.TypeText(eqprint);          
            eqn.BuildUp()
            selection.MoveRight(1,1);
            selection.TypeParagraph;        %linebreak
        end

          if DocStart == 0
            flag_print = 0;
          end 
     end
    end

    if ~isempty(paper_line_i) && numel(paper_line_i) >= 16 && strcmp(paper_line_i(1:16),'\begin{equation}') % to activate equation mode for the next line
        eqFlag_centre = 1; % don't print line
    elseif numel(paper_line_i)>=14 && strcmp(paper_line_i(1:14),'\end{equation}')
        eqFlag_centre = 0; % don't print line
    end
   
    waitbar(y/3000,h3,sprintf('%30s %5i','Generating Latex2Word: line = ',y))
    y = y + 1;
   
end

waitbar(3000/3000,h3,sprintf('%30s %5i %11s','Generating Latex2Word: line = ',y,' (finished)'))

fclose(fileID_read);

document.SaveAs2(file_outputname);          %save Document
word.Quit();                         %close Word

fclose('all') 
close(h3)

% Clean Latex Commands
function newstr = CleanLatexCom(newstr)

com_pointonwards = {'%','\label{'}; % remove all the text from this point onwards, e.g. comments
com_2delete = {'\par','\noindent ','\\'}; % 1) to commands to delete
com_2replace = {'\%'}; % 2) commands to replace
com_replacement = {'%'}; % 2.1) replacement to the above commands
com_brack = {'\textit{','\ce{'}; % 3) commands to do something withing brackets

% com_comment
for i = 1:numel(com_pointonwards)
    comloc_all = strfind(newstr,char(com_pointonwards(i)));
    if ~isempty(comloc_all)
        for j = 1:numel(comloc_all)
            comloc = comloc_all(j);
            if  comloc == 1 || ~strcmp(newstr(comloc-1),'\') % to avoid the commands, e.g. \%
                newstr = newstr(1:comloc-1);
                break
            end 
        end
    end
end

% com_delete
for i = 1:numel(com_2delete)
    comloc = strfind(newstr,char(com_2delete(i)));   
    if ~isempty(comloc)
        newstr = [newstr(1:comloc-1), newstr(comloc + numel(char(com_2delete(i))):end)];
    end 
end

% com_replace
for i = 1:numel(com_2replace)
    comloc = strfind(newstr,char(com_2replace(i)));   
    if ~isempty(comloc)
        newstr = [newstr(1:comloc-1), char(com_replacement(i)), newstr(comloc + numel(char(com_2replace(i))):end)];
    end 
end

% com_bracks (now it is just getting rid of these commands)
for i = 1:numel(com_brack)
    lenCom = numel(char(com_brack(i)));
    comloc_all = strfind(newstr,char(com_brack(i)));
    for j = 1:numel(comloc_all)
        comloc = comloc_all(1);
        findBracR = strfind(newstr(comloc:end),'}'); 
        if ~isempty(comloc)
            refpoint = comloc + findBracR(1);  
            newstr = [newstr(1:comloc-1), newstr(comloc+lenCom:refpoint-2), newstr(refpoint:end)];
        end 
        comloc_all = strfind(newstr,char(com_brack(i)));
        if isempty(comloc_all)
            break
        end
    end
end


% Index BiBtex for rapid access to citations
function IndexingCitations(filedirBiBTex,filenameBibTex) % Prepare Citation indexes for rapid access

BiBfile = [filedirBiBTex,filenameBibTex];
CitationIndex_file = [filedirBiBTex,'Citation_Index.txt'];

Bibdir = dir(BiBfile);
dateModBib = datenum(Bibdir.date);
Citdir = dir(CitationIndex_file);

run = 0;

if isempty(Citdir)
    run = 1;
else
    dateCit = datenum(Citdir.date);
    if dateCit < dateModBib
        run = 1;
    else
        run = 0;
    end
end


if run == 1

    fid_BiB = fopen(BiBfile,'r');
    fid_citationIndex = fopen(CitationIndex_file,'w');
    fclose(fid_citationIndex);
    
    BiBtxt = fgetl(fid_BiB);
    BiBtxt = fgetl(fid_BiB);

    linenum = 1;

    h5 = waitbar(0,'Indexing BiBtex entries...');
    Citation_key = []; % citation key
    Citation_year = []; % year
    CitationAuthorsnum = []; % number of authors
    CitationAuthors = []; % authors
    
    while ~strcmp(num2str(BiBtxt),'-1')
        if ~isempty(BiBtxt) && strcmp(BiBtxt(1),'@')
          open = strfind(BiBtxt,'{') + 1 ;
          close = strfind(BiBtxt,',') - 1;
          Citation_key = BiBtxt(open:close); 
          
          fflag = 0;
          y_yes = 0;
          a_yes = 0;
          while fflag < 2
              linenum = linenum + 1;
              BiBtxt = fgetl(fid_BiB);
              
              if numel(BiBtxt) > 10 && strcmp(BiBtxt(1:10),'author = {')
                  
                authstart = strfind(BiBtxt,'{');
                authend = strfind(BiBtxt,'}');
                
                authors = BiBtxt(authstart(1)+1:authend(end)-1);
                findAnd = strfind(authors,' and ');
                
                
                CitationAuthorsnum = numel(findAnd) + 1;
                commafind = strfind(authors, ', ');
                if CitationAuthorsnum == 1
                    if ~isempty(commafind)
                        CitationAuthors = authors(1:commafind(1)-1);
                    else
                        CitationAuthors = authors;
                    end
                elseif CitationAuthorsnum == 2
                    if numel(commafind == 1)
                        commafind = [commafind, numel(authors) + 1]; 
                    end
                    emptyspaces = strfind(authors,' ');
                    CitationAuthors = [authors(1:commafind(1)-1),'_and_',authors(findAnd(1)+ 5:commafind(2)-1)];
                elseif CitationAuthorsnum > 2 
                    CitationAuthors = [authors(1:commafind(1)-1),'_et_al.'];
                end              
                  fflag =  fflag + 1;
                  a_yes = 1;
              elseif numel(BiBtxt) > 8 && strcmp(BiBtxt(1:8),'year = {')
                  ystart = strfind(BiBtxt,'{');
                  yend = strfind(BiBtxt,'}');           
                  Citation_year = BiBtxt(ystart(1)+1:yend(end)-1); 
                  fflag =  fflag + 1;
                  y_yes = 1;
              end
              
              if ~isempty(BiBtxt) && strcmp(BiBtxt(1),'}') & fflag ~=2
                  if y_yes == 0
                      Citation_year = '????';
                  end
                   if a_yes == 0
                      CitationAuthors = '????';
                   end
                  fflag = 2;
              end
              
          end  
        
          % remove brackets in authors name
          leftB = strfind(CitationAuthors,'{');
          rightB = strfind(CitationAuthors,'}');
          CitationAuthors([leftB,rightB]) = '';
          
          
          forsize = ['%',num2str(numel(char(Citation_key))),'s %4s %1i %',...
          num2str(numel(char(CitationAuthors))),'s\n'];   
          fid_citationIndex = fopen(CitationIndex_file,'a');
          fprintf(fid_citationIndex,forsize,char(Citation_key),...
          Citation_year, CitationAuthorsnum, char(CitationAuthors));     
          fclose(fid_citationIndex);
        end
        linenum = linenum + 1;
        waitbar(linenum/10000,h5,sprintf('%32s %5i','Indexing BiBtex entries: line = ',linenum))
        
        BiBtxt = fgetl(fid_BiB);
    end
    
    waitbar(10000/10000,h5,sprintf('%32s %5i %12s','Indexing BiBtex entries: line = ',linenum,' (completed)'))
    close(h5)
    
    fopen(fid_BiB); 
end


function newstr = InputCite(filedirBiBTex,newstr,lineNum); % Add citations

CitationIndex_file = [filedirBiBTex,'Citation_Index.txt'];

%citepfind = strfind(newstr,'\citep{'); % 1
%citetfind = strfind(newstr,'\citet{'); % 2
%citealpfind = strfind(newstr,'\citealp{'); % 3
%citeall = [citepfind,citetfind,citealpfind];
%citetype = [citepfind./citepfind,citetfind./citetfind.*2,citealpfind./citealpfind.*3];


citeTypesDB = {'\citep{','\citet{','\citealp{','\citeauthor{'};
SizeCiteCom = []; %  size of cite command
for i = 1:numel(citeTypesDB)
    SizeCiteCom = [SizeCiteCom,numel(char(citeTypesDB(i)))]; 
end
    
citeall = strfind(newstr,'\cite'); 

if ~isempty(citeall)
    citeloc = citeall(1);
    
   for i = 1:numel(citeall)
       % Get the citation keys from the text
        FindBracL_all = strfind(newstr(citeloc:end),'{'); % find end of citation
        FindBracR_all = strfind(newstr(citeloc:end),'}'); % find end of citation
        FindBracL_i = FindBracL_all(1);
        FindBracR_i = FindBracR_all(1);
        
        flag_fod = 0;
        citetype = 0;
        while flag_fod == 0
            citetype = citetype + 1;
            flag_fod = strcmp(newstr(citeloc:citeloc+FindBracL_i-1),citeTypesDB(citetype));
        end
             
        citeKeytxt_all = newstr(citeloc+FindBracL_i:citeloc+FindBracR_i-2); % find citation key
        citmulnum = strfind(citeKeytxt_all,',');
        numcite = numel(citmulnum) + 1;
        refpoints = [0,citmulnum,numel(citeKeytxt_all)+1];
        
        % find and decompose multiple citations within the same command
        if ~isempty(citmulnum)
            citeKeytxt = {};
           for j = 1:numcite
                citeKeytxt_i = citeKeytxt_all(refpoints(j)+1:refpoints(j+1)-1);
                citeKeytxt = [citeKeytxt,citeKeytxt_i];
           end
        else
           citeKeytxt = {citeKeytxt_all};
        end
        
        for j = 1:numcite  % to address the cases where there are multiple citations being called by the same citation key
        % Look for corresponding citation in DB
            Id_citationIn = fopen(CitationIndex_file,'r');
            Citationline = fgetl(Id_citationIn);

            EmptySpF = strfind(Citationline,' '); % database
             while ~isempty(EmptySpF)                
                if isempty(EmptySpF) | EmptySpF(1)==1
                    msgbox(['Citation key not found: "',char(citeKeytxt(j)),'" (line ',num2str(lineNum+1),')']);
                    break
                end             
                
                Citation_key = Citationline(1:EmptySpF(1)- 1);

                flagfound = strcmp(char(citeKeytxt(j)),Citation_key);
                 if ~flagfound
                    Citationline = fgetl(Id_citationIn);
                    EmptySpF = strfind(Citationline,' '); % database
                 else
                    break
                 end
             end  
            
            citelinestr = num2str(Citationline);
             
            if citelinestr(1:2)==num2str(-1)
                msgbox(['Citation key not found: "',char(citeKeytxt(j)),'" (line ',num2str(lineNum+1),')']);
                return
            end 
                           
            citeauth = Citationline(EmptySpF(end)+1:end);
            undscfind = strfind(citeauth,'_');
            citeauth(undscfind) = ' ';
            
            key2print = {'in preparation','submitted','under review','in press'};
            yearstend = [];
            for kf = 1: numel(key2print)
                findkey = strfind(Citationline,key2print(kf));
                if ~isempty(findkey)
                    yearstini = findkey;
                    yearstend = findkey + numel(char(key2print(kf))) - 1;
                end
            end
            if isempty(yearstend)
                yearcite = Citationline(EmptySpF(1)+1:EmptySpF(2)-1);
            else
                yearcite = Citationline(yearstini:yearstend);
            end


            remfind = strfind(yearcite,'\noop{3001}');
            yearcite(remfind:remfind+10) = '';

            if citetype == 1 & numcite == 1
                citeform = ['(',citeauth,', ',yearcite,')'];
            elseif citetype == 1 & numcite > 1
                if j == 1
                    citeform = ['(',citeauth,', ',yearcite,'; '];
                elseif j == numcite
                    citeform = [citeform, citeauth,', ',yearcite,')']; 
                else
                    citeform = [citeform, citeauth,', ',yearcite,'; ']; 
                end
            elseif citetype== 2 
                citeform = [citeauth,' (',yearcite,')'];
            elseif citetype == 3 & numcite == 1
                citeform = [citeauth,', ',yearcite];
            elseif citetype == 3 & numcite > 1
                if j == 1
                    citeform = [citeauth,', ',yearcite,'; ']; 
               elseif j == numcite
                    citeform = [citeform, citeauth,', ',yearcite]; 
               else
                    citeform = [citeform, citeauth,', ',yearcite,'; '];
                end
            else
                citeform = citeauth;
            end
        end      
        
        newstr = [newstr(1:citeloc-1),citeform, newstr(citeloc+FindBracR_i:end)];     
        citeall = strfind(newstr,'\cite'); 
        if i ~= citeall
            citeloc = citeall(1);
        end
        end
        
end 

% Prepare Eq for Word
function eqprint = PrepEquWord(eqprint)

% Commands
com_sbscpt_all = {'_','^'}; % Subscripts


% Subscripts and Superscripts
for i = 1:numel(com_sbscpt_all)
    com_sbscpt = com_sbscpt_all(i);
    lenCom = numel(char(com_sbscpt));
    comloc_all = strfind(eqprint,char(com_sbscpt));

    if ~isempty(comloc_all)
        for j = 1:numel(comloc_all)
            comloc = comloc_all(1);
            findBracR = strfind(eqprint(comloc:end),'}'); 
            if ~isempty(comloc) && ~isempty(findBracR)
                refpoint = comloc + findBracR(1);  
                eqprint = [eqprint(1:comloc), eqprint(comloc+lenCom+1:refpoint-2),'', eqprint(refpoint:end)];
            elseif ~isempty(comloc) && isempty(findBracR)
                eqprint = [eqprint(1:comloc+1),'', eqprint(comloc+2:end)];
            end
            comloc_all = strfind(eqprint,char(com_sbscpt));
            if isempty(comloc_all)
                break
            end
        end
    end
end
  

% final Clean Up function
function newstr = CleanUp(newstr)

typotypes = {'  ','. .',' .','..',' ,'}; % types of possible typos
char2keep = [1,     1,   2,    1,  2]; % character to keep

if ~isempty(newstr) && ~strcmp(newstr(1),'%')
    for i = 1:numel(typotypes)
        typotype_i = typotypes(i);
        typocor = 1;

        while typocor == 1
            
        typo = strfind(newstr,typotype_i);
            if ~isempty(typo)
                while ~isempty(typo)
                    len = numel(char(typotype_i));
                    txtbef = (1:typo(1)-1);
                    correctx = typo(1)+char2keep(i)-1;  
                    txtaft = (typo(1)+len:numel(newstr));

                    newstr = newstr([txtbef,correctx,txtaft]);
                    typo = strfind(newstr,typotype_i);
                end
                typocor = 1;
            else
                typocor = 0;
            end 
        end
    end
end
    
% Indexing Labels and References
function InedexingLabelsRef(filedir,paper_file)

LabelIndex_file = [filedir,'Labels_Index.txt'];
fid_LablIndex = fopen(LabelIndex_file,'w');
fclose(fid_LablIndex);

LabelObj = {'\section{','\subsection','\subsubsection{','\paragraph',...
    '\begin{table}','\begin{figure}','\begin{equation}',...
    '\begin{table*}','\begin{figure*}','\begin{align}'};

StopFlag = '\end{align}'; % for multiple labels within the same element

fid_paper = fopen(paper_file,'r');
textline = fgetl(fid_paper);

% initiation
section = 0;
subsection = 0;
subsubsection = 0;
paragraph = 0;
table = 0;
figure = 0;
equation = 0;

startIndFlag = 0;
labelOneyes = 0;
while ~strcmp(textline,'\end{document}') % go line by line
    
    if strcmp(textline,'\begin{document}');
        startIndFlag = 1;
    end
    
    if  startIndFlag  && ~isempty(textline) && ~strcmp(textline(1),'%')
        % Identify the last object
           
        Objtype =[];
        for i = 1:numel(LabelObj)
            Objfind = strfind(textline, LabelObj(i));
            if ~isempty(Objfind)
                Objtype = i;  % identifies the type of object
            end
        end
        
        if ~isempty(Objtype)
            % Update Object numbering and save last label for reference
            labelOneyes = 0;
            if Objtype == 1
                section = section + 1;
                RefNum = section;
                numsiz = 1;
            elseif Objtype == 2
                subsection = section + 0.1;
                RefNum = subsection;
                numsiz = 2;
            elseif Objtype == 3
                subsubsection = subsection + 0.01;
                RefNum = subsubsection;
                numsiz = 3;
            elseif Objtype == 4
                paragraph = subsubsection + 0.001;
                RefNum = paragraph;
                numsiz = 4;
            elseif Objtype == 5 | Objtype == 8
                table = table + 1;
                RefNum = table;
                numsiz = 1;
            elseif Objtype == 6 | Objtype == 9
                figure = figure + 1;
                RefNum = figure;
                numsiz = 1;
            elseif Objtype == 7 | Objtype == 10
                equation = equation + 1;
                RefNum = equation;
                numsiz = 1;
            end            
        end

        % Look for \label
        Objfind = strfind(textline, '\label{');
        if ~isempty(Objfind)
            
            if labelOneyes % for multiple labels within the same element, e.g. \aligh{}
                equation = equation + 1;
                RefNum = equation;
                numsiz = 1;
            end
            
            findBracR = strfind(textline(Objfind:end),'}');
            lablKey = textline(Objfind+7:Objfind+findBracR-2); 
            numformat = numsiz + (numsiz-1)/10;
            integerTest=~mod(numformat,1);
            numformatStr = num2str(numformat);          
            forsize = char(['%',num2str(numel(lablKey)),'s %1s %',num2str(numel(numformatStr)),'s\n']);
            fid_LablIndex = fopen(LabelIndex_file,'a');
            fprintf(fid_LablIndex,forsize,lablKey,' ',num2str(RefNum));     
            fclose(fid_LablIndex);   
            
            labelOneyes = 1;
        end
    end    
    textline = fgetl(fid_paper);
end
fclose(fid_paper)


% Add references to Doc
function newstr = InputRef(filedir,newstr)

LabelIndex_file = [filedir,'Labels_Index.txt'];

RefFind_all = strfind(newstr,'\ref{');

while ~isempty(RefFind_all)
    RefFind_i = RefFind_all(1);
    FindBracR_all = strfind(newstr(RefFind_i:end),'}');
    FindBracR_i = FindBracR_all(1);
    
    RefKeyTxt = newstr(RefFind_i+5:RefFind_i+FindBracR_i-2);
    
    % find key match in DB
    foundFlag = 0;
    fid_LablIndex = fopen(LabelIndex_file,'r');
    DB_labelKeys_alltxt = fgetl(fid_LablIndex);
    while ~foundFlag && ~strcmp(num2str(DB_labelKeys_alltxt),'-1')
        findBlankS_all = strfind(DB_labelKeys_alltxt,' ');
        findBlankSone = findBlankS_all(1);
        findBlankSlast = findBlankS_all(end);
        DB_labelKey = DB_labelKeys_alltxt(1:findBlankSone-1);  
        DB_labelKey_refnum = DB_labelKeys_alltxt(findBlankSlast+1:end);  
        foundFlag = strcmp(RefKeyTxt,DB_labelKey);   
        DB_labelKeys_alltxt = fgetl(fid_LablIndex);
    end
    fclose(fid_LablIndex);
    
    if foundFlag
        newstr = [newstr(1:RefFind_i-1),num2str(DB_labelKey_refnum),newstr(RefFind_i+FindBracR_i:end)];
    else
        newstr = [newstr(1:RefFind_i-1),'(?)',newstr(FindBracR_i+1:end)];
    end

    RefFind_all = strfind(newstr,'\ref{');
end

% convert Tex to Math style
function eqprint = ConvertMaths(newstr)

SymbLatex = {'\cdot'};
SymbWord = {'.'};

% Remove brackets
findBrack_all = sort([strfind(newstr,'{'),strfind(newstr,'}')]);
if ~isempty(findBrack_all)                 
    while ~isempty(findBrack_all)  
        findBrack = findBrack_all(1);
        if findBrack ~= numel(newstr)
            newstr = [newstr(1:findBrack-1),newstr(findBrack+1:end)];
        else
            newstr = [newstr(1:findBrack-1)];
        end                        
        findBrack_all = sort([strfind(newstr,'{'),strfind(newstr,'}')]);
    end  
end  

eqprint = newstr;

% Replace some Math Symbols
for i = 1:numel(SymbLatex)
    findSymb_all = strfind(eqprint,SymbLatex(i));
    if ~isempty(findSymb_all)                 
        while ~isempty(findSymb_all)  
            findSymb = findSymb_all(1);
            if findSymb ~= numel(eqprint)
                len = numel(char(SymbLatex(i)));
                eqprint = [eqprint(1:findSymb-1),char(SymbWord(i)),eqprint(findSymb+len:end)];
            else
                eqprint = [eqprint(1:findSymb-1),char(SymbWord(i))];
            end                        
            findSymb_all = strfind(eqprint,SymbLatex(i));
        end  
    end 
end
     

% Import BibTex to ACCESS DB: import, Citation_keys, Year, Title and Abstract
function Import_BiBtex_2_DB_ACESS(filedirBiBTex,filenameBibTex,force_import) % Prepare Citation indexes for rapid access

BiBfile = [filedirBiBTex,filenameBibTex];
CitationIndex_file = [filedirBiBTex,'Import_BibTex_2_ACCESS.txt'];

Bibdir = dir(BiBfile);
dateModBib = datenum(Bibdir.date);
Citdir = dir(CitationIndex_file);

run = 0;

if isempty(Citdir)
    run = 1;
else
    dateCit = datenum(Citdir.date);
    if dateCit < dateModBib
        run = 1;
    else
        run = 0;
    end
end


if run == 1 | force_import == 1

    fid_BiB = fopen(BiBfile,'r');
    fid_citationIndex = fopen(CitationIndex_file,'w');
    firstRowprint = 'Citation_key * Authors * Year_ * Chemical * Study_Region * Type_of_Study * Region_Classification * Cycle * Land_Use * Media * Process * In_relation_to * Title * Abstract * Flag_for_paper ';   
    forsize = ['%',num2str(numel(char(firstRowprint))),'s\n'];   
    fprintf(fid_citationIndex,forsize,char(firstRowprint));     
    fclose(fid_citationIndex);
    
    BiBtxt = fgetl(fid_BiB);
    BiBtxt = fgetl(fid_BiB);

    linenum = 1;

    h5 = waitbar(0,'Indexing BiBtex entries...');
    Citation_key = []; % citation key
    Citation_year = []; % year
    CitationAuthorsnum = []; % number of authors
    CitationAuthors = []; % authors
    Citation_title = [];% title
    Citation_abstract = []; %abstract
    
    while ~strcmp(num2str(BiBtxt),'-1')
        if ~isempty(BiBtxt) && strcmp(BiBtxt(1),'@')
          open = strfind(BiBtxt,'{') + 1 ;
          close = strfind(BiBtxt,',') - 1;
          Citation_key = BiBtxt(open:close); 
          
          y_yes = 0;
          a_yes = 0;
          t_yes = 0;
          ab_yes = 0;
          while ~strcmp(BiBtxt,'}')
              linenum = linenum + 1;
              BiBtxt = fgetl(fid_BiB);
              
              if numel(BiBtxt) > 10 && strcmp(BiBtxt(1:10),'author = {')
                  
                authstart = strfind(BiBtxt,'{');
                authend = strfind(BiBtxt,'}');
                
                authors = BiBtxt(authstart(1)+1:authend(end)-1);
                findAnd = strfind(authors,' and ');
                
                
                CitationAuthorsnum = numel(findAnd) + 1;
                commafind = strfind(authors, ', ');
                if CitationAuthorsnum == 1
                    if ~isempty(commafind)
                        CitationAuthors = authors(1:commafind(1)-1);
                    else
                        CitationAuthors = authors;
                    end
                elseif CitationAuthorsnum == 2
                    if numel(commafind == 1)
                        commafind = [commafind, numel(authors) + 1]; 
                    end
                    emptyspaces = strfind(authors,' ');
                    CitationAuthors = [authors(1:commafind(1)-1),'_and_',authors(findAnd(1)+ 5:commafind(2)-1)];
                elseif CitationAuthorsnum > 2 
                    CitationAuthors = [authors(1:commafind(1)-1),'_et_al.'];
                end              
                  a_yes = 1;
              elseif numel(BiBtxt) > 8 && strcmp(BiBtxt(1:8),'year = {')
                  ystart = strfind(BiBtxt,'{');
                  yend = strfind(BiBtxt,'}');           
                  Citation_year = BiBtxt(ystart+1:yend-1); 
                  y_yes = 1;
              elseif numel(BiBtxt) > 8 && strcmp(BiBtxt(1:9),'title = {')
                  tstart = strfind(BiBtxt,'{');
                  tend = strfind(BiBtxt,'}');           
                  Citation_title = BiBtxt(tstart(1)+2:tend(end)-2); 
                  t_yes = 1;
              elseif numel(BiBtxt) > 8 && strcmp(BiBtxt(1:12),'abstract = {')
                  tstart = strfind(BiBtxt,'{');
                  tend = strfind(BiBtxt,'}');           
                  Citation_abstract = BiBtxt(tstart(1)+1:tend(end)-1); 
                  ab_yes = 1;
              end
                           
          end  
        
          % Put "????" if not found
          if y_yes == 0
              Citation_year = '????';
          end
           if a_yes == 0
              CitationAuthors = '????';
           end
           if t_yes == 0
              Citation_title = '????';
           end
           if ab_yes == 0
              Citation_abstract = '????';
           end

          
          % remove brackets in authors name
          leftB = strfind(CitationAuthors,'{');
          rightB = strfind(CitationAuthors,'}');
          CitationAuthors([leftB,rightB]) = '';
          
          
          forsize = ['%',num2str(numel(char(Citation_key))),'s %1s %',...
              num2str(numel(char(CitationAuthors))),'s %1s %4s %19s %',...
              num2str(numel(char(Citation_title))),'s %1s %',...
              num2str(numel(char(Citation_abstract))),'s %1s \n'];   
          
          fid_citationIndex = fopen(CitationIndex_file,'a');
          fprintf(fid_citationIndex,forsize,char(Citation_key),'*',...
              char(CitationAuthors),'*',...
              Citation_year,'* * * * * * * * * *',...
              char(Citation_title),'*',...
              char(Citation_abstract),'*');  
          
          fclose(fid_citationIndex);
        end
        linenum = linenum + 1;
        waitbar(linenum/10000,h5,sprintf('%32s %5i','Indexing BiBtex entries: line = ',linenum))
        
        BiBtxt = fgetl(fid_BiB);
    end
    
    waitbar(10000/10000,h5,sprintf('%32s %5i %12s','Indexing BiBtex entries: line = ',linenum,' (completed)'))

    fopen(fid_BiB);
    
end
close(h5)

% Get default directory for tex file search
function defaultdir = GetDefaultdir(a)
%[~, CompName] = system('hostname');

if a==1
    defaultdir = '/media/DATADRIVE1';
   % if char(CompName(1:end-1))=='ARTS-GEOG-2725'
   %     defaultdir = 'E:\OneDrive\DI_PRF_CUR\UofS\9_Journal_Publications\0_MATLAB_support\test';
    %elseif char(CompName(1:end-1))=='DIOGOCOSTA'
   %     defaultdir = 'C:\Users\diogo\OneDrive\DI_PRF_CUR\Diogo_Sync_Tablet';
   % else
   %     defaultdir = 'C:\';
   % end
elseif a==2
    defaultdir = '/home/diogo/Mendeley_DB'
   %  if char(CompName(1:end-1))=='ARTS-GEOG-2725'
   %     defaultdir = 'E:\OneDrive\DI_PRF_CUR\Mendeley_library\BibTex';
   % elseif char(CompName(1:end-1))=='DIOGOCOSTA'
   %     defaultdir = 'C:\Users\diogo\OneDrive\DI_PRF_CUR\Mendeley_library\BibTex';
   % else
   %     defaultdir = 'C:\';
   % end
end

% Convert Word to Latex
function Gen_Word2Latex(filenameWORD,filedir,filedirBiBTex)

% Open a paragraph sample from WORD doc
word = actxserver('Word.Application');
wdoc_paragraph = word.Documents.Open([pwd,'\ParagraphSample.docx']);
word_text_paragraph = wdoc_paragraph.Content.Text;
word.Quit();

filepathWORD = [filedir,filenameWORD];
dot_loc = strfind(filepathWORD,'.');
filepath_Latex = [filepathWORD(1:dot_loc-1),'.tex'];
filepath_Latex_NEW = [filepathWORD(1:dot_loc-1),'_NEW.tex'];

% Open WORD doc
word = actxserver('Word.Application');
wdoc = word.Documents.Open(filepathWORD);
%wdoc is the Document object which you can query and navigate.
word_text_raw = wdoc.Content.Text;
word.Quit();

% Organise the word file
% parag_int = strfind(word_text_raw,'?'); 
% word_raw_i = word_text_raw(1:parag_int-1);
% word_text_org = {word_raw_i};
% for i =1:numel(parag_int)-1
%     word_raw_i = word_text_raw(parag_int(i)+5:parag_int(i+1));
%     word_text_org = [word_text_org;{word_raw_i}];
% end

word_text_org = splitlines(word_text_raw);

% Open the Old Latex file and look for objects (to be used as reference)
Latexfile_id = fopen(filepath_Latex,'r');
lookforSections = {'\begin{abstract}','\section{','\subsection{','\subsubsection{','\paragraph{'};
StartDocMarker = lookforSections(1);
EndDocMarker = '\section{Acknowledgements}';
%LookforOtherObjects = {'\begin{equation}','\begin{enumerate}','\begin{figure}'};
findSectionline = [];
findSection_loc = [];
findSectiontype = [];
findSectionname = {};
fullSectionnameCommand = {};
StartDocMarker_i = [];
EndDocMarker_i = [];
line_i = 0;
while ~feof(Latexfile_id)
    line_i = line_i+1;
    lineread = fgetl(Latexfile_id);
    % Find start of doc
    findStartDocMarker = strfind(lineread,char(StartDocMarker));
    if ~isempty(findStartDocMarker)
        StartDocMarker_i = line_i;
    end
    % Find End of doc
     findSEndDocMarker = strfind(lineread,char(EndDocMarker));
    if ~isempty(findSEndDocMarker)
        EndDocMarker_i = line_i;
    end
    % Look for Sections
    for i = 1:numel(lookforSections)
        findSection_loc_i = strfind(lineread,char(lookforSections(i)));
        if ~isempty(findSection_loc_i)
            % check for comments
            findComment = strfind(lineread(1:findSection_loc_i),'%');
            if ~isempty(findComment)
                continue
            end
            % store info for sections
            findSectionline = [findSectionline,line_i];
            findSection_loc = [findSection_loc, findSection_loc_i];
            findSectiontype = [findSectiontype,i];
            findSectionname_i = GetSectioName(lineread,findSection_loc);
            findSectionname = [findSectionname;findSectionname_i];
            fullSectionnameCommand_i = [char(lookforSections(i)),char(findSectionname_i),'}'];
            findbrakRight = strfind(fullSectionnameCommand_i,'}');
            fullSectionnameCommand_i = fullSectionnameCommand_i(1:findbrakRight(1));
            fullSectionnameCommand = [fullSectionnameCommand;char(fullSectionnameCommand_i)];
        end
    end 
end
fclose(Latexfile_id)

line_total_latex = line_i;

% Generate the new Latex file

Latxfile_NEW_id = fopen(filepath_Latex_NEW,'w+');
Latexfile_id = fopen(filepath_Latex,'r');
line_i = 0;
flag_par = 0;

h = waitbar(0,'Generating the Latex file...');
while ~feof(Latexfile_id)
   line_i = line_i+1;
   lineread = fgetl(Latexfile_id);
   % copy old latex to new latex until 'StartDocMarker_i'
   if line_i < StartDocMarker_i
       CallPrint({lineread},Latxfile_NEW_id,flag_par)      
   elseif line_i > EndDocMarker_i - 1
       CallPrint({lineread},Latxfile_NEW_id,flag_par)        
   else
      findnewSec = findIfSec(lineread,fullSectionnameCommand); 
      if ~isempty(findnewSec)
          SecStr = fullSectionnameCommand(findnewSec);
          CallPrint({char(SecStr)},Latxfile_NEW_id,flag_par)
          % Get the Section text from WORD
          SectionBodyText = FindSecTextinWord(word_text_org,findnewSec,fullSectionnameCommand,word_text_paragraph);
          flag_par = 1;
          % Correct simbols, e.g. %
          SectionBodyText = CorrectSymbols(SectionBodyText);
          % Add citation comands and link to database
          SectionBodyText = addCitationComands(filedirBiBTex,SectionBodyText);
          % Print
          CallPrint(SectionBodyText,Latxfile_NEW_id,flag_par)  
          flag_par = 0;
          
          % TO do now:
            % 1) References
            % 2) Equations
            % 3) ref for figures and tables
            
         if strcmp(SecStr,'\begin{abstract}')
              CallPrint({'\end{abstract}'},Latxfile_NEW_id,flag_par)
         end
            
      end
   end
   waitbar(line_i / line_total_latex)
end
close(h) 
fclose(Latxfile_NEW_id)
fclose(Latexfile_id)

% Find Section
function findSectionname_i = GetSectioName(lineread,findSection_loc);

LocBrakLeft = strfind(lineread(findSection_loc:end),'{');
LocBrakRigh = strfind(lineread(findSection_loc:end),'}');
findSectionname_i = lineread(LocBrakLeft+1:LocBrakRigh-1);

if strfind(findSectionname_i,'abstract')
    findSectionname_i = 'Abstract';
end

% Find if section (in Latex file)
function findnewSec = findIfSec(lineread,fullSectionnameCommand)

findnewSec = [];
for i = 1:numel(fullSectionnameCommand)
    Secexist = strfind(lineread,char(fullSectionnameCommand(i)));
    if ~isempty(Secexist)
        findnewSec = i;
        return
    end
end

% Print to txt
function CallPrint(newline,file_id,flag_par)
for i=1:numel(newline)
    newline_2print = num2str(char(newline(i)));
    if flag_par
        strstyle = 's';
    else
        strstyle = 's\n';
    end
    formatSpec = ['%',num2str(numel(newline_2print)),strstyle];
    fprintf(file_id,formatSpec,newline_2print);
    if flag_par
        formatSpec = '%4s\n';
        fprintf(file_id,formatSpec,'\par');
        fprintf(file_id,formatSpec,' ');
    end
end

    
% Find Section Text in Word
function SectionBodyText = FindSecTextinWord(word_text_org,findnewSec,fullSectionnameCommand,word_text_paragraph)

% Get the name of current and following sections
secloc = findnewSec;
SectionNames = {};
for s=1:2
    LeftBrak_i = strfind(char(fullSectionnameCommand(secloc)),'{');
    RightBrak_i = strfind(char(fullSectionnameCommand(secloc)),'}');
    SectionNamesLatex = char(fullSectionnameCommand(secloc));
    SectionNames_i = SectionNamesLatex(LeftBrak_i+1:RightBrak_i-1);
    % check if introduction. If yes, then capitalize because that has
    % been done for the word version
    if contains(SectionNames_i,'abstract')
        SectionNames_i = 'ABSTRACT';
    end
    SectionNames = [SectionNames;SectionNames_i];
    secloc = secloc + 1;
end

% Identigy location of section body text in cell 
SectionBodyText = {};
LocSecStart = [];
LocSecEnd = [];
multipleline = 0;
for i=1:numel(word_text_org)

    WordLine = char(word_text_org(i));
    %if isempty(LocSecStart)
        LocSecStart = [i, strfind(WordLine,char(SectionNames(1)))]; % line (row), position in line (col)
        SectionNames
        LocSecEnd = [i, strfind(WordLine,char(SectionNames(2)))];
   % end
    
    % If found
    if numel(LocSecStart) == 2 & numel(LocSecEnd) == 2 & ~multipleline
        TextStart = LocSecStart(2) + numel(char(SectionNames(1)));
        TextEnd = LocSecEnd(2) - numel(char(SectionNames(2)));
        SectionBodyText = {WordLine(TextStart:TextEnd)};
        return
    elseif numel(LocSecStart) == 2 & numel(LocSecEnd) == 1
        multipleline = 1;
        SectionBodyText = {};
        TextStart = LocSecStart(2) + numel(char(SectionNames(1)));
        addline = checkifparagrapharrows(WordLine(TextStart:end),word_text_paragraph); % remove paragraphs
        SectionBodyText  = [SectionBodyText; {addline}];
        continue
    elseif numel(LocSecStart) == 1 & numel(LocSecEnd) == 1 & multipleline
        addline = checkifparagrapharrows(WordLine,word_text_paragraph);
        SectionBodyText = [SectionBodyText;{addline}];
        continue
    elseif numel(LocSecStart) == 1 & numel(LocSecEnd) == 2 & multipleline 
        TextEnd = LocSecEnd(2) - 2;
        addline = checkifparagrapharrows(WordLine(1:TextEnd),word_text_paragraph);
        SectionBodyText = [SectionBodyText;{addline}];
        return
    end

end

% Check if paragraph(s) and remove it(them)
function addline = checkifparagrapharrows(WordLinePortion,word_text_paragraph);
   if ~isempty(WordLinePortion)
        paragraphLoc = strfind(WordLinePortion,word_text_paragraph);
        charNuArray = [1:1:numel(WordLinePortion)];
        charNuArray(paragraphLoc) = [];
        addline = WordLinePortion(charNuArray); 
   else
       addline = WordLinePortion;
   end
 
% Correct symbols
function SectionBodyText = CorrectSymbols(SectionBodyText);

symbols = {'%','#','^','\ud'};
for i= 1:numel(SectionBodyText)
    flag_change = 0;
    lineread = char(SectionBodyText(i));
    if ~isempty(lineread)
        for s = 1:numel(symbols)
           symbol_i = symbols(s);
           symbol_loc = strfind(lineread,symbol_i); 
           if ~isempty(symbol_loc)
               for r = 1:numel(symbol_loc)
                   if lineread(symbol_loc(r)-1)~= '\'
                    flag_change = 1;
                    lineread_change = [lineread(1:symbol_loc(r)-1),...
                                '\', lineread(symbol_loc(r)),...
                                lineread(symbol_loc(r)+1:end)];
                   end              
               end
           end
        end
    end
    if flag_change == 1
        lineread = lineread_change;
    end
    SectionBodyText(i) = {lineread};
end

% Add citation comands
function SectionBodyText_wCite = addCitationComands(filedirBiBTex,SectionBodyText);
CitationIndex_file = [filedirBiBTex,'Citation_Index.txt'];
SectionBodyText_wCite = {};
citecomands = {'\citet{','\citep{','\citealt{'};
for i=1:numel(SectionBodyText)
   lineread_txt =  char(SectionBodyText(i));
   if ~isempty(lineread_txt)
       Citation_indexfile_id = fopen(CitationIndex_file);
       while ~feof(Citation_indexfile_id)
           lineread_citDB = fgetl(Citation_indexfile_id);
           lineread_citDB_name = GetCoreCitName(lineread_citDB);
           citationkey = GetCitationKey(lineread_citDB); 
           findCite = strfind(lineread_txt,lineread_citDB_name);
           if ~isempty(findCite)
             for c1 = 1:numel(findCite)
                   findCite = strfind(lineread_txt,lineread_citDB_name);
                   c =1;
                   %findCite = strfind(lineread_txt,lineread_citDB_name);
                   % 1st: check if falls in single \citep or \citet structure
                   CitFullpossTxt = ConstCitepAndCitet(lineread_citDB);
                   citetype = CheckMatchwCitType(lineread_txt,CitFullpossTxt,findCite,c); %1) citep, 2) citet, 3) citealt with e.g.
                   citationtype = find(citetype == 1);
                   
                   if ~isempty(citationtype)
                       g = 0;
                       if citationtype == 2
                           g = 1;
                       end
                       textbefore = lineread_txt(1:findCite(c)-1-g);
                       citationcomand = char(citecomands(citationtype));
                       sizecomand = numel(char(CitFullpossTxt(citationtype)));
                       if citationtype == 2
                           sizecomand = sizecomand-1;
                       end
                       textafter = lineread_txt(findCite(c)+ sizecomand:end);
                       lineread_txt = [textbefore,citationcomand,...
                                    citationkey,'}',...
                                    textafter];
                   end

                   % continue here... 
                   % 1st: it can't find many of the citations
                   % 2nd: \citealt with multiple citations is not working
                   % well because there is a e.g.
                   % 2nd: check for multiple citet and citep
                   % 3rd: check for e.g. and possibly multiple citet and citep
               end
           end
       end
   end
   SectionBodyText_wCite = [SectionBodyText_wCite;lineread_txt];
end

% Get main text of citation_index line
function lineread_citDB_name = GetCoreCitName(lineread_citDB)

findSpaces = strfind(lineread_citDB,' ');
lineread_citDB_name = lineread_citDB(findSpaces(end)+1:end);
find_underscore = strfind(lineread_citDB_name,'_');
lineread_citDB_name(find_underscore) = ' ';


% Construct \citet and \citep references
function CitFullpossTxt = ConstCitepAndCitet(lineread_citDB)

findblanks = strfind(lineread_citDB,' ');
citet_txt = [lineread_citDB(findblanks(end)+1:end),' (',lineread_citDB(findblanks(1)+1:findblanks(2)-1),')'];
citep_txt = ['(',lineread_citDB(findblanks(end)+1:end),', ',lineread_citDB(findblanks(1)+1:findblanks(2)-1),')'];
citealt_txt = [lineread_citDB(findblanks(end)+1:end),', ',lineread_citDB(findblanks(1)+1:findblanks(2)-1)];

citet_txt(strfind(citet_txt,'_')) = ' ';
citep_txt(strfind(citep_txt,'_')) = ' ';
citealt_txt(strfind(citealt_txt,'_')) = ' ';

CitFullpossTxt = {citet_txt,citep_txt,citealt_txt};


% Get citation key
function citationkey = GetCitationKey(lineread_citDB) 

BlankSpace = findstr(lineread_citDB,' ');
citationkey = lineread_citDB(1:BlankSpace(1)-1);

% look for citation in the word text
function citetype = CheckMatchwCitType(lineread_txt,CitFullpossTxt,findCite,c)
citetype = [0 0 0]; % 
if numel(findCite)>1 & c~= numel(findCite)
   lineread_txt_look = lineread_txt(findCite(c):findCite(c+1));
elseif numel(findCite)>1 & c== numel(findCite)
   lineread_txt_look = lineread_txt(findCite(c):end);
else
    lineread_txt_look = lineread_txt;
end
citetype(1) = ~isempty(strfind(lineread_txt_look,char(CitFullpossTxt(1))));
citetype(2) = ~isempty(strfind(lineread_txt_look,char(CitFullpossTxt(2))));
if citetype(2)
    citetype(3) = 0;
else
    citetype(3) = ~isempty(strfind(lineread_txt_look,char(CitFullpossTxt(3))));
end
