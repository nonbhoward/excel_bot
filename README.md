# excel_bot

just a foundation, i haven't needed an excel bot in a long time

all it does right now is..

### extract
1. parse files, workbooks, and worksheets into dicts and such

### compare
2. look..  
      ..over a certain range..  
      ..with workbooks containing certain names..  
      ..with worksheets containing certain names..  
      ..for the provided search terms..  

### list
3. keeps a list/dict/etc (depending on data) of every cell address that matches all of the above  

## not much else.
just a foundation for next time i need to do something with excel

# quick-start
1. have python installed
2. have pip installed (for your python version)  
3. clone the repo  
4. make a virtual environment  
5. install dependencies per requirements.txt  
6. put some excel files in /project_root/data_src/
7. set custom terms
8. run!

## set custom terms
# customize your filters (marked with FIXME tag as of last-update)
1. filter 1, customize search area with main_loop > get_search_area  
2. filter 2, customize search terms to find with main_loop > get_search_terms_to_find  
3. filter 3, customize workbook keywords of interest with main_loop > get_workbook_keywords_of_interest  
4. filter 4, customize worksheet keywords of interest with main_loop > get_worksheet_keywords_of_interest  
