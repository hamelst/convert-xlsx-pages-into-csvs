# loading the read_excel_allsheets function
source('read_excel_allsheets.R')

#### SETUP #### 
# edit variables here

#folder of xlsx files
xlxs_dir = './data/input'
# where to output csv files
output_folder = './data/output'
# what to separate original filename and page number (example in main loop below)
sep = '--page_'
# file extension we are looking for
file_ext = '.xlsx'
# uncomment below if you wanna create the output folder automatically
dir.create(output_folder,showWarnings = F, recursive = T)

#### END SETUP #### 

### MAIN ###
# don't edit

xlsx_files = grep(list.files(xlxs_dir, full.names = T), value = T, pattern = paste0(".*\\",file_ext,"$"))

for (xlxs_file in xlsx_files){
  xlxs_object = read_excel_allsheets(xlxs_file)
  
  page_num = 0
  for (xlxs_page in xlxs_object) {
    page_num = page_num + 1
    # every csv is the filename of the xls + a separation character and then page number 
    # Example:
    #     if the variable sep = '--page_'
    #         and the original excel file we used is named 'original_excel_file.xlsx'
    #         and the excel file has 3 pages
    #     Our csv files will be named: 
    #         original_excel_file.xlsx--page_1.csv
    #         original_excel_file.xlsx--page_2.csv
    #         original_excel_file.xlsx--page_3.csv
    #     and they will all be save to the variable `output_folder`
    
    csv_name = paste0(basename(xlxs_file),sep,page_num,'.csv')
    
    # write csv to output folder
    write.csv(x = xlxs_page, file = file.path(output_folder,csv_name))
  }
}

### END MAIN ###