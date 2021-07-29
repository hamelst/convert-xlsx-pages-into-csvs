source('read_excel_allsheets.R')


# excel file we will use
xlxs_file = './data/input/40401A-cyto-test--cyt_1--nms_0.1--prob_0.1--_firstimage--P1_PD-1_40401A.xlsx'
# where to output csv files
output_folder = './data/output'
# what to separate original filename and page number (example in loop below)
sep = '--page_'
# uncomment below if you wanna create the output folder automatically
dir.create(output_folder, showWarnings = T, recursive = T)

#read excel object
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
  #     and they will all be save to the variable output_folder
  
  csv_name = paste0(basename(xlxs_file),sep,page_num,'.csv')
  
  # write csv to output folder
  write.csv(x = xlxs_page, file = file.path(output_folder,csv_name))
}