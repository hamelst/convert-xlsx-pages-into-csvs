# Function is verbatim from this stack overflow answer: https://stackoverflow.com/questions/12945687/read-all-worksheets-in-an-excel-workbook-into-an-r-list-with-data-frames
# Credit to Jeremy Anglim, senior lecturer at Deakin (https://stackoverflow.com/users/180892/jeromy-anglim)

if (!require('readxl')) install.packages('readxl'); library('readxl') # comment out and uncomment below if errors are thrown
# library('readxl') 

read_excel_allsheets <- function(filename, tibble = FALSE) {
  # I prefer straight data.frames
  # but if you like tidyverse tibbles (the default with read_excel)
  # then just pass tibble = TRUE
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}