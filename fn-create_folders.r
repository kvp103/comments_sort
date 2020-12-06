# Function to create output folders to sort the public comments.
# Each EPBC project has one folder. 
# The number of folders to create depends on the input spreadsheet.
# There will always be an "unknown" folder created.
create_folders <- function(epbc_projects, output_wd){
  unknown_folder <- "unknown"
  folder_name_list <- c(unknown_folder, names(epbc_projects))
  directory_create_list <- paste(output_wd, folder_name_list, sep='/')
  directory_create_list %>% 
    # short apply statement to just message whether the epbc folder already exists or not.
    # when run from batch file, this will not output to terminal. 
    sapply(function(x) {
      if(file.exists(x)) {
        message('The directory: "', x, '" already exists')
      } else {
        dir.create(x)
        message('The directory: "', x, '" was created')
      }
    }) %>%
    invisible()
  return(folder_name_list)
}