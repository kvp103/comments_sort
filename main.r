################
################
#### Header ####
################
################

# Purpose of code: To sort a single directory of text files into subdirectories, based on their topics.
# Output: file movement, record file (csv) and a log file (csv). 
# Input: One excel sheet, with columns for each expected topic/project and a list of keywords.
# (How to execute code): Code will be scheduled to run automatically at minimum daily intervals.
#   Check that the data > outlook_dump folder 
#   is populated with .msg files, and that the input spreadsheet is up to date. 
#
# Author: Kirsten Vo-Phuoc
# 
# Maintainer:
#
# Date created:
# Date last major modification:


###############
## Libraries ##
###############

# In order of usage:
library(rstudioapi)
library(magrittr)
library(reticulate)
library(readtext)
library(tidyverse)
library(lubridate)
library(filesstrings) #for file.move. file.[other] is in base.
library(xlsx) #for read.xslx. Can also use read.table in base.

###############
## Variables ##
###############

current_wd <- getActiveProject()
current_wd %>% setwd()
function_wd <- "code/functions"
msg_wd <- paste(current_wd, "data/msg_txt_test", sep = '/')
#msg_raw_wd <- paste(current_wd, 'data/msg_raw_test', sep = '/')
output_wd <- paste(current_wd, "output", sep = '/')
code_wd <- "code"
data_wd <- paste(current_wd, "data", sep = "/")
input_spreadsheet <- "categories_test.xlsx"

###############
## Functions ##
###############

# maybe remove paste jobs and just go into the folder? 
# maybe make a list of function names, then apply function with pasting
# Note any function used within an "apply" statement is explicitly written in 
# code below
source(paste(function_wd, 'fn-create_email_df.r', sep= '/'))
source(paste(function_wd, 'fn-create_header.r', sep= '/'))
source(paste(function_wd, 'fn-clean_date.r', sep= '/'))
source(paste(function_wd, 'fn-separate_name_and_email.r', sep= '/'))
source(paste(function_wd, 'fn-camel_case_names.r', sep= '/'))
source(paste(function_wd, 'fn-separate_sender_name.r', sep= '/'))
source(paste(function_wd, 'fn-remove_white_space.r', sep= '/'))
source(paste(function_wd, 'fn-create_partial_tag.r', sep= '/'))
source(paste(function_wd, 'fn-read_spreadsheet.r', sep= '/'))
source(paste(function_wd, 'fn-clean_project_names.r', sep= '/'))
source(paste(function_wd, 'fn-create_folders.r', sep= '/'))
source(paste(function_wd, 'fn-compile_keywords.r', sep= '/'))
source(paste(function_wd, 'fn-search_keywords.r', sep= '/'))
source(paste(function_wd, 'fn-move_unknowns.r', sep= '/'))
source(paste(function_wd, 'fn-move_knowns.r', sep= '/'))
source(paste(function_wd, 'fn-attach_project_id.r', sep= '/'))
source(paste(function_wd, 'fn-write_output_file.r', sep= '/'))

##############
##############
#### Main ####
##############
##############

start_time = Sys.time()

##########################
## Convert .msg to .txt ##
##########################

#py_run_string("directory = ")
py_run_file(paste(current_wd, code_wd, "msgtotext.py", sep='/'))

#################################
## Read and store saved emails ##
#################################

# Store public comments from .txt files into data frame
raw_emails_df <- create_email_df(msg_wd)

#####################
## Cleaning galore ##
#####################

# Split out email header (sender, date, subject) from body
raw_emails_df %<>% create_header()

# Date manipulation - currently in dmy hms but in character format. 
raw_emails_df$date %<>% clean_date()

# From manipulation - split out sender name from their emails
raw_emails_df %<>% separate_name_and_email()

# Makes all the names camel case
raw_emails_df$from %<>% camelcase_names()

# Split the name column into two columns for firstname and lastname.
# People with just one name will have a blank as lastname.
raw_emails_df %<>% separate_sender_name()

# Strip any remaining leading and trailing white space for both first/lastnames
raw_emails_df[c("firstname", "lastname")] %<>% apply(2, remove_white_space)

# create a column with a partial name & date tag in the format:
# "Referral-Email-Comment-<Lastname><FirstInitial>-<YYYYMMDD>"
# to help with file renaming conventions.
raw_emails_df %<>% create_partial_tag()


########################################
## Create appropriate project folders ##
########################################

# Read in EPBC Projects (open comment period) spreadsheet
# check file extension in variables up top
epbc_projects <- read_spreadsheet(data_wd, input_spreadsheet)

# Clean EPBC project names
names(epbc_projects) %<>% sapply(clean_project_names)

# Make folders for each EPBC project.
# The number of folders created is assign to an R object as a record for the output log file. 
num_folders <- create_folders(epbc_projects, output_wd)


#####################################
## Search for key words of project ##
#####################################

# Make an array with each element being a string of project keywords for each EPBC project
project_keywords <- apply(epbc_projects, 2, compile_keywords)
  
# Create empty project_flags columns in data frame to be populated once comments are sorted into projects
raw_emails_df[c(names(project_keywords))] <- NA

# For each comment, search the EPBC keywords within the subject first
raw_emails_df[c(names(project_keywords))] <- sapply(project_keywords, search_keywords, y = raw_emails_df$subject) 

# Check which comments didn't get assigned a flag from the subject search
# (flag_sums == 0)
flag_sums <- raw_emails_df %>% select(all_of(names(project_keywords))) %>% rowSums()

# From the unflagged comments, search over the body/text to assign flag
raw_emails_df[c(which(flag_sums == 0)), c(names(project_keywords))] <- sapply(project_keywords, search_keywords, y = raw_emails_df$body[c(which(flag_sums == 0))])

# assign unknown flag to ones that still have no flags
flag_sums <- raw_emails_df %>% select(all_of(names(project_keywords))) %>% rowSums()
raw_emails_df$unknown <- 0
raw_emails_df$unknown[c(which(flag_sums == 0))] <- 1
raw_emails_df$unknown[c(which(flag_sums >= 2))] <- 1
  # could clean these lines up a bit


################################
## Allocate emails to folders ##
################################

# use file.copy instead of file.move for testing purposes. 

# first move unknown comments
files_df <- raw_emails_df %>% select(filename, partial_tag, unknown)
  # select all from the data frame with unknown flag
apply(files_df, 1, move_unknowns) %>% invisible()

# now move comments that were assigned a single project flag
files_df <- raw_emails_df %>% 
  subset(unknown == 0) %>% 
  select(filename, partial_tag, all_of(names(project_keywords)))
  # select all from the data frame that do not have unknwn flag
apply(files_df, 1, move_knowns) %>% invisible()


##############################
## Create Project ID column ##
##############################

# assign a project_id column to data frame to be populated
raw_emails_df$project_id <- NA
raw_emails_df$project_id <- raw_emails_df %>% apply(1, attach_project_id, y = names(project_keywords))


#######################
#######################
#### Write outputs ####
#######################
#######################

end_time <- Sys.time()
message('Execution time: ', end_time-start_time)

#####################
## File record CSV ##
#####################

# Create a record file with each row writing meta data around the comment 
# (so does not include body of comment)
output_df <- raw_emails_df[c("date", "firstname", "lastname", "project_id")] %>% 
  cbind("date_processed" = Sys.time()) %>%
  rename("date_received" = "date")

write_output_file(output_wd, "record_file.csv", output_df)


#####################
## Log file CSV    ##
#####################

# Create log file with each row containing meta data of how this script ran
log_df <- list(
  Sys.time(),
  end_time - start_time,
  nrow(output_df),
  ncol(epbc_projects),
  length(num_folders),
  nrow(output_df) - sum(raw_emails_df$unknown),
  sum(raw_emails_df$unknown)
) %>% as.data.frame()

names(log_df) <- c(
  "timestamp", 
  "execution_time", 
  "comments_processed", 
  "epbc_projects_processed",
  "epbc_folders_created", 
  "comments_allocated", 
  "comments_unknown"
)

write_output_file(output_wd, "log_file.csv", log_df)


