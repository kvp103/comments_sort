#function to separate out the header from the body

create_header <- function(raw_emails_df) {
  raw_emails_df %<>% 
    separate(data=.,col = email_text, into=c(NA, "rest"), sep="subject: ", remove=TRUE) %>%
    separate(data=.,col = rest, into=c("subject", "rest"), sep="\\nfrom: ", remove=TRUE) %>%
    separate(data=.,col = rest, into=c("from", "rest"), sep="\\nsent: ", remove=TRUE) %>%
    separate(data=.,col = rest, into=c("date", "body"), sep="\\n", remove=TRUE, extra = "merge")
  
  # NB:
  # Local time sometimes GMT +10, GMT +11 or other so can't just use '+1000' as a delimiter for date column. 
  # Look for first new line instead.
  #"(?<=.)(?=[sent: ])"
  # Don't need to get rid of "dear minister" etc because not any topic analysis yet
  
  return(raw_emails_df)
}