# Import necessary libraries
library(XML)
library(tidyverse)
library(openxlsx)

# Set working directory
setwd('C:/Users/LynchZ20/Desktop/Misc/BG_CW')

# set the url to the boston globe crossword
url <- 'https://www3.bostonglobe.com/lifestyle/crossword'

# Read in html
html <- readLines(url)

# EXTRACT ALL CLUES

# First find ordered list of clue labels
lbl <- unlist(str_extract_all(html, "<label for=\"input-[0-9]*\">"))
lbl <- gsub("<label for=\"input-", "", lbl)
lbl <- as.integer(gsub("\">", "", lbl))

# Next find list of clues themselves in same order
clue <- html[grep("<b>", html, fixed=T)+1]
clue <- str_replace_all(clue, fixed('&quot;'), '"')
clue <- str_replace_all(clue, fixed('&amp;'), '&')
clue <- clue[1:(length(clue)-1)]

# Next create repeated vector of 'across' and 'down' accordingly
dir <- lbl > lag(lbl) | is.na(lag(lbl))
dir[1:(which(dir==F)-1)] <- 'across'
dir[dir != 'across'] <- 'down'

# Lastly put all 3 vectors into table format
clues <- tibble(
  lbl = lbl, 
  clue = clue, 
  dir = dir
)
rm(lbl, clue, dir)

# Extract puzzle title
title <- html[grep('<p class=\"subhed\">', html)+1]

# CREATE TABLE OF PUZZLE

# First create table with corresponding label numbers in correct squares
puzzle <- as_tibble(readHTMLTable(html)[[1]]) %>% 
  mutate_all(as.character)

# Next identify the coordinates of all the blacked-out squares in the puzzle
blocks <- str_extract(html[grep('data-coords', html, fixed=T)], '[0-9]*,[0-9]*')
blocks <- tibble(blocks) %>% 
  rename(coord = blocks) %>% 
  separate(col = coord, into = c('y', 'x'), sep = ",") %>% 
  mutate(x = as.integer(x), 
         y = as.integer(y))
blocks <- anti_join(tibble(x = sort(rep(1:nrow(puzzle), nrow(puzzle))), 
                           y = rep(1:nrow(puzzle), nrow(puzzle))), blocks, 
                    by=c('x', 'y'))

# Next loop through each coordinate pair and place an "X" in the puzzle table 
# where the square will be blacked out
for(i in 1:nrow(blocks)){
  puzzle[blocks$x[i], blocks$y[i]] <- "X"
}

# Lastly include an index field in the puzzle table
puzzle <- puzzle %>% 
  add_column(ind = 1:nrow(puzzle), .before = "V1")
rm(i)

# EXTRACT CLUE ANSWERS

# Extract answers

# First find each letter and its respective coordinates
hidden <- anti_join(tibble(x = sort(rep(1:nrow(puzzle), nrow(puzzle))), 
                         y = rep(1:nrow(puzzle), nrow(puzzle))), blocks, 
                     by=c('x', 'y')) %>% 
  mutate(letter = unlist(str_extract_all(html[grepl("maxlength=", html, fixed=T)], 
                                         'name=\"[A-Z]\"'))) %>% 
  mutate(letter = unlist(str_extract_all(letter, '[A-Z]')))

# Next create an empty table
answers <- matrix(NA, nrow=nrow(puzzle), ncol=nrow(puzzle))

# Then fill in the table with the correct letter for the respective square
for(i in 1:nrow(hidden)){
  answers[hidden$x[i], 
          hidden$y[i]] <- hidden$letter[i]
}

# Create tibble of answers
colnames(answers) <- paste0('V', 1:15)
answers <- as_tibble(answers) %>% 
  add_column(ind = 1:nrow(puzzle), .before = "V1")
rm(i, hidden)

# UPDATE EXCEL WORKBOOK

# Open workbook
wb <- loadWorkbook('BG_CW.xlsx')

# Replace 'Raw' tab with new puzzle data
removeWorksheet(wb, 'Raw')
addWorksheet(wb, 'Raw')
writeData(wb, 'Raw', puzzle)

# Replace 'Across' tab with new across clues
removeWorksheet(wb, 'Across')
addWorksheet(wb, 'Across')
writeData(wb, 'Across', (clues %>% 
                           filter(dir=='across') %>% 
                           add_column(ind=1:sum(clues$dir == 'across'),.before='lbl') %>% 
                           unite(text, lbl, clue, sep=". ", remove=F)))

# Replace 'Down' tab with new down clues
removeWorksheet(wb, 'Down')
addWorksheet(wb, 'Down')
writeData(wb, 'Down', (clues %>% 
                         filter(dir=='down') %>% 
                         add_column(ind=1:sum(clues$dir == 'down'),.before='lbl') %>% 
                         unite(text, lbl, clue, sep=". ", remove=F)))

# Replace 'Title' tab with new title
removeWorksheet(wb, 'Title')
addWorksheet(wb, 'Title')
writeData(wb, 'Title', title)

# Replace 'Solution_Raw' tab with new solutions
removeWorksheet(wb, 'Solution_Raw')
addWorksheet(wb, 'Solution_Raw')
writeData(wb, 'Solution_Raw', answers)

# Save Workbook
saveWorkbook(wb, 'BG_CW.xlsx', overwrite=T)
rm(wb)
