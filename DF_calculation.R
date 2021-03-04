rm(list=ls())
install.packages("FIACH")
install.packages("openxlsx")
library(FIACH)
library(openxlsx)

setwd("path")
marker <- "path"
cur_files <- list.dirs(full.name = FALSE, recursive = FALSE)
print(cur_files)

DF_calculation = c()

for(i in 1:length(cur_files)){
  
  if(!grepl("\\d$", cur_files[i]) && !grepl("[[:upper:]]$", cur_files[i]))
    next
  
  x <- paste(marker, cur_files[i], sep = "")
  rp_files <- list.files(path = x, pattern = "rp_", full.names = FALSE, recursive = TRUE)
  
  rp_1 <- paste(x, rp_files[1], sep = "/")
  rp_2 <- paste(x, rp_files[2], sep = "/")
  rp_3 <- paste(x, rp_files[3], sep = "/")
  
  if(file.exists(rp_1)){
    fd_rp1 <- fd(rp_1)
    rp_1_mean <- mean(fd_rp1)
  }
  if(file.exists(rp_2)){
    fd_rp2 <- fd(rp_2)
    rp_2_mean <- mean(fd_rp2)
  }
  if(file.exists(rp_3)){
    fd_rp3 <- fd(rp_3)
    rp_3_mean <- mean(fd_rp3)
  }
  total <- 0
  
  if(file.exists(rp_1) & file.exists(rp_2) & file.exists(rp_3)){
    total <- fd_rp1 + fd_rp2 + fd_rp3
    DF_calculation[i] <- mean(total)
  }
  if(file.exists(rp_1) & file.exists(rp_2) & !file.exists(rp_3)){
    total <- fd_rp1 + fd_rp2
    DF_calculation[i] <- mean(total)
  }
  if(file.exists(rp_1) & !file.exists(rp_2) & file.exists(rp_3)){
    total <- fd_rp1 + fd_rp3
    DF_calculation[i] <- mean(total)
  }
  if(!file.exists(rp_1) & file.exists(rp_2) & file.exists(rp_3)){
    total <- fd_rp2 + fd_rp3
    DF_calculation[i] <- mean(total)
  }
  if(file.exists(rp_1) & !file.exists(rp_2) & !file.exists(rp_3)){
    DF_calculation[i] <- rp_1_mean
  }
  if(!file.exists(rp_1) & file.exists(rp_2) & !file.exists(rp_3)){
    DF_calculation[i] <- rp_2_mean
  }
  if(!file.exists(rp_1) & !file.exists(rp_2) & file.exists(rp_3)){
    DF_calculation[i] <- rp_3_mean
  }
  if(!file.exists(rp_1) & !file.exists(rp_2) & !file.exists(rp_3)){
    DF_calculation[i] <- total
  }
  
  print(cur_files[i])
}

#print(DF_calculation)

#for(k in 1:length(DF_calculation)){
#  print(DF_calculation[k])
#}



setwd("new path")

wb <- loadWorkbook("spreadsheet name")
writeData(wb, sheet = "Sheet1", DF_calculation, startCol = "HY", startRow = 2, colNames = TRUE)
saveWorkbook(wb,"spreadsheet name",overwrite = TRUE)
