#################################################################################################
#authors: "Valentin Scherz and Linda Mueller"
#date: 25/06/2020
#version: 1.0.0

#################################################################################################
################################### LICENSE #####################################################
#################################################################################################

#MIT License

#Copyright (c) 2020 Valentin Scherz

#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:

#The above copyright notice and this permission notice shall be included in all
#copies or substantial portions of the Software.

#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#SOFTWARE.


###################################################################################################
###################################### PARAMS #####################################################
###################################################################################################

## Those variables should be adapted to fit your setting and the structure of the table with results

### ".xlsx" table with new results to be added to the database of results. This table should be 
#   structured with one line per analysis and variables (columns) matching all the parameters defined 
#   in the next chunk of code. The structure of this table must remain stable for all execution of the
#   script. 
input_table_path <- "extraction1202-1206.xlsx"

### Sample type dictionary file(see requirements for structure)
dict_path <- "../Ressources/sample_source_dictionnary.xlsx"

### Script file name (a copy of this script script is copied after each run for tractability)
script_path <- "20200423_discordance_analysis_for_publication.R"

### Some QC samples could be included in the provided input table. Provide a variable name (column)
#   in the input spread sheet and a value (pattern) matching these controls to filter these out.
QC_column <- "Nom pat."
QC_pattern <- "ZZ IMU"

### Variable (column) describing the test results and the possible values to describe positivity 
#   and negativity (e.g. NEGATIVE/POSITIVE)
Results_col  <- "COWE"  # Column name describing test result
Results_variables <- c("POSITIVE", "NEGATIVE") # Column name describing acceptable values in this column

### Variable (column) with patient identifier
patient_ID <- "NIP"

### Variable (column) with sample identifier (unique)
sample_ID <- "Numero Alias"

### Variable (column) with sample reception date. See e.g. https://www.r-bloggers.com/date-formats-in-r/ to see how to format your dates
date_col <- "Dt.recept."
date_form <- "%d/%m/%Y"

#################################################################################################
################################### DON'T TOUCH #################################################
#################################################################################################

## Libraries 
### These R packages sould be installed beforehand
library(readxl)
library(tidyr)
library(dplyr)
library(writexl)
library(ggplot2)
library(scales)
library(furniture)
library(htmlTable)
library(colorspace)
library(forcats)
library(rtf)
library(data.table)
library(rlang)

### Define a simple useful "not in" function for later use
'%!in%' <- function(x,y)!('%in%'(x,y))


## Import tables
### New cases to add to the database of results. 
new_data <- read_xlsx(input_table_path , col_types = "guess", guess_max = 1000000) 

### Validated cases (input)
if(file.exists("2_solved_cases_updated.xlsx")){
  validated_discordant <- read_xlsx("2_solved_cases_updated.xlsx", col_types = "guess", guess_max = 1000000) 
}else{
  validated_discordant <- data.frame(matrix(ncol=length(colnames(data)),nrow=0, dimnames=list(NULL, colnames(data))), check.names = FALSE)
  write_xlsx(validated_discordant, "2_solved_cases_updated.xlsx")
}

### Previously compiled database of samples
if(file.exists("0_previous_results_database.xlsx")){
  sample_db <- read_xlsx("0_previous_results_database.xlsx", col_types = "guess", guess_max = 1000000) 
}else{
  sample_db <- data.frame()
}


### Dictionaries translating and regrouping the "sample type" column
specimen_dictionary <- read_xlsx(dict_path, sheet = 1)
bad_sample <-  read_xlsx(dict_path, sheet = 2)

## Determine time and user flag. Create an archive directory with it
st <- format(Sys.time(), "%Y%m%d_%H%M%S")
us <- Sys.info()[["user"]]
file_flag <- paste(st, us,sep = "_")
### Create a directory
dir.create(file_flag)

## Back-up this script in the archive directory for tractability
file.copy(from = script_path, to = paste0(file_flag, "/", "R_script_", file_flag, ".R")) 

### Treat newly extracted data
#### Filter out and write in a table problematic samples missing a patient ID (which are anyway problematic but also are not compatible with the way this comparative script is built)
missing_NIP <- new_data[is.na(new_data[[patient_ID]]),]

#### write the list of these samples if not empty
if(length(missing_NIP[[sample_ID]])>0){
  ##### Write these samples to be valited for traceability
  write_xlsx(x = missing_NIP, path = paste0(file_flag, "/","3_missing_IPP_",file_flag,".xlsx"))
  write_xlsx(x = missing_NIP, path = "3_missing_IPP.xlsx")
  
}else{
  
  #### else, remove the table with missinig table, if previously existing
  file.remove("3_missing_IPP.xlsx")
}

#### Filter out any result not "POSITIVE" nor "NEGATIVE" or which are missing a patient ID
new_data_P_N <- new_data[new_data[[Results_col]] %in% Results_variables & new_data[[sample_ID]] %!in% missing_NIP[[sample_ID]],]


#### Back-up the filtered input table for traceability 
write_xlsx(x = new_data_P_N, path = paste0(file_flag, "/","0_input_",file_flag,".xlsx" ))


### Add these to the longitudinal databases
#### Look for samples already in the db but for which result could have change in the newly extracted database
##### Find samples in the newly extracted table that were already in the DB.
data_old_samples <- new_data_P_N[new_data_P_N[[sample_ID]] %in% sample_db[[sample_ID]],]

##### Match order of old samples with newly extracted database
m <- match(new_data_P_N[[sample_ID]], sample_db[[sample_ID]])
changed_results <- new_data_P_N[[sample_ID]][new_data_P_N[[Results_col]] != sample_db[[Results_col]][m]]

##### Remove changed samples from the database
sample_db_filt <- sample_db[sample_db[[sample_ID]] %!in% changed_results,]

#### Remove samples from new extraction already in filtered sample_db
data_new_samples <- new_data_P_N[new_data_P_N[[sample_ID]] %!in% sample_db_filt[[sample_ID]],]

#### Join these new samples to the longitudinal database
data_completed <- rbind(sample_db_filt, data_new_samples)

#### Filter controls out
data_completed <- data_completed[grep(data_completed[[QC_column]], pattern = QC_pattern , invert = TRUE),]

#### Write the new completed table in the root file and in a back-up
write_xlsx(x = data_completed, path = paste0(file_flag, "/","0_previous_results_database.xlsx",file_flag,".xlsx" ))
write_xlsx(x = data_completed, path = "0_previous_results_database.xlsx")

### Plot an histogram of positive/negatives cases
#### Plot
tests_hist <- ggplot(data=data_completed, aes(as.Date(!!rlang::sym(date_col), format = "%d/%m/%Y"), fill = !!rlang::sym(Results_col))) +
  geom_histogram(binwidth = 1, colour="white", size=0.1, alpha = 0.8) +
  theme_bw() +
  xlab("Analyses") +
  scale_y_continuous(breaks= pretty_breaks()) +
  labs(title = "N. of daily tests",
       subtitle = paste("Concerns", length(unique(data_completed[[patient_ID]])), "unique patients and", length(unique(data_completed[[sample_ID]])), "samples") ,
       caption = "All tested samples for the extracted period")


#### Save this plot
ggsave(tests_hist, filename =  paste0(file_flag, "/", "0_input_extracted_",file_flag,"_tests_hist.png"), width = 7, height = 7)


## Already validated cases
### Translate in data the material to the new description and if it is a Low yield sample
m <- match(data_completed$Materiel, specimen_dictionary$specimen_orig)
data_completed$Materiel_ENG <- specimen_dictionary$specimen_ENG[m]

n <- match(data_completed$Materiel_ENG, bad_sample$specimen_ENG)
data_completed$Bad_sample <- bad_sample$Bad_sample[n]


## Find samples with undefined material
missing_sample_type <-  data_completed$Materiel[is.na(data_completed$Materiel_ENG)]

#### write the list of these samples if not empty
if(length(missing_sample_type)>0){
  ##### Write these samples to be valited for traceability
  write_xlsx(x = as.data.frame(missing_sample_type), path = paste0(file_flag, "/","3_sample_source_dictionary_missing",file_flag,".xlsx"))
  write_xlsx(x = as.data.frame(missing_sample_type), path = "3_sample_source_dictionary_missing.xlsx")
  
  stop('Unkown sample types. Those were written in "3_sample_source_dictionary_missing.xlsx". These should be added to "sample_source_dictionary.xlsx" and the script re-run')
  
}else{
  
  #### else, remove the table with missinig table, if previously existing
  file.remove("3_sample_source_dictionary_missing.xlsx")
}



#### Compute a time to positivity for discordant cases   
## A big loop to evaluate potential discordancy in samples

### Generate columns to be filled
data_completed$compared_with <- NA
data_completed$compared_to_previous <- NA
data_completed$time_to_conco <- NA
data_completed$time_to_disco <- NA
data_completed$detailed_explanation <- NA
data_completed$broad_explanation <- NA

### Find patient with multiple analyses
multi_test_pat <- unique(data_completed$NIP[duplicated(data_completed$NIP)])

### Loop over each patient tested multiple time
for (p in multi_test_pat){
  
  data_completed_p <- filter(data_completed, NIP == p) %>% arrange(as.POSIXct(paste(Dt.recept., H.recept.), format = "%d/%m/%Y %H:%M"))# arrange by patient and time.
  
  if (length(data_completed_p$`Numero Alias`)>1){
    
    #### Reset to NA these values when starting with a new patient
    previous_sample_ID <- NA
    previous_result <- NA
    previous_date <- NA
    previous_Ct <- NA
    previous_bad_sample <- NA
    previous_sample_type <- NA
    
    time_to <- NA
    
    #### Loop over all samples from a patient to apply classification algorithm
    for (a in data_completed_p$`Numero Alias`){
      
      ##### Extract some data for the currently evaluated sample
      current_sample_ID <- a
      current_result <- data_completed_p$COWE[data_completed_p$`Numero Alias`== a]
      current_date <- data_completed_p$Dt.recept.[data_completed_p$`Numero Alias`== a]
      current_Ct <- max(as.numeric(as.character(data_completed_p[data_completed_p$`Numero Alias`== a, c("COVWIP", "COWEIP", "CTCWE", "CTCWEC", "CTC1", "CTC2", "CTCNGX", "CTCEGX")], na.rm = TRUE)), na.rm =  TRUE)
      current_bad_sample <- data_completed_p$Bad_sample[data_completed_p$`Numero Alias`== a]
      current_sample_type <- data_completed_p$Materiel_ENG[data_completed_p$`Numero Alias`== a]
      
      time_to <- as.Date(current_date, format = "%d/%m/%Y") - as.Date(previous_date, format = "%d/%m/%Y")
      
      ##### If there was no previous sample, compared is NA
      if(is.na(previous_result)){
        
        data_completed$compared_to_previous[data_completed$`Numero Alias`== a] <- NA
        
        ##### If previous was the same, compared_to is "concordant"
      } else if(previous_result == current_result){
        
        data_completed$compared_to_previous[data_completed$`Numero Alias`== a] <- "Concordant"
        
        data_completed$time_to_conco[data_completed$`Numero Alias`== a] <- time_to
        
        data_completed$compared_with[data_completed$`Numero Alias`== a] <- previous_sample_ID
        
        ##### If previous was different, then we apply a list of criteria.  
      } else if (previous_result != current_result){
        
        data_completed$compared_to_previous[data_completed$`Numero Alias`== a] <- "Discordant"
        
        data_completed$compared_with[data_completed$`Numero Alias`== a] <- previous_sample_ID
        
        if(current_result == "POSITIVE"){
          
          data_completed$time_to_disco[data_completed$`Numero Alias`== a] <- time_to
          
          if(previous_bad_sample == "Y"){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Previous sample was a Low yield sample"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Low yield sample"
            
          } else if(current_Ct > 35 & previous_sample_type == current_sample_type){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Current sample has high Ct, previous sample could have been affected by stochasticity"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Stochastic"
            
          } else if (time_to > 10 ){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Previous NEG sample was > 10 days, new infection likely"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Time delay"
            
          } else{
            
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "To be investigated"
            
            
          }
          
          
        }else if(current_result == "NEGATIVE"){
          
          data_completed$time_to_disco[data_completed$`Numero Alias`== a] <- -time_to 
          
          if(current_bad_sample == "Y"){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Current sample is a Low yield sample"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Low yield sample"
            
            next()
            
          } else if(previous_Ct > 35 & previous_sample_type == current_sample_type){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Previous sample had high Ct, this sample could be affected by stochasticity"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Stochastic"
            
            next()
            
          } else if (time_to > 10 ){
            
            data_completed$detailed_explanation[data_completed$`Numero Alias`== a] <- "Previous POS sample was > 10 days, infection resolution likely"
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "Time delay"
            
          } else{
            
            data_completed$broad_explanation[data_completed$`Numero Alias`== a] <- "To be investigated"
            
            next()
            
          }
          
        }
        
        
      }
      
      
      ### Report the current as previous  
      previous_sample_ID <- current_sample_ID  
      previous_result <- current_result
      previous_date <- current_date
      previous_Ct <- current_Ct
      previous_bad_sample <- current_bad_sample
      previous_sample_type <- current_sample_type
      
    }
  }
  
  
}

### Keep only patients not belonging to the list of multiple tested patients
unique_test_pat_df <- data_completed[data_completed$NIP %!in% multi_test_pat,]

### Write this table of uniquely tested patients for traceability
write_xlsx(x = unique_test_pat_df, path = paste0(file_flag, "/","1_unique_test_patients_",file_flag,".xlsx" ))

## Patients tested multiple times
### Keep only samples from patients belonging to the list of multiple tested patients
multi_test_pat_df <- data_completed[data_completed$NIP %in% multi_test_pat,]

### Write this table of multiple tested patients for traceability
write_xlsx(x = multi_test_pat_df, path = paste0(file_flag, "/","1_multiple_tests_patients_",file_flag,".xlsx" ))


### Concordant only analyses
#### Filter
concordant_time_to <- data_completed %>% 
  filter(compared_to_previous == "Concordant") %>% 
  group_by(NIP, Bad_sample) %>%
  arrange(`Nom pat.`, as.Date(Dt.recept., format = "%d/%m/%Y")) # arrange by patient and time.


#### Plot time to for concordant cases
concord_hist <- ggplot(data=concordant_time_to, aes(abs(time_to_conco), fill = COWE)) +
  geom_histogram(binwidth = 1, colour="black") +
  facet_grid(Dem.group ~ .) +
  theme_bw() +
  xlab("Time lap between concordant analyses") +
  scale_y_continuous(breaks= pretty_breaks()) +
  labs(title = "Time to concordant SARS-CoV-2 analyses",
       subtitle = paste("Concerns", length(unique(concordant_time_to$NIP)), "unique patients and", length(unique(concordant_time_to$`Numero Alias`)), "samples") ,
       caption = "Time lag for concordant SARS-CoV-2 analyses")


#### Save this plot
ggsave(concord_hist, filename =  paste0(file_flag, "/", "2_multiple_tests_patients_concordant_",file_flag,"_time_to_hist.png"), width = 7, height = 7)

#### Write them for traceability
write_xlsx(x = concordant_time_to, path = paste0(file_flag, "/","2_multiple_tests_patients_concordant_",file_flag,".xlsx"))


### Discordant cases analysis
#### Filter and format discordant cases
disco_patients <- unique(data_completed$NIP[data_completed$compared_to_previous == "Discordant"])

discordant_time_to <- data_completed %>% 
  filter(NIP %in% disco_patients) %>% 
  group_by(NIP, Bad_sample) %>%
  arrange(`Nom pat.`, as.Date(Dt.recept., format = "%d/%m/%Y")) # arrange by patient and time.


#### Plot already validated cases in a histogram indicating the time to discordancy
##### Adapat labels for bad samples 
discordant_time_to_relabeled <- discordant_time_to
discordant_time_to_relabeled$Bad_sample <- factor(discordant_time_to_relabeled$Bad_sample, 
                                                  levels=c("Y", NA),
                                                  labels=c('Bad samples', 'Usual samples'), exclude = NULL)

##### Plot
hist_time_to <- ggplot(data=discordant_time_to_relabeled, aes(time_to_disco, fill = broad_explanation, colour=broad_explanation)) + 
  geom_histogram(binwidth = 1) +
  theme_bw() +
  xlab("Time to positivity [Days]") +
  #scale_y_continuous(breaks= pretty_breaks()) +
  #scale_x_continuous(breaks= pretty_breaks(n = 10)) +
  labs(title = "Time laps between discordant analyses",
       subtitle = paste("Concerns", length(unique(discordant_time_to_relabeled$NIP)), "unique patients and", length(unique(discordant_time_to_relabeled$`Numero Alias`)), "samples"),
       caption = "Positive values for NEG --> POS, negative for POS --> NEG transitions") 
#facet_grid(Bad_sample ~ .)


##### Save this plot
ggsave(hist_time_to, filename =  paste0(file_flag, "/","2_input_discordant_time_to_hist.png"), width = 7, height = 7)



#### Filter out of these discordant cases already treated samples
to_be_investigated_patients <- data_completed$NIP[which(data_completed$broad_explanation == "To be investigated")]  
analyses_of_to_be_investigated_patients <- data_completed[data_completed$NIP %in% to_be_investigated_patients,] 

to_be_validated_by_hand <- analyses_of_to_be_investigated_patients[analyses_of_to_be_investigated_patients$`Numero Alias` %!in% validated_discordant$`Numero Alias`,]
to_be_validated_by_hand <- to_be_validated_by_hand %>% 
  group_by(NIP, Bad_sample) %>%
  arrange(`Nom pat.`, as.Date(Dt.recept., format = "%d/%m/%Y"))


#### Write these samples to be valited for traceability
write_xlsx(x = to_be_validated_by_hand, path = paste0(file_flag, "/","3_to_be_validated.xlsx"))

#### Write these samples to be valited in the active file
write_xlsx(x = to_be_validated_by_hand, path = "3_to_be_validated_.xlsx")




