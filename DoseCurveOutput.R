# Analysis Goal: Do the actual calculations for EC50. 
# This is designed for an R-super user, someone who is familiar with R and can do some programming.
# If people unfamiliar with R want to use the code, it should be put into RShiny or RMarkdown.

# Inputs: data.frame with size data and dose data
#         data can be your own data, or data included in the drc library
#         this is going to have problems, I assume your data is formatted correctly
#         (One long column of numerical dose, one long column of numerical size data, wrong inputs removed)

# Analysis Path:
#1 Import your data, or use the S.alba data from the drc package
#2 calculate means, for all clones and controls
#3 decide which data is in the correct range, and add comments to the table
#4 calculate the curve and graph the clone
#5 make a bar graph of EC50 for all clones.
#6 create an output in excel spreadshet for your data

# Outputs: data.frame with summarized size
#          IC50 Results and Model
#          graphs of size with curve
#          excel spreadsheet with data from this assay

#Problems: If you input your own variable, be wary of formatting issues

###############################################################################
rm(list = ls()) #this removes everything in the working directory (no old data to confuse Assay_SampleAlias).

#Install these libraries (this can take a minute or so)
library(drc)

###############################################################################
# 1 Import your data, let us know your name for the dose and continuous variable
getwd() # allows you to see where R is looking now.
#setwd("C:\\Documents") # allows you to send R looking somewhere else
#Data_Raw = read.table("yourdata.csv", header = TRUE, sep = ",")
Data_Raw = S.alba #remove this if you use your own data
head(Data_Raw)

Continuous_Variable = "DryMatter" # name of column in your table
Cont_Var_Name = "Dry Matter Collected (g)" # name to use for graphs
Dose_Variable = "Dose" # name of your dose column
Dose_Var_Name = "Concentration of Herbicide (ppm)" # name to use for graphs
Categorical_Variable = "Herbicide" # name of your sample alias column
Cat_Var_Name = "Herbicide" # name to use for graphs
# Some Colors that are good for colorblind people
cbPalette = c( "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7", "#000000","#999999")
Color_Choice1 = c("#E69F00", "#56B4E9") # orange and blue
Color_Choice2 = c("#009E73", "#F0E442") # green and yellow
Color_Choice3 = c("#0072B2", "#D55E00") # dark blue and dark orange
#the following two variables allow you to track this assay and compare with others
Assay_Start_Date = as.Date("2018-10-01")# date in format year-month-day
Assay_Name = "AP2142"


colnames(Data_Raw)[colnames(Data_Raw) == Continuous_Variable] <- "Continuous_Variable"
colnames(Data_Raw)[colnames(Data_Raw) == Dose_Variable] <- "Dose_Variable"
colnames(Data_Raw)[colnames(Data_Raw) == Categorical_Variable] <- "Categorical_Variable"
head(Data_Raw)
# how many samples?
Num_Samples = unique(Data_Raw$Categorical_Variable)
length(Num_Samples)

# Let's Leap right into modeling!
# We are using a four-parameter log-logistic curve here.  There are many others
# Ask if you would prefer a Weibull curve, 3 parameter logistic curve or something else
# A different curve might provide a better fit, especially at the inflection points.

# Laboratory assays often have the same asymptotes (lower limit and upper limit)
#   i.e., the biggest a plant can get is the same for all treatments, same for the smallest plants
#   If this is not the case, change the pmodels part of the drm below.

# Sometimes you need to play with different kinds of fits to find the curve that works best
# try this website for more info: http://rstats4ag.org/dose-response-curves.html#which-dose-response-model
# here I provide code for:
# Four Parameter Logistic
LL4_Model = drm(Continuous_Variable ~ Dose_Variable, Categorical_Variable, data = Data_Raw, 
                fct = LL.4(names = c("Slope", "Lower Limit", "Upper Limit", "ED50")),
                pmodels = data.frame(Categorical_Variable, 1, 1, Categorical_Variable))
# pmodels this sets the asymptotes  to be equal for each sample, true with most assays
# remove it if your data has different controls
summary(LL4_Model) # show the model
modelFit(LL4_Model) # lack of fit test (p<0.05 if the data does not fit)
ED(LL4_Model,c(10,50,90),interval = "delta")
plot(LL4_Model, broken = TRUE, xlab = Dose_Var_Name, ylab = Cont_Var_Name,
     main = Assay_Name, col = Color_Choice1, 
     pch = 16, cex.axis = 0.8, cex.lab = 1.2, bty = "l")

# The Weibull model tends to be steeper at the asymptotes
W4_Model = drm(Continuous_Variable ~ Dose_Variable, Categorical_Variable, data = Data_Raw,
                fct = W1.4(), pmodels = data.frame(Categorical_Variable, 1, 1, Categorical_Variable))
summary(W4_Model) # show the model
modelFit(W4_Model) # lack of fit test (p<0.05 if the data does not fit)
ED(W4_Model,c(10,50,90),interval = "delta")
plot(W4_Model, add = TRUE, type = "none", col = Color_Choice2)

# Three parameter  logistic models assume that the smallest size is zero 
# If you don't have enough doses at the lower end, this may be a better fit
LL3_Model = drm(Continuous_Variable ~ Dose_Variable, Categorical_Variable, data = Data_Raw, 
                fct = LL.3(names = c("Slope", "Upper Limit", "ED50")),
                pmodels = data.frame(Categorical_Variable, 1, Categorical_Variable))
summary(LL3_Model) # show the model
modelFit(LL3_Model) # lack of fit test (p<0.05 if the data does not fit)
ED(LL3_Model,c(10,50,90),interval = "delta")
plot(LL3_Model, add = TRUE, type = "none", col = Color_Choice3)

# How to pick the best model?
# first are the different?
anova(W4_Model, LL4_Model) # larger F means no difference
AIC(W4_Model, LL4_Model) # smaller AIC means better fit



# The Model summary (model fit coefficients) needs to be saved.
# The Following Code includes formatting for that
Model_Summary = data.frame(summary(LL4_Model)$coef)
Model_Summary$Coefficients = row.names(Model_Summary)
Model_Summary$Sample.Alias = as.character(lapply(strsplit # separate out Name from Coefficient type 
                                           (as.character(Model_Summary$Coefficients), split = ":"), "[", 2))
Model_Summary$Type = as.character(lapply(strsplit # separate out Name from Coefficient type 
                                         (as.character(Model_Summary$Coefficients), split = ":"), "[", 1))
Simple_Model_Summary = Model_Summary [,c(6,7,1:4)]# I like this for reporting in xlsx
names(Simple_Model_Summary)

# Including assay name and date helps later if this is put in a database
Model_Summary$Date = Assay_Start_Date
Model_Summary$Assay = Assay_Name
Model_Summary2 = Model_Summary[,c(8,9,6,7,1:4)]   
# The following saves an unformatted copy of the models (as a .csv)
# This type of file is easy to take up into databases, or use in RShiny
write.csv(Model_Summary2, "Model_Summary.csv", row.names = FALSE)

#The EC50 is what most researchers are interested in 
#(is this sample more active than another sample)
# if you need, you can add EC25 or EC90 by adding in numbers
ED(LL4_Model,c(10,50,90),interval = "delta")
   
EC50_Response = data.frame(ED(LL4_Model,50,interval = "delta"))
EC50_Response$Coefficients = row.names(EC50_Response)
EC50_Response$Sample.Alias = as.character(lapply(strsplit # separate out Name from Coefficient type 
                                         (as.character(EC50_Response$Coefficients), split = ":"), "[", 2))

EC50_Response$Type = as.character(lapply(strsplit # separate out Name from Coefficient type 
                                         (as.character(EC50_Response$Coefficients), split = ":"), "[", 3))
Simple_EC50_Response = EC50_Response[,c(6,1:4)]
names(Simple_EC50_Response) = c("Sample Alias", "EC50","Std Error","Lower 95CI","Upper 95CI")
# Including assay name and date helps later if this is put in a database
EC50_Response$Date = Assay_Start_Date
EC50_Response$Assay = Assay_Name
EC50_Response2 = EC50_Response[,c(8,9,6,7,1:4)]   

# The following saves an unformatted copy of the models (as a .csv)
# This type of file is easy to take up into databases, or use in RShiny
write.csv(EC50_Response2, "EC50_Response.csv", row.names = FALSE)

# The other thing people want in the Report is a graph of the actual values and curve fit
# This graph shows all points, you can change that to the mean by changing type
# These functions are used in the graphs.

Add_ErrorBars = function(dose, continuous, color1){
  # This adds error bars to the figures (for each dose of each toxin)
  # Inputs: Column of dose of each toxin, Column of continuous variable, Color preference for bars
  # Outputs: Point on graph with mean size at each dose.
  #          error bars (95% CI) around each point
  uda = unique(dose)
  ud.a = sort(uda)
  mean.a = tapply(continuous, list(dose), mean, na.rm = TRUE)
  ci.a = tapply(continuous, list(dose), Calculate_95_Confidence_Interval)
  suppressWarnings(arrows(ud.a, mean.a-ci.a, ud.a, mean.a+ci.a, col = color1, code = 3, 
         angle = 90, length = 0.05))
  }

Calculate_95_Confidence_Interval = function(x){
  # This calculates the 95% Confidence Interval.  There isn't one in base R.
  # Inputs: vector (column) needed for calculation
  # Outputs: one number that is the 95% Confidence Interval
  sd1 = sd(x, na.rm = TRUE)
  l1 = length(x)
  standard.error = sd1/(sqrt(l1))
  ci = 1.96*standard.error
  ci
}

Dose_Curve_Graph = paste(Assay_Name, Assay_Start_Date, "EC50.png", sep = "_")
jpeg(Dose_Curve_Graph)
plot(LL4_Model, broken = TRUE, xlab = Dose_Var_Name, ylab = Cont_Var_Name,
     main = Assay_Name, col = Color_Choice1, 
     pch = 16, cex.axis = 0.8, cex.lab = 1.2, bty = "l")

# for data with many samples it can be easier to display the mean and CL of each dose.
for (i in 1:length(Num_Samples)){
  # First get a data frame of the unique clones
  # Then plot the 95% CI from the mean
  Sample_Data = Data_Raw[Data_Raw$Categorical_Variable == Num_Samples[i],]
  Add_ErrorBars(Sample_Data$Dose, Sample_Data$Continuous_Variable,
                       Color_Choice1[i])
  # Note that the buffer point does not graph
  # This is something that I need to improve on
}
dev.off() 

# Sometimes people like to see the output based on each dose.
# This provides a summary of the continuous variable at each dose
library(tidyverse)
head(Data_Raw)
Cat_Var_By_Dose <- Data_Raw %>% 
  group_by(Categorical_Variable, Dose_Variable) %>% 
  summarise(N.Observations = sum(!is.na(Continuous_Variable)),
            Mean.Size = mean(Continuous_Variable, na.rm = TRUE),
            Median.Size = median(Continuous_Variable, na.rm = TRUE),
            StDev = sd(Continuous_Variable, na.rm = TRUE),
            CV1 = StDev/Mean.Size,
            Size.95CI = Calculate_95_Confidence_Interval(Continuous_Variable))
Simple_Cat_Var_By_Dose = Cat_Var_By_Dose
names(Simple_Cat_Var_By_Dose) = c("Sample Alias", "Dose", "N.Obs", 
                  "Mean Size", "Median Size", "StDev", "CV", "Size 95CI")

# Including assay name and date helps later if this is put in a database
Cat_Var_By_Dose$Date = Assay_Start_Date
Cat_Var_By_Dose$Assay = Assay_Name

Cat_Var_By_Dose2 = Cat_Var_By_Dose[,c(9,10,1:8)]   
# save it for later use
write.csv(data.frame(Cat_Var_By_Dose2), "Summary_by_Dose.csv", row.names = FALSE)

EC50_Response_Graph = paste(Assay_Name, "EC50_bargraph.png", sep = "_")
jpeg(EC50_Response_Graph)
names(EC50_Response2)
ggplot(EC50_Response2, aes(x = EC50_Response2$Sample.Alias, y = Estimate, fill = EC50_Response2$Sample.Alias ))+
  geom_bar(position = position_dodge(.9), colour = "black", stat = "identity")+
  geom_errorbar(position = position_dodge(.9), width = .25, aes(ymin = Lower, ymax = Upper))+
  theme_bw() +
  labs(fill = Categorical_Variable,  x = Categorical_Variable)+
  theme(axis.text.x = element_text(angle = 90, hjust = 1))+
  ggtitle(paste(Assay_Name, "EC50 and 95%CL", sep = " "))
dev.off()

#  
library(rJava)
library(xlsxjars)
library(xlsx)

###############################################################################
#This creates the final spreadsheet.  You can change the Column Width, 
#change the format (here all I do is use different significant digits), 
#it also includes a loop for adding figures to the file.
my.wb = createWorkbook(type = "xlsx")
# First: Name sheets, and set column width
sheet.1 = createSheet(my.wb, sheetName = "Results")
setColumnWidth(sheet.1, c(1), 19)
sheet.2 = createSheet(my.wb, sheetName = "SummaryData")
setColumnWidth(sheet.2, c(1), 17)
sheet.3 = createSheet(my.wb, sheetName = "ModelFit")
setColumnWidth(sheet.3, c(1), 17)
setColumnWidth(sheet.3, c(2), 13)
# Second: Set formatting, decimal places, font, etc
cs1 = CellStyle(my.wb, dataFormat = DataFormat("0"))    # no decimal
cs2 = CellStyle(my.wb, dataFormat = DataFormat("0.0%")) # for percentage results
cs3 = CellStyle(my.wb, dataFormat = DataFormat("0.00")) # two decimal points
cs5 = CellStyle(my.wb, dataFormat = DataFormat("0.00000")) # 5 decimal points (p-value)
cs4 = CellStyle(my.wb) + Font(my.wb, isBold = TRUE) + Alignment(wrapText = TRUE) + Border()
# by column, set style (decimal places)
model1 = list('2' = cs3, '3' = cs3, '4' = cs3, '5' = cs3, '6' = cs5)
Response.1 = list('1' = cs3,'2' = cs3,'3' = cs3, '4' = cs3, '5' = cs3, '6' = cs3)
Mean.data = list('2' = cs3,'3' = cs1, '4' = cs3, '5' = cs3, '6' = cs3, '7' = cs3, '8' = cs3)
# Third fill in sheets with data and graphs
# sheet.1 holds EC50 data, and graphs 
addDataFrame(Simple_EC50_Response, sheet = sheet.1, row.names = FALSE, startRow = 1, 
             startColumn = 1, colnamesStyle = cs4, colStyle = Response.1)
addPicture(file = Dose_Curve_Graph, sheet = sheet.1, scale = 1, startRow = length(Num_Samples)+10,
           startColumn = 1)
addPicture(file = EC50_Response_Graph, sheet = sheet.1, scale = 1, startRow = length(Num_Samples)+10,
           startColumn = 10)
# sheet.2 holds summary data
addDataFrame(as.data.frame(Simple_Cat_Var_By_Dose), sheet = sheet.2, row.names = FALSE, startRow = 1, 
             startColumn = 1, colnamesStyle = cs4, colStyle = Mean.data)
# sheet.3 holds model fit info
addDataFrame(Simple_Model_Summary, sheet = sheet.3, row.names = FALSE, startRow = 1, 
             startColumn = 1, colnamesStyle = cs4, colStyle = model1)
# Fourth: name file and save.
Workbook_Name = paste(Assay_Name,"EC50_Results.xlsx", sep = "_")
saveWorkbook(my.wb, Workbook_Name)

A = 42 # This just tells you the program is done.
