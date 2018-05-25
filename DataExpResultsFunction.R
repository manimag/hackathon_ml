library(scales)
library(sqldf)
library(XLConnect)
library(ggplot2)
library(rJava)
library(dplyr)


RemoveDuplicates <- function(data){
 data <-  unique(data)
 return(data)
}

separateVariables <- function(data)
{
char <- c()
num <- c()
bin <- c()
date <- c()

for(i in 1:ncol(data))
{
  if(class(data[,i])=="character" | class(data[,i]) =="factor") {
    char[i] <- c(names(data)[i])
  } else if(class(data[,i])=="numeric" | class(data[,i])=="integer") {
    num[i] <- c(names(data)[i])
  } else if(class(data[,i])=="logical") {
    bin[i] <- c(names(data)[i])
  } else if(class(data[,i])=="Date") {
    date[i] <- c(names(data)[i])
  }
}
char <- char[!is.na(char)]
char <- data[,char]

num <- num[!is.na(num)]
num <- data[,num]

bin <- bin[!is.na(bin)]
bin <- data[,bin]

date <- date[!is.na(date)]
date <- data[,date]
return(list(num,char,bin,date))
}


VariablesInformation <- function(data){
  Names <- names(data)
  Class <- sapply(data, function(x) class(x))
  NoofObservation <- sapply(data, function(x) length(x))
  Missing_Count <- sapply(data, function(x) sum(is.na(x)))
  Missing_Percentage <- percent((Missing_Count/nrow(data)))
  VI <- data.frame(Names, Class, NoofObservation, Missing_Count, Missing_Percentage)
  return(VI)
}

GetNominaldataVariablenames <- function(data){
  Noofuniquevales <- (sapply(num, function(x) length(unique(x))))
}

ExportExcel <- function(data, files, sheetNames){
  writeWorksheetToFile(file = files, data = data, sheet = sheetNames)
}

ExportExcel.Specified.Location <- function(data, files, sheetNames, startRow, startCol){
  writeWorksheetToFile(file = files, data = data, sheet = sheetNames, startRow = startRow,
                       startCol = startCol)
}

SummaryNumericalVariables <- function(data){

  Names <- names(data)
  Mean <- sapply(data, function(x) mean(x, na.rm = T))
  Median <- sapply(data, function(x) median(x, na.rm = T))
  Min <- sapply(data, function(x) min(x, na.rm = T))
  SD <- sapply(data, function(x) sd(x, na.rm = T))
  CV <- ((SD/Mean)*100)
  P1 <- sapply(data, function(x) quantile(x, 0.01, na.rm = T))
  P5 <- sapply(data, function(x) quantile(x, 0.05, na.rm = T))
  P10 <- sapply(data, function(x) quantile(x, 0.10, na.rm = T))
  P20 <- sapply(data, function(x) quantile(x, 0.20, na.rm = T))
  P25 <- sapply(data, function(x) quantile(x, 0.25, na.rm = T))
  P30 <- sapply(data, function(x) quantile(x, 0.30, na.rm = T))
  P40 <- sapply(data, function(x) quantile(x, 0.40, na.rm = T))
  P50 <- sapply(data, function(x) quantile(x, 0.50, na.rm = T))
  P60 <- sapply(data, function(x) quantile(x, 0.60, na.rm = T))
  P70 <- sapply(data, function(x) quantile(x, 0.70, na.rm = T))
  P75 <- sapply(data, function(x) quantile(x, 0.75, na.rm = T))
  P80 <- sapply(data, function(x) quantile(x, 0.80, na.rm = T))
  P90 <- sapply(data, function(x) quantile(x, 0.90, na.rm = T))
  P95 <- sapply(data, function(x) quantile(x, 0.95, na.rm = T))
  P99 <- sapply(data, function(x) quantile(x, 0.99, na.rm = T))
  Max <- sapply(data, function(x) max(x, na.rm = T))
  IQR <- sapply(data, function(x) IQR(x, na.rm = T))
  SNV <- data.frame(Names, Mean, Median, Min, SD, CV, P1, P5, P10, P20, P25, P30, P40, P50,
                    P60, P70, P75, P80, P90, P95, P99, Max, IQR)
  return(SNV)
}

Getnames <- function(data){
  Names <- names(data)
  return(Names)
}

AnalysisForEachNumericVariable <- function(num, files, Dep.var, Dep.var.status, Dep.Var.Type){
                                           
  dep.data <- data[Dep.var]
  Names <- Getnames(data = num)
  for (i in 1:ncol(num)){
    mydata <- num[Names[i]]
    VI <- VariablesInformation(data = mydata)
    SNV <- SummaryNumericalVariables(data = mydata)
    Analysis <- merge(x=VI, y=SNV, by = "Names", all = TRUE)
    ExportExcel(data = Analysis, files = files, sheetNames = Names[i])
    wb = loadWorkbook(files, create=FALSE)
    boxp <- paste(Names[i],'boxp' , sep = '.')
    
    formula.box <- paste(Names[i],'!$L$8', sep = '')
    
    createName(wb, name = boxp, formula = formula.box)
    
    png(filename = "box.png", width = 500, height = 400)
    boxplot(mydata[, Names[i]], main = paste('Boxplot of', Names[i], sep = " "), xlab = Names[i])
    dev.off()
    
    addImage(wb, filename = "box.png", name = boxp, originalSize = TRUE)
    
    saveWorkbook(wb)
    
    removeName(wb, boxp)
    gc(reset = TRUE)
    
    if(Dep.var.status == "Yes"){
      if(Dep.Var.Type == "continuous"){
        if(Dep.var != Names[i]){
          
          histo <- paste(Names[i],'histo' , sep = '.')
          scatterp <- paste(Names[i],'scatter' , sep = '.')
          
          formula.hist <- paste(Names[i],'!$B$8', sep = '')
          formula.scatter <- paste(Names[i],'!$B$30', sep = '')
          
          createName(wb, name = histo, formula = formula.hist)
          createName(wb, name = scatterp, formula = formula.scatter)
          
          png(filename = "hist.png", width = 500, height = 400)
          hist(mydata[, Names[i]], main = paste('Histogram of', Names[i], sep = " "), xlab = Names[i], col = "red")
          dev.off()
          
          png(filename = "scatter.png", width = 500, height = 400)
          plot(mydata[, Names[i]], dep.data[, Dep.var], main = paste('Scatter plot between', Names[i], "Vs", Dep.var, sep = " "), 
               xlab = Names[i], ylab = Dep.var, col = "blue")
          dev.off()
          
          addImage(wb, filename = "hist.png", name = histo, originalSize = TRUE)
          addImage(wb, filename = "scatter.png", name = scatterp, originalSize = TRUE)
          
          saveWorkbook(wb)
          
          removeName(wb, histo)
          removeName(wb, scatterp)
        }
      } else if(Dep.Var.Type == "binary"){
          
          if(Dep.var != Names[i]){
        
          combinedata <- cbind(dep.data, mydata)
          names(combinedata)[1] <- "Group"
          names(combinedata)[2] <- "Var"
          
          histo <- paste(Names[i],'histo' , sep = '.')
          
          formula.hist <- paste(Names[i],'!$B$8', sep = '')
          
          createName(wb, name = histo, formula = formula.hist)
          
          png(filename = "hist.png", width = 500, height = 400)
          ## first plot - left half of x-axis, right margin set to 0 lines
          par(fig = c(0, .5, 0, 1), mar = c(5,4,3,0))
          hist(combinedata$Var[combinedata$Group == 0], ann = FALSE, las = 1, col = "green")
          
          ## second plot - right half of x-axis, left margin set to 0 lines
          par(fig = c(.5, 1, 0, 1), mar = c(5,0,3,2), new = TRUE)
          hist(combinedata$Var[combinedata$Group == 1], ann = FALSE, axes = FALSE, col = "yellow")
          axis(1)
          axis(2, lwd.ticks = 0, labels = FALSE)
          
          title(main = paste('Histogram of', Names[i], sep = " "), xlab = Names[i], outer = TRUE, line = -2)
          dev.off()
          
          addImage(wb, filename = "hist.png", name = histo, originalSize = TRUE)
          
          saveWorkbook(wb)
          
          removeName(wb, histo)
          
          group <- group_by(combinedata, Group)
          Analysis <- as.data.frame(summarise(group, mean(Var), median(Var), 
                                              sd(Var), sum(Var)))
          Analysis.N <- Assignnames(x = c(Dep.var, "Mean", "Median", "SD", "SUM"), data = Analysis)
          Analysis.N$Count <- table(combinedata$Group)
          for( j in 1:nrow(Analysis.N)){
            Analysis.N$Count.Percentage[j] <- percent((Analysis.N$Count[j]/nrow(data)))
          }
          ExportExcel.Specified.Location(data = Analysis.N, files = files, sheetNames = Names[i], 
                                         30, 3)
          }
      }
    
    } else {
      
      histo <- paste(Names[i],'histo' , sep = '.')
      
      formula.hist <- paste(Names[i],'!$B$8', sep = '')
      
      createName(wb, name = histo, formula = formula.hist)
      
      png(filename = "hist.png", width = 500, height = 400)
      hist(mydata[, Names[i]], main = paste('Histogram of', Names[i], sep = " "), xlab = Names[i], col = "red")
      dev.off()
      
      addImage(wb, filename = "hist.png", name = histo, originalSize = TRUE)
      
      saveWorkbook(wb)
      
      removeName(wb, histo)
    }
  }   
} 

CorrelationMatrix <- function(data, method) {
  removedNA <- na.omit(data)
  m <- cor(removedNA)
  return(m)
}

AnalysisForEachCharacterVariable <- function(char, files, Dep.var, Dep.var.status,Dep.Var.Type){
  dep.data <- data[Dep.var]
  Names_f <- Getnames(data = char)
  for (i in 1:ncol(char)){
    mydata <- char[Names_f[i]]
    Names <- names(mydata)
    Type <- sapply(mydata, function(x) class(x))
    NoofObservation <- sapply(mydata, function(x) length(x))
    Missing_Count <- sapply(mydata, function(x) sum(is.na(x)))
    Missing_Percentage <- percent((Missing_Count/nrow(mydata)))
    data.f <- data.frame(Names, Type, Missing_Percentage)
    ExportExcel(data = data.f, files = files, sheetNames = Names_f[i])
    
    wb = loadWorkbook(files, create=FALSE)
    
    piec <- paste(Names_f[i],'piec' , sep = '.')
    formula.piec <- paste(Names_f[i],'!$B$8', sep = '')
    
    createName(wb, name = piec, formula = formula.piec)

    png(filename = "piec.png", width = 500, height = 400)
    mytable <- table(mydata[, Names_f[i]])
    lbls <- paste(names(mytable), "\n", mytable, sep="")
    pie(mytable, labels = lbls, 
        main = paste("Pie Chart of", Names_f[i], sep = " "))
    dev.off()

    addImage(wb, filename = "piec.png", name = piec, originalSize = TRUE)
    
    saveWorkbook(wb)
    
    removeName(wb, piec)
    
    if(Dep.var.status == "Yes"){
      if(Dep.Var.Type == "continuous"){
        combinedata <- cbind(dep.data, mydata)
        names(combinedata)[1] <- "Dep.Var"
        names(combinedata)[2] <- "Group"
        group <- group_by(combinedata, Group)
        Analysis <- as.data.frame(summarise(group, mean(Dep.Var), median(Dep.Var), 
                                              sd(Dep.Var)))
        Analysis.N <- Assignnames(x = c(Names_f[i], "Mean", "Median", "SD"), data = Analysis)
        Analysis.N$Count <- table(combinedata$Group)
        for( j in 1:nrow(Analysis.N)){
          Analysis.N$Count.Percentage[j] <- percent((Analysis.N$Count[j]/nrow(data)))
        }
        ExportExcel.Specified.Location(data = Analysis.N, files = files, sheetNames = Names_f[i], 
                                       8, 12)
      } else if(Dep.Var.Type == "binary") {
        
        combinedata <- cbind(dep.data, mydata)
        names(combinedata)[1] <- "var1"
        names(combinedata)[2] <- "Var2"
        
        wb = loadWorkbook(files, create=FALSE)
        
        barc <- paste(Names_f[i],'barc' , sep = '.')
        formula.barc <- paste(Names_f[i],'!$M$8', sep = '')
        
        createName(wb, name = barc, formula = formula.barc)
        
        counts <- table(combinedata$var1, combinedata$Var2)
        png(filename = "barc.png", width = 500, height = 400)
        barplot(counts, main = paste("Bar Chart of", Names_f[i], "Vs", Dep.var, sep = " "),
                xlab = Names_f[i], col=c("darkblue","red"),
                legend = rownames(counts), border = TRUE, legend.text = TRUE)
        dev.off()
        
        addImage(wb, filename = "barc.png", name = barc, originalSize = TRUE)
        
        saveWorkbook(wb)
        
        removeName(wb, barc)
      }
    }
  }
}


Assignnames <- function(x, data){
  for(i in 1:length(x)){
    names(data)[i] <- x[i]
  }
  return(data)
}

AnalysisForEachDateVariable <- function(data, files){
  if (ncol(data) != 0){
    Names_f <- Getnames(data = data)
    for (i in 1:ncol(data)){
      mydata <- data[Names_f[i]]
      names(mydata)[1] <- "Date"
      mydata$Date <-as.Date(mydata$Date, format = c("%d-%m-%y"))
      Names <- Names_f[i]
      Type <- class(mydata$Date)
      NoofObservation <- length(mydata$Date)
      Missing_Count <- sum(is.na(mydata$Date))
      Missing_Percentage <- percent((Missing_Count/NoofObservation))
      Min <- min(mydata$Date)
      Max <- max(mydata$Date)
      data.f <- data.frame(Names, Type, Missing_Percentage, Min, Max)
      ExportExcel(data = data.f, files = files, sheetNames = Names_f[i])
    }
  } else {
    return(NULL)
  }
}



DataExpResult <- function(data, files, Dep.var, Dep.var.status, Dep.Var.Type, drops, status){
  
  #Remove the duplicates in the data
  data <- RemoveDuplicates(data = data)
  
  #Creating the excel file
  wb = loadWorkbook(files, create=TRUE)
  
  #Getting information about all variables & exporting to excel
  VI <- VariablesInformation(data = data)
  ExportExcel(data = VI, files = files, sheetNames = "Over All")
  
  #Removing unused variables
  if(status){
    data <- data[ , !(names(data) %in% drops)]
  }
  
  #Getting statiscal summary for all numerical variables & exporting to excel
  num <- separateVariables(data = data)[[1]]
  SNV <- SummaryNumericalVariables(data = num)
  ExportExcel(data = SNV, files = files, sheetNames = "Statiscal Summary")
  
  #Correlation Matrix
  CM <- CorrelationMatrix(data = num, method = "number")
  ExportExcel(data = CM, files = files, sheetNames = "Correlation Matrix")
  
  #Analysis for each numeric variables
  AnalysisForEachNumericVariable(num = num, files = files, Dep.var, Dep.var.status, Dep.Var.Type)
  
  
  #Analysis for each character variables
  char <- separateVariables(data = data)[[2]]
  AnalysisForEachCharacterVariable(char = char, files = files, Dep.var, Dep.var.status, 
                                   Dep.Var.Type)

}

