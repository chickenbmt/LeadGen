cat("\014")
rm(list=ls())
library(readxl)
library(data.table)
library(WriteXLS)

#-------------------------------------------------------------------
#BATCH REVOLUTION FUNNEL
#-------------------------------------------------------------------

path.old = getwd()
setwd(file.path(path.old,"Data/FromBank"))

file.list = list.files(pattern = "\\.xlsx$")
file.date = gsub("_sent.xlsx$","",file.list)
file.date = gsub("^.*_","",file.date)

#Daily Funnel
sheet.name = "Daily Funnel"
skip.row = 2

excel.batch.template = as.data.frame(matrix(0,ncol = 0,nrow=29))
excel.batch.template$RowNames = c("BatchSentDate","ReportDate","Lead received","Not Called Yet",
                                   "Called","Not Contacted","Follow 1st time","Follow 2nd time",
                                   "Follow 3rd time","Other","Contacted",
                                   "Not Interested","Cancelled","Interested",
                                   "CIC","Not Eligible","Loan/app at FE/VPB",
                                   "Processing","Meeting","Processing & Meeting","APP",
                                   "IN PROCESS","REJECT","CANCEL",
                                   "APPROVED","Waiting for cust","Document checking (PDOC)",
                                   "Document checking (DOV)","Disbursed")
excel.batch.append = as.data.frame(matrix(0,ncol = 29,nrow=0))
colnames(excel.batch.append) = excel.batch.template$RowNames

for (i in 1:length(file.list)){
  a = excel_sheets(file.list[i])
  if(sheet.name %in% a){
  exceltemp = read_excel(file.list[i], sheet = sheet.name, col_names=TRUE,skip=skip.row)
  exceltemp = exceltemp[!is.na(exceltemp[,1]),]
  exceltemp = exceltemp[,c(1,2,grep("2016*",colnames(exceltemp)))]
  exceltemp = merge(excel.batch.template,exceltemp,sort = FALSE,all=TRUE,by=1)
  row.names(exceltemp) = exceltemp[,1]
  exceltemp = exceltemp[,-c(1,2)]
  exceltemp = as.data.frame(t(exceltemp))
  exceltemp$BatchSentDate = row.names(exceltemp)
  exceltemp$ReportDate = file.date[i]
  exceltemp = exceltemp[,excel.batch.template$RowNames]
  excel.batch.append = rbind(excel.batch.append, exceltemp)
  rm(exceltemp)
  }
  rm(a)
} 

row.names(excel.batch.append) = NULL

excel.batch.append$`Processing & Meeting` = ifelse(is.na(excel.batch.append$`Processing & Meeting`),
                                                   excel.batch.append$Processing + excel.batch.append$Meeting,
                                                   excel.batch.append$`Processing & Meeting`)

#Not Eligible
sheet.name = "Not eligible"
skip.row = 0
excel.ne.append = as.data.frame(matrix(0,ncol = 0,nrow=6))
excel.ne.append[,1] = c("Total Non-Eligible","Unsatisfactory condition","Cannot supplement required document(s)",
                        "Not in supported location","CIC","Others")
colnames(excel.ne.append) = "Rownames"
for (i in 1:length(file.list)){
  a = excel_sheets(file.list[i])
  if(sheet.name %in% a){
    exceltemp = read_excel(file.list[i], sheet = sheet.name, col_names=F,skip=skip.row)
    colnames(exceltemp) = c("Rownames",file.date[i])
    excel.ne.append = merge(excel.ne.append,exceltemp,1,all.x=T,sort=F)
    rm(exceltemp)
  }
  rm(a)
}

setwd(path.old)
#----------------------------------------------------------------
#ANALYSE
#----------------------------------------------------------------
#Accumulated Data

data.acc = data.table(excel.batch.append)[,lapply(.SD, sum),by=ReportDate,
                                          .SDcols=c(colnames(excel.batch.append[,3:length(excel.batch.append)]))]
data.acc = as.data.frame(data.acc)
row.names(data.acc) = data.acc[,1]
data.acc = data.acc[,-1]

data.acc.per = data.acc/data.acc[,1]
data.acc.per = t(data.acc.per)
data.acc = as.data.frame(t(data.acc))
data.acc.per = as.data.frame(data.acc.per)

data.compare = data.acc.per[,(length(data.acc.per)-1):length(data.acc.per)]
data.compare$Delta = data.compare[,2]-data.compare[,1]
data.compare = data.compare[order(abs(data.compare$`Delta`),decreasing = TRUE),]

#Detail by features
data.detail = excel.batch.append[excel.batch.append$ReportDate%in%c(colnames(data.acc.per)[
  (length(data.acc.per)-1):length(data.acc.per)]),]
data.detail[,3:length(data.detail)] = data.detail[,3:length(data.detail)]/data.detail[,3]
data.detail.list = split(data.detail,data.detail$ReportDate)
data.detail = excel.batch.append[excel.batch.append$ReportDate%in%c(colnames(data.acc.per)[
                                                                      length(data.acc.per)]),]
data.detail[3:length(data.detail)] = data.detail[3:length(data.detail)]/data.detail[,3]

data.detail[1:(nrow(data.detail)-1),3:length(data.detail)] = data.detail.list[[2]][-nrow(data.detail.list[[2]]),3:length(data.detail.list[[2]])] - 
              data.detail.list[[1]][,3:length(data.detail.list[[1]])]
data.detail =data.detail[order(data.detail$BatchSentDate,decreasing = T),]
row.names(data.detail) = data.detail[,1]
data.detail = data.detail[,-(1:2)]
rm(data.detail.list)

#Not eligible
data.ne.compare = excel.ne.append[,c(colnames(data.acc.per))]
row.names(data.ne.compare) = excel.ne.append$Rownames
data.ne.compare$`Delta(last2days)` = data.ne.compare[,length(data.ne.compare)] - data.ne.compare[,length(data.ne.compare)-1]
data.ne.compare$`%Delta(last2days)` = round(data.ne.compare$`Delta(last2days)`/data.ne.compare[,(length(data.ne.compare)-1)],4)

#Follow up & Processing aging
data.aging = as.data.frame(matrix(0,ncol=0,nrow=7))
data.aging[,1] = c("1 day", "2 days", "3 days", "4 days", "5 days", "6-10 days", ">10 days")
a = excel.batch.append[excel.batch.append$ReportDate%in%colnames(data.acc)[(length(data.acc)-1):length(data.acc)],
                       c("BatchSentDate","ReportDate","Follow 1st time","Follow 2nd time","Follow 3rd time","Processing & Meeting","APP")]
colnames(a)[6] = "Processing & Meeting (not converted)"
a$cal = ifelse(sign(a[,6]-a[,7])!=-1,a[,6]-a[,7],0)
a$`Processing & Meeting (not converted)` = a$cal
a = a[,-c(7,8)]
data.list = split(a,a$ReportDate)
rm(a)

for (i in 1:2){
  temp = data.list[[i]]
  temp = temp[order(temp$BatchSentDate,decreasing = T),]
  temp = data.table(temp)
  temp.6.10 = temp[6:10,lapply(.SD,sum),.SD=c(colnames(temp)[3:length(temp)])]
  temp.11.end = temp[11:.N,lapply(.SD,sum),.SD=c(colnames(temp)[3:length(temp)])]
  temp1 = rbind(temp[1:5,.SD,.SD=c(colnames(temp)[3:length(temp)])],temp.6.10,temp.11.end)
  temp1$`Total Follow up` = temp1$`Follow 1st time`+temp1$`Follow 2nd time`+temp1$`Follow 3rd time`
  temp1$Lead_aging = c("1 day", "2 days", "3 days", "4 days", "5 days", "6-10 days", ">10 days")
  colnames(temp1) = paste(colnames(temp1),unique(temp$ReportDate))
  temp1 = as.data.frame(temp1)
  temp1 = temp1[,c(length(temp1),1:(length(temp1)-1))]
  data.aging = merge(data.aging,temp1,1,sort=F)
  rm(temp,temp1,temp.11.end,temp.6.10)
}

data.aging$Delta_1st = data.aging[,grep("1st",colnames(data.aging))[2]] - data.aging[,grep("1st",colnames(data.aging))[1]]
data.aging$Delta_2nd = data.aging[,grep("2nd",colnames(data.aging))[2]] - data.aging[,grep("2nd",colnames(data.aging))[1]]
data.aging$Delta_3rd = data.aging[,grep("3rd",colnames(data.aging))[2]] - data.aging[,grep("3rd",colnames(data.aging))[1]]
data.aging$Delta_PnM = data.aging[,grep("Processing",colnames(data.aging))[2]] - data.aging[,grep("Processing",colnames(data.aging))[1]]
data.aging$Delta_Follow_Up = data.aging[,grep("Total Follow",colnames(data.aging))[2]] - data.aging[,grep("Total Follow",colnames(data.aging))[1]]

data.aging = data.aging[,c(1,5,10,15,6,11,16,2,7,12,3,8,13,4,9,14)]
row.names(data.aging) = data.aging[,1]
data.aging = data.aging[,-1]

#Formating
data.acc.per = round(data.acc.per,digit=4)
data.compare = round(data.compare, digits = 4)
data.detail = round(data.detail, digits=4)

#WriteXLS
excel.file.name = paste0("TS_Daily_Summary_Report",".xlsx")
table.list = list(data.acc,data.acc.per,data.compare,data.detail,data.ne.compare,data.aging)
sheets.name = c("AccumulatedFunnel","Percetage","Delta","Delta_detail","Not_Eligible","Follow up & Processing Aging")

WriteXLS(table.list, excel.file.name, SheetNames =sheets.name, row.names=TRUE, col.names=TRUE,
         AdjWidth=TRUE, FreezeRow=1, FreezeCol=1)
