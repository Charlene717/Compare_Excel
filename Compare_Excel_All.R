# http://www.zendei.com/article/2409.html
#作用：讀取每個文件夾下的excel，並將其合併成一個文件。
#共有3級：第一級：文件夾名，第二級：文件夾中的xlsx文件名，第三級：xlsx文件中的每行
#代碼後面註釋中若有：【修改】字樣，則表示如果要在你機器上運行該段代碼時，需要進行相應的修改。

##########方法一：最終單獨保存在每個文件夾下
rm(list=ls())
# setwd("E:/cnblogs")                                    #設定工作目錄【修改】
setwd(getwd())
PathName = setwd(getwd())

FolderName = c("Metastasis") 
dir.create(paste0(PathName,"/",FolderName,"_Result_All"))

MainFileName = c("LGG XIAP Positive")

library(xlsx)
# https://www.cnblogs.com/loca/p/4682568.html
# https://earthworm2016.pixnet.net/blog/post/320495233-r-%E8%B3%87%E6%96%99%E5%8C%AF%E5%87%BA-excel%E6%AA%94
# https://stackoverflow.com/questions/55185436/rjava-non-zero-exit-status

########## Try another method to read Excel
# https://readxl.tidyverse.org/
# install.packages("tidyverse")
# install.packages("readxl")
library(readxl)


########## Load the file
#Main <- read.xlsx(file = MainFileName,sheetIndex=Enrichment,startRow = 3,endRow = 26,header = F,colIndex =3:5,encoding = "UTF-8")

#read_excel('test.xlsx',1) ##1 is the position of the sheet, it can also enter the name of a sheet. 

library("data.table")
library("dplyr")
# https://www.zhihu.com/question/39755875

# https://tutorials.methodsconsultants.com/posts/reading-and-writing-excel-files-with-r-using-readxl-and-writexl/
library(writexl)




######### Loop

Main_1 <- as.data.frame(read_excel(paste0(PathName,"/",MainFileName,".xlsx"),"Enrichment"))
Main_1 <- data.table(Main_1)

first_category_name = list.files(FolderName)             #list.files命令得到"APP整理”文件夾下所有文件夾的名稱【修改】
dir = paste("/",FolderName,"/",first_category_name,sep="")    #用paste命令構建路徑變數dir,第一級目錄的詳細路徑【修改】
n = length(dir)                                       #讀取dir長度，也就是：總共有多少個一級目錄                                                   

##########
Intersect_1 <- c()
CountInterset <- c()
SubFileName <- c()
merge_1 <- as.data.frame(matrix(nrow=1,ncol=9))
names(merge_1)<-c('GroupID','Category','Term','Description','LogP','Log(q-value)','InTerm_InList','Genes','Symbols')

for(i in 1:n){         #對於每個一級目錄(文件夾)
  b=list.files(dir[i]) #b是列出每個一級目錄(文件夾)中每個xlsx文件的名稱

  
  Sub_1 <- as.data.frame(read_excel(paste0(PathName,dir[i],sep=""),"Enrichment"))
  Sub_1 <- data.table(Sub_1)
  SubFileName[i] <- first_category_name[i]
  SubFileName2 <- first_category_name[i]

  # https://www.zhihu.com/question/39755875  
  Intersect_1<- semi_join(Main_1,Sub_1,by = c("Term"="Term","Description"="Description"))
  Intersect_1<- unique(Intersect_1, by = c("Term"="Term","Description"="Description"))
#  dim(Intersect_1)
  names(Intersect_1)<-c('GroupID','Category','Term','Description','LogP','Log(q-value)','InTerm_InList','Genes','Symbols')
  
  write_xlsx(list(Intersec = Intersect_1,Main = Main_1,Sub = Sub_1),paste0(PathName,"/",FolderName,"_Result_All","/Intersec_",MainFileName,"_AND_",SubFileName2))
  
  CountInterset[i] <- nrow(Intersect_1)
  merge_1 <- rbind(merge_1,Intersect_1)
  
#  SumTable[i] <- c(SubFileName,CountInterset)
#  SubFileName[i] <- SubFileName
 
}
merge_1 <- merge_1[!is.na(merge_1$Term),]
merge_2<- unique(merge_1)
#merge_2<- unique(merge_1, by = c("Term"="Term","Description"="Description"))

SumTable <- as.data.frame(cbind(SubFileName,CountInterset))
write_xlsx(list(Count = SumTable,merge = merge_2,merge_ori = merge_1),paste0(PathName,"/",FolderName,"_Result_All","/Intersec_",MainFileName,"_SumTable.xlsx"))


####################################################################################################################################################################################
