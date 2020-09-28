# http://www.zendei.com/article/2409.html
#�@�ΡGŪ���C�Ӥ�󧨤U��excel�A�ñN��X�֦��@�Ӥ��C
#�@��3�šG�Ĥ@�šG��󧨦W�A�ĤG�šG��󧨤���xlsx���W�A�ĤT�šGxlsx��󤤪��C��
#�N�X�᭱�������Y���G�i�ק�j�r�ˡA�h���ܦp�G�n�b�A�����W�B��Ӭq�N�X�ɡA�ݭn�i��������ק�C

##########��k�@�G�̲׳�W�O�s�b�C�Ӥ�󧨤U
rm(list=ls())
# setwd("E:/cnblogs")                                    #�]�w�u�@�ؿ��i�ק�j
setwd(getwd())
PathName = setwd(getwd())


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


FolderName = c("Metastasis") 


Main_1 <- as.data.frame(read_excel(paste0(PathName,"/",MainFileName,".xlsx"),"Enrichment"))
Main_1 <- data.table(Main_1)

first_category_name = list.files(FolderName)            #list.files�R�O�o��"APP��z����󧨤U�Ҧ���󧨪��W�١i�ק�j
dir = paste("/",FolderName,"/",first_category_name,sep="")   #��paste�R�O�c�ظ��|�ܼ�dir,�Ĥ@�ťؿ����ԲӸ��|�i�ק�j
n = length(dir)                                       #Ū��dir���סA�]�N�O�G�`�@���h�֭Ӥ@�ťؿ�                                                     

##########
Intersect_1 <- c()
CountInterset <- c()
SubFileName <- c()

for(i in 1:n){         #���C�Ӥ@�ťؿ�(���)
  b=list.files(dir[i]) #b�O�C�X�C�Ӥ@�ťؿ�(���)���C��xlsx��󪺦W��

  
  Sub_1 <- as.data.frame(read_excel(paste0(PathName,dir[i],sep=""),"Enrichment"))
  Sub_1 <- data.table(Sub_1)
  SubFileName[i] <- first_category_name[i]
  SubFileName2 <- first_category_name[i]

  # https://www.zhihu.com/question/39755875  
  Intersect_1<- semi_join(Main_1,Sub_1,by="Term")
  Intersect_1<- unique(Intersect_1, by = "Term")
#  dim(Intersect_1)
  
  write_xlsx(list(Intersec = Intersect_1,Main = Main_1,Sub = Sub_1),paste0(PathName,"/Intersec_",MainFileName,"_AND_",SubFileName2))
  
  CountInterset[i] <- nrow(Intersect_1)
#  SumTable[i] <- c(SubFileName,CountInterset)
#  SubFileName[i] <- SubFileName
 
}

SumTable <- as.data.frame(cbind(SubFileName,CountInterset))
write_xlsx(SumTable,paste0(PathName,"/Intersec_",MainFileName,"_SumTable.xlsx"))


####################################################################################################################################################################################