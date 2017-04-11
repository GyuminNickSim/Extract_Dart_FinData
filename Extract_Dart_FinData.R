###################################################################
## Extract Financial Data from DART
## Designed by: Hyunjin Jeong
## 
## 2017.04.08. (2017.04.11. Modified)
## Alpha 7.0
## 
## Please execute this code in R Studio
###################################################################

###################################################################
## Excute Function firstly
##
## StartExtract, DartJSONtoExcel 앞에 커서 두고 Ctrl+R 누르면 함수 실행됨 
###################################################################

## Main function to extract financial information data

StartExtract <- function () {
  
  Sys.setlocale("LC_ALL", "Korean_Korea.949")
  
  cat("Write down your DART API authentication code. You can get your own code from DART.")
  cat("\nIf you want to finish this program, press 'q!'\n")
  auth <- as.character(readline(prompt=""))
  
  if (auth!="q!") {
    cat("Write down company's code. (6 numbers)")
    cat("\nIf you want to finish this program, press 'q!'\n")
    companyCode <- as.character(readline(prompt=""))
    
    if (companyCode!="q!") {
      cat("Write down start date. (yyyymmdd) (Required to input date after 1st Apr, 2012)")
      cat("\nIf you want to finish this program, press 'q!'\n")
      startDate <- as.character(readline(prompt=""))
     
      if (startDate!="q!") {
        cat("Write down finish date. (yyyymmdd) (If you want to set finish date as today, you can just press [Enter].)")
        cat("\nIf you want to finish this program, press 'q!'\n")
        finishDate <- as.character(readline(prompt=""))
        
        if (finishDate!="q!") {
          if (finishDate=="" || finishDate==as.character(format(Sys.Date(), "%Y%m%d"))) {
            DartJSONtoExcel(auth, companyCode, startDate)
          } else {
            DartJSONtoExcel(auth, companyCode, startDate, finishDate)
          }
        }
      }
    }
  }
  
  Sys.setlocale()
}


## Sub function to post request, receive JSON file, and change to Excel csv file

DartJSONtoExcel <- function (auth, companyCode, startDate, finishDate=NULL) {
  
  #######################################################
  ## Install Packages
  
  pkgs <- c("readxl", "gdata", "plyr")
  tt <- sapply(pkgs, require, character.only=T)
  
  for (index in 1:length(tt)) {
    if (tt[index] == F) {
      install.packages(names(tt[index]))
    }
  }
  
  ##
  #######################################################
  
  if (as.numeric(startDate)>=20120401) {
    
    ## Set web address to download JSON file
    Prototype_Auth = paste("http://dart.fss.or.kr/api/search.json?auth=", auth, sep="")
    Prototype_Company = paste("crp_cd=", companyCode, sep="")
    Prototype_startDate = paste("start_dt=", startDate, sep="")
    
    if (is.null(finishDate)==F) {
      Prototype_finishDate = paste("end_dt=", finishDate, sep="")
      url_Prototype = paste(Prototype_Auth, Prototype_Company, Prototype_startDate, Prototype_finishDate, "bsn_tp=A001&bsn_tp=A002&bsn_tp=A003&bsn_tp=F002&fin_rpt=Y&page_set=100", sep="&")
      
      if (as.numeric(substr(startDate,5,8))<1001) {
        startdate_temp <- paste("start_dt=",paste(as.numeric(substr(startDate,1,4))-1,substr(startDate,5,8),sep=""),sep="")
        url_temp = paste(Prototype_Auth, Prototype_Company, startdate_temp, Prototype_finishDate, "bsn_tp=A001&bsn_tp=A002&bsn_tp=A003&bsn_tp=F002&fin_rpt=Y&page_set=100", sep="&")
      }
      
    } else {
      Prototype_finishDate = NULL
      url_Prototype = paste(Prototype_Auth, Prototype_Company, Prototype_startDate, "bsn_tp=A001&bsn_tp=A002&bsn_tp=A003&bsn_tp=F002&fin_rpt=Y&page_set=100", sep="&")
      
      if (as.numeric(substr(startDate,5,8))<1001) {
        startdate_temp <- paste("start_dt=",paste(as.numeric(substr(startDate,1,4))-1,substr(startDate,5,8),sep=""),sep="")
        url_temp = paste(Prototype_Auth, Prototype_Company, startdate_temp, "bsn_tp=A001&bsn_tp=A002&bsn_tp=A003&bsn_tp=F002&fin_rpt=Y&page_set=100", sep="&")
      }
    }
    
    ## Download JSON file from DART
    fjson <- jsonlite::fromJSON(url_Prototype)
    fjson_temp <- jsonlite::fromJSON(url_temp)
    
    ## Make folder to store your data
    mainDir <- "D:\\"
    subDir <- "DartDownload"
    
    setwd("D:\\")
    if (file.exists(subDir)){
      setwd(file.path(mainDir, subDir))
    } else {
      dir.create(file.path(mainDir, subDir))
      setwd(file.path(mainDir, subDir))
    }
    
    mainDir2 <- "D:\\DartDownload"
    subDir2 <- "Data"
    
    setwd("D:\\")
    if (file.exists(subDir2)){
      setwd(file.path(mainDir2, subDir2))
    } else {
      dir.create(file.path(mainDir2, subDir2))
      setwd(file.path(mainDir2, subDir2))
    }
    
    mainDir2_1 <- "D:\\DartDownload\\Data"
    subDir2_1 <- companyCode
    
    setwd("D:\\")
    if (file.exists(subDir2_1)){
      setwd(file.path(mainDir2_1, subDir2_1))
    } else {
      dir.create(file.path(mainDir2_1, subDir2_1))
      setwd(file.path(mainDir2_1, subDir2_1))
    }
    
    if (as.numeric(substr(startDate,5,8))<=1001) {
      LEN = length(fjson_temp$list$rcp_no)
      LEN3 = LEN
      LEN2 = length(fjson$list$rcp_no)
    } else {
      LEN = length(fjson$list$rcp_no)
      LEN3 = LEN
      LEN2 = LEN
    }
    
    temp <- NULL
    
    ## Set web address to download Excel file about Financial Information
    for(i in 1:LEN) {
      if (i<=LEN-(LEN3-LEN2)) {
        down_Prototype <- paste("http://dart.fss.or.kr/pdf/download/excel.do?rcp_no=", fjson_temp$list$rcp_no[i], "&lang=ko", sep="")
        
        FSdata_address <- paste(
          companyCode,
          substr(fjson$list$rcp_no[i],1,8),
          "FSdata.xls",
          sep="_"
        )
        
        ### Download Excel file
        download.file(down_Prototype, FSdata_address, method="internal", mode="wb")
      }
      
      if (as.numeric(substr(startDate,5,8))<=1001 && i>LEN-(LEN3-LEN2)) {
        down_temp <- paste("http://dart.fss.or.kr/pdf/download/excel.do?rcp_no=", fjson_temp$list$rcp_no[i], "&lang=ko", sep="")
        
        FSdata_address <- paste(
          companyCode,
          substr(fjson_temp$list$rcp_no[i],1,8),
          "FSdata.xls",
          sep="_"
        )
        
        temp[i] <- FSdata_address
        
        ### Download Excel file
        download.file(down_temp, FSdata_address, method="internal", mode="wb")
        
      } else if (as.numeric(substr(startDate,5,8))>1001 && i>=LEN-3) {
        down_Prototype <- paste("http://dart.fss.or.kr/pdf/download/excel.do?rcp_no=", fjson$list$rcp_no[i], "&lang=ko", sep="")
        
        FSdata_address <- paste(
          companyCode,
          substr(fjson$list$rcp_no[i],1,8),
          "FSdata.xls",
          sep="_"
        )
        
        ### Download Excel file
        download.file(down_Prototype, FSdata_address, method="internal", mode="wb")
      }
      
      if (file.size(FSdata_address)>5000 && file.exists(FSdata_address)==T) {
        
        ### Financial Statement
        if (i!=1) FS_bind <- FS_work
        
        for (j in 1:7) {
          if (is.na(regexpr("단위", readxl::read_excel(FSdata_address, 2)[j,1])[1])==F &&
              regexpr("단위", readxl::read_excel(FSdata_address, 2)[j,1])[1]!=-1) j2=j;
        }
        FS <- readxl::read_excel(FSdata_address, 2)[-c(1:j2-1),c(1:2)]
        
        K <- gdata::trim(substr(FS[1,1], regexpr(":", FS[1,1])[1]+1, regexpr(")", FS[1,1])[1]-1))
        
        if (K=="원") {
          G=1
        } else if (K=="천원") {
          G=1000
        } else if (K=="백만원") {
          G=1000000
        } else {
          G=1000000000
        }
        
        FS <- FS[-1,]
        
        FS <- subset(FS,FS[,1]!="자산" & FS[,1]!="부채" & FS[,1]!="자본")
        FS[1,1] <- "Date"
        
        #### Transpose Financial Statement matrix
        Transpose_FS <- as.data.frame(t(as.matrix(FS)), stringsAsFactors=F)
        names(Transpose_FS) <- gdata::trim(as.character(unlist(Transpose_FS[1,])))
        Transpose_FS <- Transpose_FS[-1,]
        Transpose_FS[,2:ncol(Transpose_FS)] <- apply(Transpose_FS[,2:ncol(Transpose_FS)],2, FUN=as.numeric) * G
        
        if (as.numeric(substr(startDate,5,8))<=1001 && i>LEN-(LEN3-LEN2)) {
          Transpose_FS$index=1;
        } else {
          Transpose_FS$index=0;
        }
        
        #### Bind quarterly Financial Statement data
        if (i!=1) {
          FS_work <- plyr::rbind.fill(Transpose_FS, FS_bind)
        } else {
          FS_work <- Transpose_FS
        }
        
        FS_work <- FS_work[!duplicated(FS_work["Date"]),]
        
        
        ### Income Statement
        if (i!=1) IS_bind <- IS_work
        
        for (j in 1:7) {
          if (is.na(regexpr("단위", readxl::read_excel(FSdata_address, 3)[j,1])[1])==F &&
              regexpr("단위", readxl::read_excel(FSdata_address, 3)[j,1])[1]!=-1) j2=j;
        }
        IS <- readxl::read_excel(FSdata_address, 3)[-c(1:j2-1),c(1:2)]
        IS <- subset(IS, is.na(IS[,1])==F)
        IS[,1] <- gsub(x=IS[,1], "\\d+", "")
        
        K2 <- gdata::trim(substr(IS[1,1], regexpr(":", IS[1,1])[1]+1, regexpr(")", IS[1,1])[1]-1))
        
        
        if (K2=="원") {
          G2=1
        } else if (K2=="천원") {
          G2=1000
        } else if (K2=="백만원") {
          G2=1000000
        } else {
          G2=1000000000
        }
        
        IS <- IS[-1,]
        
        IS <- subset(IS,IS[,2]!="3개월")
        
        IS[1,1] <- "Date"
        
        #### Transpose Income Statement matrix
        Transpose_IS <- as.data.frame(t(as.matrix(IS)), stringsAsFactors=F)
        names(Transpose_IS) <- gdata::trim(as.character(unlist(Transpose_IS[1,])))
        
        #### Standardize variable names 
        for (k in 1:length(names(Transpose_IS))) {
          if (is.na(names(Transpose_IS)[k])==F) {
            if (names(Transpose_IS)[k]=="수익(매출액)") {
              names(Transpose_IS)[k] = "매출액"
            } else if (names(Transpose_IS)[k]=="매출총이익(손실)") {
              names(Transpose_IS)[k] = "매출총이익"
            } else if (names(Transpose_IS)[k]=="영업이익(손실)") {
              names(Transpose_IS)[k] = "영업이익"
            } else if (names(Transpose_IS)[k]=="법인세비용차감전순이익(손실)") {
              names(Transpose_IS)[k] = "법인세비용차감전순이익"
            } else if (names(Transpose_IS)[k]=="계속영업이익(손실)") {
              names(Transpose_IS)[k] = "계속영업이익"
            } else if (names(Transpose_IS)[k]=="당기순이익(손실)" || 
                       names(Transpose_IS)[k]=="반기순이익" || 
                       names(Transpose_IS)[k]=="반기순이익(손실)" ||
                       names(Transpose_IS)[k]=="분기순이익" ||
                       names(Transpose_IS)[k]=="분기순이익(손실)") {
              names(Transpose_IS)[k] = "당기순이익"
            } else if (names(Transpose_IS)[k]=="당기순이익(손실)의 귀속" ||
                       names(Transpose_IS)[k]=="반기순이익의 귀속" ||
                       names(Transpose_IS)[k]=="반기순이익(손실)의 귀속" ||
                       names(Transpose_IS)[k]=="분기순이익의 귀속" ||
                       names(Transpose_IS)[k]=="분기순이익(손실)의 귀속") {
              names(Transpose_IS)[k] = "당기순이익의 귀속"
              
              ##### 아래 2가지의 경우 변수 통일이 사실상 불가능 
            } else if (names(Transpose_IS)[k]=="지배기업의 소유주에게 귀속되는 당기순이익(손실)") {
              names(Transpose_IS)[k] = "지배기업의 소유주에게 귀속되는 당기순이익"
            } else if (names(Transpose_IS)[k]=="비지배지분에 귀속되는 당기순이익(손실)") {
              names(Transpose_IS)[k] = "비지배지분에 귀속되는 당기순이익"
              
              ##### 변수명에 단위가 있을 수 있음  
            } else if (names(Transpose_IS)[k]=="기본주당이익(손실)" || names(Transpose_IS)[k]=="기본주당이익(손실) (단위 : 원)") {
              names(Transpose_IS)[k] = "기본주당이익"
            } else if (names(Transpose_IS)[k]=="희석주당이익(손실)" || names(Transpose_IS)[k]=="희석주당이익(손실) (단위 : 원)") {
              names(Transpose_IS)[k] = "희석주당이익"
            }
          }
        }
        
        Transpose_IS <- Transpose_IS[-1,]
        Transpose_IS[,2:ncol(Transpose_IS)] <- apply(Transpose_IS[,2:ncol(Transpose_IS)],2, FUN=as.numeric) * G2
        
        if (as.numeric(substr(startDate,5,8))<=1001 && i>LEN-(LEN3-LEN2)) {
          Transpose_IS$index=1;
        } else {
          Transpose_IS$index=0;
        }
        
        #### Bind quarterly Financial Statement data
        if (i!=1) {
          IS_work <- plyr::rbind.fill(Transpose_IS, IS_bind)
        } else {
          IS_work <- Transpose_IS
        }
        
        IS_work <- IS_work[!duplicated(IS_work["Date"]),]
        
        
        ### Cash Flow Statement
        if (i!=1) CFS_bind <- CFS_work
        
        cfs_index=0
        for (j in 1:length(readxl::excel_sheets(FSdata_address))) {
          if (readxl::excel_sheets(FSdata_address)[j]=="연결 현금흐름표") {
            cfs_index=cfs_index+1
            
          }
        }
        
        if (cfs_index>0) {
          for (j in 1:8) {
            if (is.na(regexpr("단위", readxl::read_excel(FSdata_address, "연결 현금흐름표")[j,1])[1])==F &&
                regexpr("단위", readxl::read_excel(FSdata_address, "연결 현금흐름표")[j,1])[1]!=-1) j2=j;
          }
          CFS <- readxl::read_excel(FSdata_address, "연결 현금흐름표")[-c(1:j2-1),c(1:2)]
        } else {
          for (j in 1:8) {
            if (is.na(regexpr("단위", readxl::read_excel(FSdata_address, "현금흐름표")[j,1])[1])==F &&
                regexpr("단위", readxl::read_excel(FSdata_address, "현금흐름표")[j,1])[1]!=-1) j2=j;
          }
          CFS <- readxl::read_excel(FSdata_address, "현금흐름표")[-c(1:j2-1),c(1:2)]
        }
  
        K3 <- gdata::trim(substr(CFS[1,1], regexpr(":", CFS[1,1])[1]+1, regexpr(")", CFS[1,1])[1]-1))
        
        
        if (K3=="원") {
          G3=1
        } else if (K2=="천원") {
          G3=1000
        } else if (K2=="백만원") {
          G3=1000000
        } else {
          G3=1000000000
        }
        
        CFS <- CFS[-1,]
        
        CFS[1,1] <- "Date"
        
        #### Transpose Income Statement matrix
        Transpose_CFS <- as.data.frame(t(as.matrix(CFS)), stringsAsFactors=F)
        names(Transpose_CFS) <- gdata::trim(as.character(unlist(Transpose_CFS[1,])))
        names(Transpose_CFS) <- gsub(pattern=" ", x=names(Transpose_CFS), replacement="")
        
        #### Standardize variable names 
        for (k in 1:length(names(Transpose_CFS))) {
          if (is.na(names(Transpose_CFS)[k])==F) {
            if (names(Transpose_CFS)[k]=="단기차입금의순증가(감소)") {
              names(Transpose_CFS)[k] = "단기차입금의순증가"
            } else if (names(Transpose_CFS)[k]=="단기금융상품의순감소(증가)") {
              names(Transpose_CFS)[k] = "단기금융상품의순감소"
            } else if (names(Transpose_CFS)[k]=="현금및현금성자산의순증가(감소)") {
              names(Transpose_CFS)[k] = "현금및현금성자산의순증가"
            } else if (names(Transpose_CFS)[k]=="환율변동효과반영전현금및현금성자산의순증가(감소)") {
              names(Transpose_CFS)[k] = "환율변동효과반영전현금및현금성자산의순증가"
            } else if (names(Transpose_CFS)[k]=="단기차입금의순차입(상환)") {
              names(Transpose_CFS)[k] = "단기차입금의순차입"
            }
          }
        }
        
        Transpose_CFS <- Transpose_CFS[-1,]
        Transpose_CFS[,2:ncol(Transpose_CFS)] <- apply(Transpose_CFS[,2:ncol(Transpose_CFS)],2, FUN=as.numeric) * G3
        
        if (as.numeric(substr(startDate,5,8))<=1001 && i>LEN-(LEN3-LEN2)) {
          Transpose_CFS$index=1;
        } else {
          Transpose_CFS$index=0;
        }
        
        #### Bind quarterly Financial Statement data
        if (i!=1) {
          CFS_work <- plyr::rbind.fill(Transpose_CFS, CFS_bind)
        } else {
          CFS_work <- Transpose_CFS
        }
        
        CFS_work <- CFS_work[!duplicated(CFS_work["Date"]),]
        
      } else {
        file.remove(FSdata_address)
      }
    }
    
    ## Calculate Quarter Income
    for (j in 5:nrow(IS_work)) {
      if (nchar(IS_work[j,1])<=8) {
        IS_work[j,2:ncol(IS_work)] = IS_work[j,2:ncol(IS_work)] - IS_work[j-1,2:ncol(IS_work)] - IS_work[j-2,2:ncol(IS_work)] - IS_work[j-3,2:ncol(IS_work)]
      }
    }
    
    ## Calculate Quarter Cash Flow
    substrRight <- function(x, n){
      substr(x, nchar(x)-n+1, nchar(x))
    }
    
    for (j in 5:nrow(CFS_work)) {
      if (substrRight(CFS_work$Date[1], 3)!="1분기") {
        CFS_work[j,2:ncol(CFS_work)] = CFS_work[j,2:ncol(CFS_work)] - CFS_work[j-1,2:ncol(CFS_work)]
      }
    }
    
    ## Remove Temperary Files
    for (ii in 1:length(subset(temp, is.na(temp)==F))) {
      file.remove(subset(temp, is.na(temp)==F)[ii])
    }
    
    
    ## Make folder to store your data
    subDir3 <- "FSData"
    subDir4 <- "ISData"
    subDir5 <- "CFSData"
    
    
    ## Write FS Data
    setwd("D:\\")
    if (file.exists(subDir3)){
      setwd(file.path(mainDir2, subDir3))
    } else {
      dir.create(file.path(mainDir2, subDir3))
      setwd(file.path(mainDir2, subDir3))
    }
    
    if (is.null(finishDate)==F) {
      FSdata_address2 <- paste(
        companyCode,
        startDate,
        finishDate,
        "Financial_Statement_Quarterly.csv",
        sep="_"
      )
    } else {
      FSdata_address2 <- paste(
        companyCode,
        startDate,
        format(Sys.Date(), "%Y%m%d"),
        "Financial_Statement_Quarterly.csv",
        sep="_"
      )
    }
    
    if (as.numeric(substr(startDate,5,8))<1001) {
      FS_work <- subset(FS_work, index==0)
    }
    FS_work <- FS_work[,-which(names(FS_work) %in% "index")]
    
    write.csv(FS_work, FSdata_address2, row.names=F, na="")
    
    
    ## Write IS Data
    setwd("D:\\")
    if (file.exists(subDir4)){
      setwd(file.path(mainDir2, subDir4))
    } else {
      dir.create(file.path(mainDir2, subDir4))
      setwd(file.path(mainDir2, subDir4))
    }
    
    if (is.null(finishDate)==F) {
      FSdata_address3 <- paste(
        companyCode,
        startDate,
        finishDate,
        "Income_Statement_Quarterly.csv",
        sep="_"
      )
    } else {
      FSdata_address3 <- paste(
        companyCode,
        startDate,
        format(Sys.Date(), "%Y%m%d"),
        "Income_Statement_Quarterly.csv",
        sep="_"
      )
    }
    
    if (as.numeric(substr(startDate,5,8))<1001) {
      IS_work <- subset(IS_work, index==0)
    }
    IS_work <- IS_work[,-which(names(IS_work) %in% "index")]
    
    for (k in 1:length(names(IS_work))) {
      if (names(IS_work)[k]=="기본주당이익" || names(IS_work)[k]=="희석주당이익") {
        IS_work[,k] <- IS_work[,k] / G2
      }
    }
    
    write.csv(IS_work, FSdata_address3, row.names=F, na="")
    
    
    ## Write CFS Data
    setwd("D:\\")
    if (file.exists(subDir5)){
      setwd(file.path(mainDir2, subDir5))
    } else {
      dir.create(file.path(mainDir2, subDir5))
      setwd(file.path(mainDir2, subDir5))
    }
    
    if (is.null(finishDate)==F) {
      FSdata_address4 <- paste(
        companyCode,
        startDate,
        finishDate,
        "Cash_Flow_Statement_Quarterly.csv",
        sep="_"
      )
    } else {
      FSdata_address4 <- paste(
        companyCode,
        startDate,
        format(Sys.Date(), "%Y%m%d"),
        "Cash_Flow_Statement_Quarterly.csv",
        sep="_"
      )
    }
    
    if (as.numeric(substr(startDate,5,8))<1001) {
      CFS_work <- subset(CFS_work, index==0)
    }
    CFS_work <- CFS_work[,-which(names(CFS_work) %in% "index")]
    
    write.csv(CFS_work, FSdata_address4, row.names=F, na="")
  
  } else {
    message("Please input start date from 1st April, 2012.")
  }
}


###################################################################
## Excution Code
##
## You just execute StartExtract().
###################################################################

StartExtract()
