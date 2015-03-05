lft <- "https://www.tesouro.fazenda.gov.br/documents/10180/137713/LFT_2015.xls"
ltn <- "https://www.tesouro.fazenda.gov.br/documents/10180/137713/LTN_2015.xls"
ntnbp <- "https://www.tesouro.fazenda.gov.br/documents/10180/137713/NTN-B_Principal_2015.xls"

setwd("C://Users//sandor.caetano//Google Drive//R//TD//")

lista <- list(tits = c(lft,ltn,ntnbp))

library(xlsx)

Res <- NULL


for (ii in 1:length(lista[[1]])){
        temp <- tempfile()
        t <- download.file(lista[[1]][ii],temp, mode="wb")
        
        #carrega as sheets da pasta
        t <- loadWorkbook(temp) 
        sheets <- names(getSheets(t))
        
        
        for (i in 1:length(sheets)){
                jj <- read.xlsx(temp,i,startRow=2)        
                P <- jj$PU.Venda.ManhÃ[length(jj$PU.Venda.ManhÃ)]
                Data <- as.Date(jj$Dia[length(jj$Dia)],"%d/%m/%Y")
                Res <- rbind(c(sheets[i], as.character(Data),P),Res)
        }
        
}

Res



#wb <- createWorkbook()
#.jmethods(wb)

#dLFT <- "LFT_2015.xls"
#dLTN <- "LTN_2015.xls"
#dNTN <- "NTNBP_2015.xls"

#download.file(lft,dLFT, mode="wb")
#download.file(ltn,dLTN, mode="wb")
#download.file(ntnbp,dNTN, mode="wb")
