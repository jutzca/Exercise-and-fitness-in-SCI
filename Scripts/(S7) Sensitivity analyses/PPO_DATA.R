install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/SENS/'
wdir <- '~/Documents/R meta-analysis/SENS_DATA/'
setwd(data_dir)

PPODATA <- read.xlsx("SENS_DATA.xlsx", sheet = 6)

PPODATA$n.post <- as.numeric(PPODATA$n.post)
PPODATA$mean.post <- as.numeric(PPODATA$mean.post)
PPODATA$sd.post <- as.numeric(PPODATA$sd.post)
PPODATA$n.pre <- as.numeric(PPODATA$n.pre)
PPODATA$mean.pre <- as.numeric(PPODATA$mean.pre)
PPODATA$sd.pre <- as.numeric(PPODATA$sd.pre)

# Naming subgroups
PPODATA$subgroup <- as.factor(PPODATA$subgroup)

# Effect size calculation
PPODATA_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=PPODATA,var.names=c('PPODATA_MD','PPODATA_VAR'))

# Overall random effects model
PPODATA<- rma(PPODATA_MD, PPODATA_VAR, data=PPODATA_ESCALC, digits=3, 
               slab=study)
PPODATA

# Subgroups random effects models
IMPUTED_REM_PPODATA <- rma(PPODATA_MD, PPODATA_VAR, data=PPODATA_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='imputed'))
IMPUTED_REM_PPODATA

NOTIMPUTED_REM_PPODATA <- rma(PPODATA_MD, PPODATA_VAR, data=PPODATA_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='notimputed'))
NOTIMPUTED_REM_PPODATA

IMPUTEDFIG_REM_PPODATA <- rma(PPODATA_MD, PPODATA_VAR, data=PPODATA_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='figure'))
IMPUTEDFIG_REM_PPODATA

# Test for subgroup differences
PPOSEI <- sqrt(PPODATA_ESCALC$PPODATA_VAR)
SGDIFF <- metagen(PPODATA_MD, PPOSEI, studlab=PPODATA$study, data=PPODATA_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF
