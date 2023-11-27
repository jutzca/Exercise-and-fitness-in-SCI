install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("devtools")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/SENS/'
wdir <- '~/Documents/R meta-analysis/SENS_DATA/'
setwd(data_dir)

RVO2DATA <- read.xlsx("SENS_DATA.xlsx", sheet = 4)

RVO2DATA$n.post <- as.numeric(RVO2DATA$n.post)
RVO2DATA$mean.post <- as.numeric(RVO2DATA$mean.post)
RVO2DATA$sd.post <- as.numeric(RVO2DATA$sd.post)
RVO2DATA$n.pre <- as.numeric(RVO2DATA$n.pre)
RVO2DATA$mean.pre <- as.numeric(RVO2DATA$mean.pre)
RVO2DATA$sd.pre <- as.numeric(RVO2DATA$sd.pre)

# Naming subgroups
RVO2DATA$subgroup <- as.factor(RVO2DATA$subgroup)

# Effect size calculation
RVO2DATA_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=RVO2DATA,var.names=c('RVO2DATA_MD','RVO2DATA_VAR'))

# Overall random effects model
RVO2DATA<- rma(RVO2DATA_MD, RVO2DATA_VAR, data=RVO2DATA_ESCALC, digits=3, 
               slab=study)
RVO2DATA

# Subgroups random effects models
IMPUTED_REM_RVO2DATA <- rma(RVO2DATA_MD, RVO2DATA_VAR, data=RVO2DATA_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='imputed'))
IMPUTED_REM_RVO2DATA

NOTIMPUTED_REM_RVO2DATA <- rma(RVO2DATA_MD, RVO2DATA_VAR, data=RVO2DATA_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='notimputed'))
NOTIMPUTED_REM_RVO2DATA

IMPUTEDFIG_REM_RVO2DATA <- rma(RVO2DATA_MD, RVO2DATA_VAR, data=RVO2DATA_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='figure'))
IMPUTEDFIG_REM_RVO2DATA

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2DATA_ESCALC$RVO2DATA_VAR)
SGDIFF <- metagen(RVO2DATA_MD, RVO2SEI, studlab=RVO2DATA$study, data=RVO2DATA_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF
