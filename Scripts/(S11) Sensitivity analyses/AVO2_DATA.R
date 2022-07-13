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

AVO2DATA <- read.xlsx("SENS_DATA.xlsx", sheet = 2)

AVO2DATA$n.post <- as.numeric(AVO2DATA$n.post)
AVO2DATA$mean.post <- as.numeric(AVO2DATA$mean.post)
AVO2DATA$sd.post <- as.numeric(AVO2DATA$sd.post)
AVO2DATA$n.pre <- as.numeric(AVO2DATA$n.pre)
AVO2DATA$mean.pre <- as.numeric(AVO2DATA$mean.pre)
AVO2DATA$sd.pre <- as.numeric(AVO2DATA$sd.pre)

# Naming subgroups
AVO2DATA$subgroup <- as.factor(AVO2DATA$subgroup)

# Effect size calculation
AVO2DATA_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2DATA,var.names=c('AVO2DATA_MD','AVO2DATA_VAR'))

# Overall random effects model
AVO2DATA<- rma(AVO2DATA_MD, AVO2DATA_VAR, data=AVO2DATA_ESCALC, digits=3, 
                  slab=study)
AVO2DATA

# Subgroups random effects models
IMPUTED_REM_AVO2DATA <- rma(AVO2DATA_MD, AVO2DATA_VAR, data=AVO2DATA_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='imputed'))
IMPUTED_REM_AVO2DATA

NOTIMPUTED_REM_AVO2DATA <- rma(AVO2DATA_MD, AVO2DATA_VAR, data=AVO2DATA_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='notimputed'))
NOTIMPUTED_REM_AVO2DATA

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2DATA_ESCALC$AVO2DATA_VAR)
SGDIFF <- metagen(AVO2DATA_MD, AVO2SEI, studlab=AVO2DATA$study, data=AVO2DATA_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

