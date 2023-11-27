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

AVO2ROB <- read.xlsx("SENS_DATA.xlsx", sheet = 1)

AVO2ROB$n.post <- as.numeric(AVO2ROB$n.post)
AVO2ROB$mean.post <- as.numeric(AVO2ROB$mean.post)
AVO2ROB$sd.post <- as.numeric(AVO2ROB$sd.post)
AVO2ROB$n.pre <- as.numeric(AVO2ROB$n.pre)
AVO2ROB$mean.pre <- as.numeric(AVO2ROB$mean.pre)
AVO2ROB$sd.pre <- as.numeric(AVO2ROB$sd.pre)

# Naming subgroups
AVO2ROB$subgroup <- as.factor(AVO2ROB$subgroup)

# Effect size calculation
AVO2ROB_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2ROB,var.names=c('AVO2ROB_MD','AVO2ROB_VAR'))

# Overall random effects model
REM_AVO2ROB<- rma(AVO2ROB_MD, AVO2ROB_VAR, data=AVO2ROB_ESCALC, digits=3, 
                  slab=study)
REM_AVO2ROB

# Subgroups random effects models
HIGH_REM_AVO2ROB <- rma(AVO2ROB_MD, AVO2ROB_VAR, data=AVO2ROB_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='high'))
HIGH_REM_AVO2ROB

LOW_REM_AVO2ROB <- rma(AVO2ROB_MD, AVO2ROB_VAR, data=AVO2ROB_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='low'))
LOW_REM_AVO2ROB

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2ROB_ESCALC$AVO2ROB_VAR)
SGDIFF <- metagen(AVO2ROB_MD, AVO2SEI, studlab=AVO2ROB$study, data=AVO2ROB_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF




