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

RVO2ROB <- read.xlsx("SENS_DATA.xlsx", sheet = 3)

RVO2ROB$n.post <- as.numeric(RVO2ROB$n.post)
RVO2ROB$mean.post <- as.numeric(RVO2ROB$mean.post)
RVO2ROB$sd.post <- as.numeric(RVO2ROB$sd.post)
RVO2ROB$n.pre <- as.numeric(RVO2ROB$n.pre)
RVO2ROB$mean.pre <- as.numeric(RVO2ROB$mean.pre)
RVO2ROB$sd.pre <- as.numeric(RVO2ROB$sd.pre)

# Naming subgroups
RVO2ROB$subgroup <- as.factor(RVO2ROB$subgroup)

# Effect size calculation
RVO2ROB_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=RVO2ROB,var.names=c('RVO2ROB_MD','RVO2ROB_VAR'))

# Overall random effects model
REM_RVO2ROB <- rma(RVO2ROB_MD, RVO2ROB_VAR, data=RVO2ROB_ESCALC, digits=3, 
                  slab=study)
REM_RVO2ROB

# Subgroups random effects models
HIGH_REM_RVO2ROB <- rma(RVO2ROB_MD, RVO2ROB_VAR, data=RVO2ROB_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='high'))
HIGH_REM_RVO2ROB

LOW_REM_RVO2ROB <- rma(RVO2ROB_MD, RVO2ROB_VAR, data=RVO2ROB_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='low'))
LOW_REM_RVO2ROB

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2ROB_ESCALC$RVO2ROB_VAR)
SGDIFF <- metagen(RVO2ROB_MD, RVO2SEI, studlab=RVO2ROB$study, data=RVO2ROB_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF
