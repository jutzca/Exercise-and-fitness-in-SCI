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

PPOROB <- read.xlsx("SENS_DATA.xlsx", sheet = 5)

PPOROB$n.post <- as.numeric(PPOROB$n.post)
PPOROB$mean.post <- as.numeric(PPOROB$mean.post)
PPOROB$sd.post <- as.numeric(PPOROB$sd.post)
PPOROB$n.pre <- as.numeric(PPOROB$n.pre)
PPOROB$mean.pre <- as.numeric(PPOROB$mean.pre)
PPOROB$sd.pre <- as.numeric(PPOROB$sd.pre)

# Naming subgroups
PPOROB$subgroup <- as.factor(PPOROB$subgroup)

# Effect size calculation
PPOROB_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=PPOROB,var.names=c('PPOROB_MD','PPOROB_VAR'))

# Overall random effects model
REM_PPOROB <- rma(PPOROB_MD, PPOROB_VAR, data=PPOROB_ESCALC, digits=3, 
                   slab=study)
REM_PPOROB

# Subgroups random effects models
HIGH_REM_PPOROB <- rma(PPOROB_MD, PPOROB_VAR, data=PPOROB_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='high'))
HIGH_REM_PPOROB

LOW_REM_PPOROB <- rma(PPOROB_MD, PPOROB_VAR, data=PPOROB_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='low'))
LOW_REM_PPOROB

# Test for subgroup differences
PPOSEI <- sqrt(PPOROB_ESCALC$PPOROB_VAR)
SGDIFF <- metagen(PPOROB_MD, PPOSEI, studlab=PPOROB$study, data=PPOROB_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF