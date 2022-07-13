install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/Cross-sec cohorts/'
wdir <- '~/Documents/R meta-analysis/COHORT_DATA/'
setwd(data_dir)

RVO2CS <- read.xlsx("COHORT_DATA.xlsx", sheet = 2)

RVO2CS$n.con <- as.numeric(RVO2CS$n.inactive)
RVO2CS$mean.con <- as.numeric(RVO2CS$mean.inactive)
RVO2CS$sd.con <- as.numeric(RVO2CS$sd.inactive)
RVO2CS$n.int <- as.numeric(RVO2CS$n.active)
RVO2CS$mean.int <- as.numeric(RVO2CS$mean.active)
RVO2CS$sd.int <- as.numeric(RVO2CS$sd.active)

# Effect size calculation
RVO2CS_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                         m2i=mean.con,sd2i=sd.con,n2i=n.con,
                         data=RVO2CS,var.names=c('RVO2CS_MD','RVO2CS_VAR'))

# Overall random effects model
REM_RVO2CS <- rma(RVO2CS_MD, RVO2CS_VAR, data=RVO2CS_ESCALC, digits=3, 
                   slab=study)
REM_RVO2CS

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2CS, ylim = c(-2,12), xlim = c(-100, 65), at = c(-20, -10, 0, 10, 20), rows = c(9:1), cex=0.4,
       ilab=format(cbind(RVO2CS$mean.inactive, RVO2CS$sd.inactive, RVO2CS$n.inactive, RVO2CS$mean.active, RVO2CS$sd.active, RVO2CS$n.active), digits = 2),
       ilab.xpos = c(-65, -58, -51, -44, -37, -30),
       showweights = TRUE, slab = study, xlab = 'Higher in inactive         Higher in active', col=10)

text(c(-65, -58, -51, -44, -37, -30), 10.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-58, -37), 11, c('Inactive', 'Active'), cex=0.35)
text(33, 10.5, cex=0.4, bquote(paste(bold("Weight"))))
text(53, 10.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-92, 10.5, cex=0.4, bquote(paste(bold("Study"))))

text(-59, -0.94, cex=0.4, bquote(paste(": p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2CS$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2CS$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2CS$I2, digits=1, format="f")), "%"))) 
# Funnel plot
par(mar=c(4,4,1,2))
funnel(REM_RVO2CS, ylim = c(0,8), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
