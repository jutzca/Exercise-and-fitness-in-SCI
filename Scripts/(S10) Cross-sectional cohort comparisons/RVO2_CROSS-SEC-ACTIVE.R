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

RVO2CSACT <- read.xlsx("COHORT_DATA.xlsx", sheet = 3)

RVO2CSACT$n.con <- as.numeric(RVO2CSACT$n.inactive)
RVO2CSACT$mean.con <- as.numeric(RVO2CSACT$mean.inactive)
RVO2CSACT$sd.con <- as.numeric(RVO2CSACT$sd.inactive)
RVO2CSACT$n.int <- as.numeric(RVO2CSACT$n.active)
RVO2CSACT$mean.int <- as.numeric(RVO2CSACT$mean.active)
RVO2CSACT$sd.int <- as.numeric(RVO2CSACT$sd.active)

# Effect size calculation
RVO2CSACT_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                           m2i=mean.con,sd2i=sd.con,n2i=n.con,
                           data=RVO2CSACT,var.names=c('RVO2CSACT_MD','RVO2CSACT_VAR'))

# Overall fixed effects model
REM_RVO2CSACT <- rma(RVO2CSACT_MD, RVO2CSACT_VAR, data=RVO2CSACT_ESCALC, digits=3, 
                     slab=study, method = "REML")
REM_RVO2CSACT

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2CSACT, ylim = c(-2,7), xlim = c(-95, 65), at = c(-20, -10, 0, 10, 20), rows = c(4:1), cex=0.4, digits = c(1,0),
       ilab=format(cbind(RVO2CSACT$mean.inactive, RVO2CSACT$sd.inactive, RVO2CSACT$n.inactive, RVO2CSACT$mean.active, RVO2CSACT$sd.active, RVO2CSACT$n.active), digits = 2),
       ilab.xpos = c(-65, -58, -51, -44, -37, -30),
       showweights = TRUE, slab = study, xlab = 'Higher in inactive         Higher in active', col=10)

text(c(-65, -58, -51, -44, -37, -30), 5.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-58, -37), 6, c('Inactive', 'Active'), cex=0.35)
text(44.5, 5.5, cex=0.4, bquote(paste(bold("Weight"))))
text(57, 5.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-88.7, 5.5, cex=0.4, bquote(paste(bold("Study"))))

text(-69, -0.94, cex=0.4, bquote(paste(": p < 0.001", 
                                       "; ",Z, " = ",.(formatC(REM_RVO2CSACT$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_RVO2CSACT$I2, digits=1, format="f")), "%")))


