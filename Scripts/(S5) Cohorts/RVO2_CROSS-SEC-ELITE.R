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

RVO2CSEL <- read.xlsx("COHORT_DATA.xlsx", sheet = 4)

RVO2CSEL$n.con <- as.numeric(RVO2CSEL$n.inactive)
RVO2CSEL$mean.con <- as.numeric(RVO2CSEL$mean.inactive)
RVO2CSEL$sd.con <- as.numeric(RVO2CSEL$sd.inactive)
RVO2CSEL$n.int <- as.numeric(RVO2CSEL$n.active)
RVO2CSEL$mean.int <- as.numeric(RVO2CSEL$mean.active)
RVO2CSEL$sd.int <- as.numeric(RVO2CSEL$sd.active)

# Effect size calculation
RVO2CSEL_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                           m2i=mean.con,sd2i=sd.con,n2i=n.con,
                           data=RVO2CSEL,var.names=c('RVO2CSEL_MD','RVO2CSEL_VAR'))

# Overall fixed effects model
REM_RVO2CSEL <- rma(RVO2CSEL_MD, RVO2CSEL_VAR, data=RVO2CSEL_ESCALC, digits=3, 
                     slab=study, method = "FE")
REM_RVO2CSEL

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2CSEL, ylim = c(-2,9), xlim = c(-95, 65), at = c(-20, -10, 0, 10, 20), rows = c(6:1), cex=0.4, digits = c(1,0),
       ilab=format(cbind(RVO2CSEL$mean.inactive, RVO2CSEL$sd.inactive, RVO2CSEL$n.inactive, RVO2CSEL$mean.active, RVO2CSEL$sd.active, RVO2CSEL$n.active), digits = 2),
       ilab.xpos = c(-65, -58, -51, -44, -37, -30),
       showweights = TRUE, slab = study, xlab = 'Higher in inactive         Higher in elite', col=10)

text(c(-65, -58, -51, -44, -37, -30), 7.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-58, -37), 8, c('Inactive', 'Elite'), cex=0.35)
text(41.4, 7.5, cex=0.4, bquote(paste(bold("Weight"))))
text(55, 7.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-88.4, 7.5, cex=0.4, bquote(paste(bold("Study"))))

text(-65.2, -0.94, cex=0.4, bquote(paste(": p < 0.001", 
                                       "; ",Z, " = ",.(formatC(REM_RVO2CSEL$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_RVO2CSEL$I2, digits=1, format="f")), "%")))

