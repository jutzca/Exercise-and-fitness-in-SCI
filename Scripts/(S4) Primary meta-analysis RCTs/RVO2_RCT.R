install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/RCT CTRLS/'
wdir <- '~/Documents/R meta-analysis/RCTS_DATA/'
setwd(data_dir)

RVO2RCT <- read.xlsx("RCTS_DATA.xlsx", sheet = 2)

RVO2RCT$n.con <- as.numeric(RVO2RCT$n.con)
RVO2RCT$mean.con <- as.numeric(RVO2RCT$mean.con)
RVO2RCT$sd.con <- as.numeric(RVO2RCT$sd.con)
RVO2RCT$n.int <- as.numeric(RVO2RCT$n.int)
RVO2RCT$mean.int <- as.numeric(RVO2RCT$mean.int)
RVO2RCT$sd.int <- as.numeric(RVO2RCT$sd.int)

# Effect size calculation
RVO2RCT_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                         m2i=mean.con,sd2i=sd.con,n2i=n.con,
                         data=RVO2RCT,var.names=c('RVO2RCT_MD','RVO2RCT_VAR'))

# Overall random effects model
REM_RVO2RCT <- rma(RVO2RCT_MD, RVO2RCT_VAR, data=RVO2RCT_ESCALC, digits=3, 
                   slab=study)
REM_RVO2RCT

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2RCT, ylim = c(-1,19), xlim = c(-50, 25), at = c(-10, -5, 0, 5, 10), cex=0.4, digits = c(1,0),
       ilab=format(cbind(RVO2RCT$mean.con, RVO2RCT$sd.con, RVO2RCT$n.con, RVO2RCT$mean.int, RVO2RCT$sd.int, RVO2RCT$n.int), digits = 1),
       ilab.xpos = c(-35, -31, -27, -23, -19, -15),
       showweights = TRUE, slab = study, col =10, xlab = 'Favours Control             Favours Intervention')

text(c(-35, -31, -27, -23, -19, -15), 17.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-31, -19), 18.5, c('Control', 'Intervention'), cex=0.3)
text(21.25, 17.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-47.5, 17.5, cex=0.4, bquote(paste(bold("Study"))))
text(16.25, 17.5, cex=0.4, bquote(paste(bold("Weight"))))


text(-33.5, -0.95, cex=0.4, bquote(paste("for All Studies: p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2RCT$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2RCT$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2RCT$I2, digits=1, format="f")), "%"))) 

# Egger's test
regtest(REM_RVO2RCT, model = "rma", ret.fit = FALSE, digits = 2)

#Funnel plot
par(mar=c(4,4,1,2))
funnel(REM_RVO2RCT, ylim = c(0,2), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')

