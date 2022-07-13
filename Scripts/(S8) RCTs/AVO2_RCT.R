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

AVO2RCT <- read.xlsx("RCTS_DATA.xlsx", sheet = 1)

AVO2RCT$n.con <- as.numeric(AVO2RCT$n.con)
AVO2RCT$mean.con <- as.numeric(AVO2RCT$mean.con)
AVO2RCT$sd.con <- as.numeric(AVO2RCT$sd.con)
AVO2RCT$n.int <- as.numeric(AVO2RCT$n.int)
AVO2RCT$mean.int <- as.numeric(AVO2RCT$mean.int)
AVO2RCT$sd.int <- as.numeric(AVO2RCT$sd.int)

# Effect size calculation
AVO2RCT_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                         m2i=mean.con,sd2i=sd.con,n2i=n.con,
                         data=AVO2RCT,var.names=c('AVO2RCT_MD','AVO2RCT_VAR'))

# Overall random effects model
REM_AVO2RCT <- rma(AVO2RCT_MD, AVO2RCT_VAR, data=AVO2RCT_ESCALC, digits=3, 
                   slab=study)
REM_AVO2RCT

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2RCT, ylim = c(-1,14), xlim = c(-5.5,3), alim = c(-1,1), cex=0.4,
       ilab=format(cbind(AVO2RCT$mean.con, AVO2RCT$sd.con, AVO2RCT$n.con, AVO2RCT$mean.int, AVO2RCT$sd.int, AVO2RCT$n.int), digits = 2),
       ilab.xpos = c(-4, -3.5, -3, -2.5, -2, -1.5),
       showweights = TRUE, slab = study, xlab = 'Favours Control         Favours Intervention', col=10)

text(c(-4, -3.5, -3, -2.5, -2, -1.5), 12.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-3.5, -2), 13.5, c('Control', 'Intervention'), cex=0.3)
text(2.55, 12.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-5.22, 12.5, cex=0.4, bquote(paste(bold("Study"))))
text(1.8, 12.5, cex=0.4, bquote(paste(bold("Weight"))))


text(-3.62, -0.94, cex=0.4, bquote(paste("for All Studies: p = 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_AVO2RCT$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_AVO2RCT$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_AVO2RCT$I2, digits=1, format="f")), "%"))) 

# Egger's test
regtest(REM_AVO2RCT, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))
funnel(REM_AVO2RCT, ylim = c(0,0.9), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (L/min)')



