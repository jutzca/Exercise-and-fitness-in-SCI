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

PPORCT <- read.xlsx("RCTS_DATA.xlsx", sheet = 3)

PPORCT$n.con <- as.numeric(PPORCT$n.con)
PPORCT$mean.con <- as.numeric(PPORCT$mean.con)
PPORCT$sd.con <- as.numeric(PPORCT$sd.con)
PPORCT$n.int <- as.numeric(PPORCT$n.int)
PPORCT$mean.int <- as.numeric(PPORCT$mean.int)
PPORCT$sd.int <- as.numeric(PPORCT$sd.int)

# Effect size calculation
PPORCT_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                         m2i=mean.con,sd2i=sd.con,n2i=n.con,
                         data=PPORCT,var.names=c('PPORCT_MD','PPORCT_VAR'))

# Overall random effects model
REM_PPORCT <- rma(PPORCT_MD, PPORCT_VAR, data=PPORCT_ESCALC, digits=3, 
                   slab=study)
REM_PPORCT

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPORCT, ylim = c(-1,14), xlim = c(-90, 60), at = c(-30, -15, 0, 15, 30), cex=0.4, digits = c(0,0),
       ilab=format(cbind(PPORCT$mean.con, PPORCT$sd.con, PPORCT$n.con, PPORCT$mean.int, PPORCT$sd.int, PPORCT$n.int), digits = 0),
       ilab.xpos = c(-65, -59, -53, -47, -41, -35),  
       showweights = TRUE, slab = study, col =10, xlab = 'Favours Control                                   Favours Intervention')

text(c(-65, -59, -53, -47, -41, -35), 12.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-59, -41), 13.5, c('Control', 'Intervention'), cex=0.3)
text(53, 12.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-85, 12.5, cex=0.4, bquote(paste(bold("Study"))))
text(44, 12.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-56.5, -0.95, cex=0.4, bquote(paste("for All Studies: p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_PPORCT$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_PPORCT$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_PPORCT$I2, digits=1, format="f")), "%"))) 

# Egger's test
regtest(REM_PPORCT, model = "rma", ret.fit = FALSE, digits = 2)

#Funnel plot
par(mar=c(4,4,1,2))
funnel(REM_PPORCT, ylim = c(0,10), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (W)')
