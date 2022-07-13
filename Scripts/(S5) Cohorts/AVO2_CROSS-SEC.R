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

AVO2CS <- read.xlsx("COHORT_DATA.xlsx", sheet = 1)

AVO2CS$n.con <- as.numeric(AVO2CS$n.inactive)
AVO2CS$mean.con <- as.numeric(AVO2CS$mean.inactive)
AVO2CS$sd.con <- as.numeric(AVO2CS$sd.inactive)
AVO2CS$n.int <- as.numeric(AVO2CS$n.active)
AVO2CS$mean.int <- as.numeric(AVO2CS$mean.active)
AVO2CS$sd.int <- as.numeric(AVO2CS$sd.active)

# Effect size calculation
AVO2CS_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                        m2i=mean.con,sd2i=sd.con,n2i=n.con,
                        data=AVO2CS,var.names=c('AVO2CS_MD','AVO2CS_VAR'))

# Overall fixed effects model
REM_AVO2CS <- rma(AVO2CS_MD, AVO2CS_VAR, data=AVO2CS_ESCALC, digits=3, 
                  slab=study, method = "FE")
REM_AVO2CS

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2CS, ylim = c(-2,12), xlim = c(-12.5, 7), alim = c(-2,2), rows = c(9:1), cex=0.4,
       ilab=format(cbind(AVO2CS$mean.inactive, AVO2CS$sd.inactive, AVO2CS$n.inactive, AVO2CS$mean.active, AVO2CS$sd.active, AVO2CS$n.active), digits = 2),
       ilab.xpos = c(-8, -7, -6, -5, -4, -3),
       showweights = TRUE, header = FALSE, slab = study, xlab = 'Higher in inactive       Higher in active', col=10)

text(c(-8, -7, -6, -5, -4, -3), 10.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-7, -4), 11.5, c('Inactive', 'Active'), cex=0.38)
text(3.3, 10.5, cex=0.4, bquote(paste(bold("Weight"))))
text(5.6, 10.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-11.55, 10.5, cex=0.4, bquote(paste(bold("Study"))))

text(-8.32, -0.94, cex=0.4, bquote(paste(": p < 0.001", 
                                        "; ",Z, " = ",.(formatC(REM_AVO2CS$zval, digits = 2, format="f")),
                                        "; ",I^2, " = ",.(formatC(REM_AVO2CS$I2, digits=1, format="f")), "%"))) 

par(mar=c(4,4,1,2))
funnel(REM_AVO2CS, ylim = c(0,0.4), digits = c(2,1), cex = 0.5, xlab = 'Mean Difference (L/min)')



