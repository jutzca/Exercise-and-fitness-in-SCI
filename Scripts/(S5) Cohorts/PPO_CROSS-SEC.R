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

PPOCS <- read.xlsx("COHORT_DATA.xlsx", sheet = 5)

PPOCS$n.con <- as.numeric(PPOCS$n.inactive)
PPOCS$mean.con <- as.numeric(PPOCS$mean.inactive)
PPOCS$sd.con <- as.numeric(PPOCS$sd.inactive)
PPOCS$n.int <- as.numeric(PPOCS$n.active)
PPOCS$mean.int <- as.numeric(PPOCS$mean.active)
PPOCS$sd.int <- as.numeric(PPOCS$sd.active)

# Effect size calculation
PPOCS_ESCALC <- escalc(measure="MD",m1i=mean.int,sd1i=sd.int,n1i=n.int,
                        m2i=mean.con,sd2i=sd.con,n2i=n.con,
                        data=PPOCS,var.names=c('PPOCS_MD','PPOCS_VAR'))

# Overall fixed effects model
REM_PPOCS <- rma(PPOCS_MD, PPOCS_VAR, data=PPOCS_ESCALC, digits=3, 
                  slab=study, method = "FE")
REM_PPOCS

forest(REM_PPOCLASS, ylim = c(-1, 79), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(74:72, 67:46, 41:8, 3:2), digits = c(0,0),
       ilab=format(cbind(PPOCLASS$mean.pre,
                         PPOCLASS$sd.pre,
                         PPOCLASS$mean.post,
                         PPOCLASS$sd.post,
                         PPOCLASS$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)



# Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOCS, ylim = c(-2,7), xlim = c(-145, 110), at = c(-60, -40, -20, 0, 20, 40, 60), rows = c(4:1), cex=0.4, digits = c(0,0),
       ilab=format(cbind(PPOCS$mean.inactive, PPOCS$sd.inactive, PPOCS$n.inactive, PPOCS$mean.active, PPOCS$sd.active, PPOCS$n.active), digits = 0),
       ilab.xpos = c(-110, -102, -94, -86, -78, -70), showweights = TRUE, slab = study, xlab = 'Higher in inactive                                       Higher in active', col=10)

text(c(-110, -102, -94, -86, -78, -70), 5.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-102, -78), 6, c('Inactive', 'Active'), cex=0.35)
text(79, 5.5, cex=0.4, bquote(paste(bold("Weight"))))
text(96, 5.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-135, 5.5, cex=0.4, bquote(paste(bold("Study"))))

text(-103, -0.94, cex=0.4, bquote(paste(": p < 0.001", 
                                       "; ",Z, " = ",.(formatC(REM_PPOCS$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_PPOCS$I2, digits=1, format="f")), "%"))) 

# Funnel plot
par(mar=c(4,4,1,2))
funnel(REM_PPOCS, ylim = c(0,30), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')

