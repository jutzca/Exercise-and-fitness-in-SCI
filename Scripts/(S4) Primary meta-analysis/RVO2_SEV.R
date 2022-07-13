install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/RVO2 DATA/'
wdir <- '~/Documents/R meta-analysis/RVO2_DATA/'
setwd(data_dir)

RVO2SEV <- read.xlsx("RVO2_DATA.xlsx", sheet = 3)

RVO2SEV$n.post <- as.integer(RVO2SEV$n.post)
RVO2SEV$mean.post <- as.numeric(RVO2SEV$mean.post)
RVO2SEV$sd.post <- as.numeric(RVO2SEV$sd.post)
RVO2SEV$n.pre <- as.integer(RVO2SEV$n.pre)
RVO2SEV$mean.pre <- as.numeric(RVO2SEV$mean.pre)
RVO2SEV$sd.pre <- as.numeric(RVO2SEV$sd.pre)

# Naming subgroups
RVO2SEV$subgroup <- as.factor(RVO2SEV$subgroup)

# Effect size calculation
RVO2SEV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=RVO2SEV,var.names=c('RVO2SEV_MD','RVO2SEV_VAR'))

# Overall random effects model
REM_RVO2SEV <- rma(RVO2SEV_MD, RVO2SEV_VAR, data=RVO2SEV_ESCALC, digits=3, 
                     slab=study)
REM_RVO2SEV

# Subgroups random effects models
COMP_REM_RVO2SEV <- rma(RVO2SEV_MD, RVO2SEV_VAR, data=RVO2SEV_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='comp'))
COMP_REM_RVO2SEV

INCOMP_REM_RVO2SEV <- rma(RVO2SEV_MD, RVO2SEV_VAR, data=RVO2SEV_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='incomp'))
INCOMP_REM_RVO2SEV

MIXED_REM_RVO2SEV <- rma(RVO2SEV_MD, RVO2SEV_VAR, data=RVO2SEV_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='mixed'))
MIXED_REM_RVO2SEV

NR.CD_REM_RVO2SEV <- rma(RVO2SEV_MD, RVO2SEV_VAR, data=RVO2SEV_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_RVO2SEV

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2SEV_ESCALC$RVO2SEV_VAR)
SGDIFF <- metagen(RVO2SEV_MD, RVO2SEI, data=RVO2SEV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2SEV, ylim = c(-1, 92), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(87:59, 54:42, 37:17, 12:2), digits = 1,
       ilab=format(cbind(RVO2SEV$mean.pre,
                         RVO2SEV$sd.pre,
                         RVO2SEV$mean.post,
                         RVO2SEV$sd.post,
                         RVO2SEV$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(COMP_REM_RVO2SEV, row=57.5, cex=0.4, col = 11, mlab="")
text(-57.5, 57.5, cex=0.38, bquote(paste("RE Model for Motor-Complete: p < 0.002",
                                          "; ", I^2, " = ", .(formatC(COMP_REM_RVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(INCOMP_REM_RVO2SEV, row=40.5, cex=0.4, col = 11, mlab="")
text(-58, 40.5, cex=0.38, bquote(paste("RE Model for Motor-Incomplete: p = 0.02",
                                          "; ", I^2, " = ", .(formatC(INCOMP_REM_RVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_RVO2SEV, row=15.5, cex=0.4, col = 11, mlab="")
text(-63, 15.5, cex=0.38, bquote(paste("RE Model for Mixed: p < 0.002",
                                          "; ", I^2, " = ", .(formatC(MIXED_REM_RVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_RVO2SEV, row=0.5, cex=0.4,, col = 11, mlab="")
text(-62, 0.5, cex=0.38, bquote(paste("RE Model for NR/CD: p < 0.002",
                                          "; ", I^2, " = ", .(formatC(NR.CD_REM_RVO2SEV$I2, digits = 1, format="f")), "%")))


text(c(-40, -35, -30, -25, -20), 91, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 92, c('Pre', 'Post'), cex=0.3)
text(38.5, 91, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2SEV$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2SEV$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2SEV$I2, digits=1, format="f")), "%)"))) 

text(-68.7, 88.5, cex=0.4, bquote(paste(bolditalic("Motor-Complete (AIS A-B)"))))
text(-68, 55.5, cex=0.4, bquote(paste(bolditalic("Motor-Incomplete (AIS C-D)"))))
text(-74, 38.5, cex=0.4, bquote(paste(bolditalic("Mixed (AIS A-D)"))))
text(-66.7, 13.5, cex=0.4, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-55.5, -2, cex=0.38, "Test for Subgroup Differences: Q = 3.81, df = 3, p = 0.28")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(RVO2SEV$subgroup, c("comp", "incomp", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_RVO2SEV, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Motor-complete","Motor-incomplete","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))

