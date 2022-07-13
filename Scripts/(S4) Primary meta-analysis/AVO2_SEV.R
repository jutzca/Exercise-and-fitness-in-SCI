install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/AVO2 DATA/'
wdir <- '~/Documents/R meta-analysis/AVO2_DATA/'
setwd(data_dir)

AVO2SEV <- read.xlsx("AVO2_DATA.xlsx", sheet = 3)

AVO2SEV$n.post <- as.numeric(AVO2SEV$n.post)
AVO2SEV$mean.post <- as.numeric(AVO2SEV$mean.post)
AVO2SEV$sd.post <- as.numeric(AVO2SEV$sd.post)
AVO2SEV$n.pre <- as.numeric(AVO2SEV$n.pre)
AVO2SEV$mean.pre <- as.numeric(AVO2SEV$mean.pre)
AVO2SEV$sd.pre <- as.numeric(AVO2SEV$sd.pre)

# Naming subgroups
AVO2SEV$subgroup <- as.factor(AVO2SEV$subgroup)

# Effect size calculation
AVO2SEV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2SEV,var.names=c('AVO2SEV_MD','AVO2SEV_VAR'))

# Overall random effects model
REM_AVO2SEV<- rma(AVO2SEV_MD, AVO2SEV_VAR, data=AVO2SEV_ESCALC, digits=3, 
                   slab=study)
REM_AVO2SEV

# Subgroups random effects models
COMP_REM_AVO2SEV <- rma(AVO2SEV_MD, AVO2SEV_VAR, data=AVO2SEV_ESCALC, digits=3, 
                        slab=study, subset=(subgroup=='comp'))
COMP_REM_AVO2SEV

INCOMP_REM_AVO2SEV <- rma(AVO2SEV_MD, AVO2SEV_VAR, data=AVO2SEV_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='incomp'))
INCOMP_REM_AVO2SEV

MIXED_REM_AVO2SEV <- rma(AVO2SEV_MD, AVO2SEV_VAR, data=AVO2SEV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_AVO2SEV

NR.CD_REM_AVO2SEV <- rma(AVO2SEV_MD, AVO2SEV_VAR, data=AVO2SEV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_AVO2SEV

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2SEV_ESCALC$AVO2SEV_VAR)
SGDIFF <- metagen(AVO2SEV_MD, AVO2SEI, studlab=AVO2SEV$study, data=AVO2SEV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2SEV, ylim = c(-1,87), xlim = c(-11.5,7), alim = c(-2,2), rows = c(82:56, 51:44, 39:18, 13:2),
       ilab=format(cbind(AVO2SEV$mean.pre, AVO2SEV$sd.pre, AVO2SEV$mean.post, AVO2SEV$sd.post, AVO2SEV$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(COMP_REM_AVO2SEV, row=54.5, cex=0.4, col=11, mlab="")
text(-8, 54.5, cex=0.4, bquote(paste("RE Model for Motor-Complete: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(COMP_REM_AVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(INCOMP_REM_AVO2SEV, row=42.5, cex=0.4, col=11, mlab="")
text(-8.05, 42.5, cex=0.4, bquote(paste("RE Model for Motor-Incomplete: p = 0.08",
                                       "; ", I^2, " = ", .(formatC(INCOMP_REM_AVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_AVO2SEV, row=16.5, cex=0.4, col=11, mlab="")
text(-8.65, 16.5, cex=0.4, bquote(paste("RE Model for Mixed: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_AVO2SEV$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_AVO2SEV, row=0.5, cex=0.4, col=11, mlab="")
text(-8.55, 0.7, cex=0.4, bquote(paste("RE Model for NR/CD: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(NR.CD_REM_AVO2SEV$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 86, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 87, c('Pre', 'Post'), cex=0.3)
text(3.9, 86, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2SEV$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2SEV$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2SEV$I2, digits=1, format="f")), "%)"))) 

text(-9.4, 83.5, cex=0.44, bquote(paste(bolditalic("Motor-Complete (AIS A-B)"))))
text(-9.35, 52.5, cex=0.44, bquote(paste(bolditalic("Motor-Incomplete (AIS C-D)"))))
text(-10.1, 40.5, cex=0.44, bquote(paste(bolditalic("Mixed (AIS A-D)"))))
text(-9.12, 14.5, cex=0.44, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 7.24, df = 3, p = 0.06")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(AVO2SEV$subgroup, c("comp", "incomp", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_AVO2SEV, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Motor-complete","Motor-incomplete","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
