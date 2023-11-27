install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/AVO2 DATA/'
wdir <- '~/Documents/R meta-analysis/AVO2_FDATA/'
setwd(data_dir)

AVO2VOLUME <- read.xlsx("AVO2_DATA_mixedout.xlsx", sheet = 9)

AVO2VOLUME$n.post <- as.numeric(AVO2VOLUME$n.post)
AVO2VOLUME$mean.post <- as.numeric(AVO2VOLUME$mean.post)
AVO2VOLUME$sd.post <- as.numeric(AVO2VOLUME$sd.post)
AVO2VOLUME$n.pre <- as.numeric(AVO2VOLUME$n.pre)
AVO2VOLUME$mean.pre <- as.numeric(AVO2VOLUME$mean.pre)
AVO2VOLUME$sd.pre <- as.numeric(AVO2VOLUME$sd.pre)

# Naming subgroups
AVO2VOLUME$subgroup <- as.factor(AVO2VOLUME$subgroup)

# Effect size calculation
AVO2VOLUME_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=AVO2VOLUME,var.names=c('AVO2VOLUME_MD','AVO2VOLUME_VAR'))

# Overall random effects model
REM_AVO2VOLUME <- rma(AVO2VOLUME_MD, AVO2VOLUME_VAR, data=AVO2VOLUME_ESCALC, digits=3, 
                      slab=study)
REM_AVO2VOLUME

# Subgroups random effects models
FITNESS_REM_AVO2VOLUME <- rma(AVO2VOLUME_MD, AVO2VOLUME_VAR, data=AVO2VOLUME_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='fitness'))
FITNESS_REM_AVO2VOLUME

CMETAB_REM_AVO2VOLUME <- rma(AVO2VOLUME_MD, AVO2VOLUME_VAR, data=AVO2VOLUME_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='cmetab'))
CMETAB_REM_AVO2VOLUME

GENPOP_REM_AVO2VOLUME <- rma(AVO2VOLUME_MD, AVO2VOLUME_VAR, data=AVO2VOLUME_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='genpop'))
GENPOP_REM_AVO2VOLUME

NR.CD_REM_AVO2VOLUME <- rma(AVO2VOLUME_MD, AVO2VOLUME_VAR, data=AVO2VOLUME_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_AVO2VOLUME

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2VOLUME_ESCALC$AVO2VOLUME_VAR)
SGDIFF <- metagen(AVO2VOLUME_MD, AVO2SEI, studlab=AVO2VOLUME$study, data=AVO2VOLUME_ESCALC, control=list(stepadj = 0.5, maxiter=10000), fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2VOLUME, ylim = c(-1,74), xlim = c(-11.5,7), alim = c(-2,2), rows = c(69:56, 51:20, 15:2), cex = 0.4,
       ilab=format(cbind(AVO2VOLUME$mean.pre, AVO2VOLUME$sd.pre, AVO2VOLUME$mean.post, AVO2VOLUME$sd.post, AVO2VOLUME$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(FITNESS_REM_AVO2VOLUME, row=54.5, cex=0.4, col=11, mlab="")
text(-7.8, 54.5, cex=0.4, bquote(paste("RE Model for 40-89 Minutes/Week: p < 0.001",
                                        "; ", I^2, " = ", .(formatC(FITNESS_REM_AVO2VOLUME$I2, digits = 1, format="f")), "%")))

addpoly(CMETAB_REM_AVO2VOLUME, row=18.5, cex=0.4, col=11, mlab="")
text(-7.65, 18.5, cex=0.4, bquote(paste("RE Model for 90-149 Minutes/Week: p < 0.001",
                                        "; ", I^2, " = ", .(formatC(CMETAB_REM_AVO2VOLUME$I2, digits = 1, format="f")), "%")))

addpoly(GENPOP_REM_AVO2VOLUME, row=0.5, cex=0.4, col=11, mlab="")
text(-7.8, 0.7, cex=0.4, bquote(paste("RE Model for " >="150 Minutes/Week: p < 0.001",
                                        "; ", I^2, " = ", .(formatC(GENPOP_REM_AVO2VOLUME$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 73, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 74, c('Pre', 'Post'), cex=0.3)
text(3.9, 73, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2VOLUME$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2VOLUME$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2VOLUME$I2, digits=1, format="f")), "%)")))

text(-7.3, 70.5, cex=0.44, bquote(paste(bolditalic("SCI-Specific Exercise Guidelines for Fitness (40-89 Min/Wk)"))))
text(-6.2, 52.5, cex=0.44, bquote(paste(bolditalic("SCI-Specific Exercise Guidelines for Cardiometabolic Health (90-149 Min/Wk)"))))
text(-6.9, 16.5, cex=0.44, bquote(paste(bolditalic("Achieving General Population Exercise Guidelines (">= "150 Min/Wk)"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 1.28, df = 3, p = 0.53")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(AVO2VOLUME$subgroup,c("fitness", "cmetab", "genpop"))]
funnelplotdata <- funnel(REM_AVO2VOLUME, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("40 - 89 min/wk","90 - 149 min/wk","150 min/wk or more"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))


