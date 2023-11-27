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

RVO2INTENS <- read.xlsx("RVO2_DATA_mixedout.xlsx", sheet = 5)

RVO2INTENS$n.post <- as.integer(RVO2INTENS$n.post)
RVO2INTENS$mean.post <- as.numeric(RVO2INTENS$mean.post)
RVO2INTENS$sd.post <- as.numeric(RVO2INTENS$sd.post)
RVO2INTENS$n.pre <- as.integer(RVO2INTENS$n.pre)
RVO2INTENS$mean.pre <- as.numeric(RVO2INTENS$mean.pre)
RVO2INTENS$sd.pre <- as.numeric(RVO2INTENS$sd.pre)

# Naming subgroups
RVO2INTENS$subgroup <- as.factor(RVO2INTENS$subgroup)

# Effect size calculation
RVO2INTENS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=RVO2INTENS,var.names=c('RVO2INTENS_MD','RVO2INTENS_VAR'))

# Overall random effects model
REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                     slab=study)
REM_RVO2INTENS

# Subgroups random effects models
MOD_REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='mod'))
MOD_REM_RVO2INTENS

MODVIG_REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='modvig'), control = list(stepadj = 0.5, maxiter=10000))
MODVIG_REM_RVO2INTENS

VIG_REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='vig'), control = list(stepadj = 0.5, maxiter=10000))
VIG_REM_RVO2INTENS

SUPRA_REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='supra'))
SUPRA_REM_RVO2INTENS

MIXED_REM_RVO2INTENS <- rma(RVO2INTENS_MD, RVO2INTENS_VAR, data=RVO2INTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='mixed/cd'), control = list(stepadj = 0.5, maxiter=10000))
MIXED_REM_RVO2INTENS

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2INTENS_ESCALC$RVO2INTENS_VAR)
SGDIFF <- metagen(RVO2INTENS_MD, RVO2SEI, studlab=RVO2INTENS$study, data=RVO2INTENS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup, 
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2INTENS, ylim = c(-1, 79), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(74:61, 56:31, 26:7, 2:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2INTENS$mean.pre,
                         RVO2INTENS$sd.pre,
                         RVO2INTENS$mean.post,
                         RVO2INTENS$sd.post,
                         RVO2INTENS$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(MOD_REM_RVO2INTENS, row=59.5, cex=0.4, col=11, mlab="")
text(-60.8, 59.5, cex=0.38, bquote(paste("RE Model for Moderate: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(MOD_REM_RVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(MODVIG_REM_RVO2INTENS, row=29.5, cex=0.4, col=11, mlab="")
text(-55.3, 29.5, cex=0.38, bquote(paste("RE Model for Moderate-to-Vigorous: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(MODVIG_REM_RVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(VIG_REM_RVO2INTENS, row=5.5, cex=0.4, col=11, mlab="")
text(-61.7, 5.5, cex=0.38, bquote(paste("RE Model for Vigorous: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(VIG_REM_RVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(SUPRA_REM_RVO2INTENS, row=0.5, cex=0.4, col=11, mlab="")
text(-59.6, 0.7, cex=0.38, bquote(paste("RE Model for Supramaximal: p = 0.82",
                                         "; ", I^2, " = ", .(formatC(SUPRA_REM_RVO2INTENS$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 78, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 79, c('Pre', 'Post'), cex=0.3)
text(38.5, 78, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2INTENS$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2INTENS$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2INTENS$I2, digits=1, format="f")), "%)"))) 

text(-77.5, 75.5, cex=0.4, bquote(paste(bolditalic("Moderate"))))
text(-70.8, 57.5, cex=0.4, bquote(paste(bolditalic("Moderate-to-Vigorous"))))
text(-77.9, 27.5, cex=0.4, bquote(paste(bolditalic("Vigorous"))))
text(-75.1, 3.5, cex=0.4, bquote(paste(bolditalic("Supramaximal"))))

text(-55.5, -2, cex=0.4, "Test for Subgroup Differences: Q = 1.95, df = 3, p = 0.58")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(RVO2INTENS$subgroup,c("mod", "modvig", "vig", "supra"))]
funnelplotdata <- funnel(REM_RVO2INTENS, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Moderate","Mod-to-Vig","Vigorous","Supramaximal"),cex=0.5, pch=20, col=c("purple","blue","orange","red"))

