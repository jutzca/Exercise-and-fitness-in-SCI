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

RVO2VOLUME <- read.xlsx("RVO2_DATA_mixedout.xlsx", sheet = 9)

RVO2VOLUME$n.post <- as.integer(RVO2VOLUME$n.post)
RVO2VOLUME$mean.post <- as.numeric(RVO2VOLUME$mean.post)
RVO2VOLUME$sd.post <- as.numeric(RVO2VOLUME$sd.post)
RVO2VOLUME$n.pre <- as.integer(RVO2VOLUME$n.pre)
RVO2VOLUME$mean.pre <- as.numeric(RVO2VOLUME$mean.pre)
RVO2VOLUME$sd.pre <- as.numeric(RVO2VOLUME$sd.pre)

# Naming subgroups
RVO2VOLUME$subgroup <- as.factor(RVO2VOLUME$subgroup)

# Effect size calculation
RVO2VOLUME_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=RVO2VOLUME,var.names=c('RVO2VOLUME_MD','RVO2VOLUME_VAR'))

# Overall random effects model
REM_RVO2VOLUME <- rma(RVO2VOLUME_MD, RVO2VOLUME_VAR, data=RVO2VOLUME_ESCALC, digits=3, 
                    slab=study)
REM_RVO2VOLUME

# Subgroups random effects models
FITNESS_REM_RVO2VOLUME <- rma(RVO2VOLUME_MD, RVO2VOLUME_VAR, data=RVO2VOLUME_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='fitness'), control=list(stepadj=0.5, maxiter=10000))
FITNESS_REM_RVO2VOLUME

CMETAB_REM_RVO2VOLUME <- rma(RVO2VOLUME_MD, RVO2VOLUME_VAR, data=RVO2VOLUME_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='cmetab'), control=list(stepadj=0.5, maxiter=10000))
CMETAB_REM_RVO2VOLUME

GENPOP_REM_RVO2VOLUME <- rma(RVO2VOLUME_MD, RVO2VOLUME_VAR, data=RVO2VOLUME_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='genpop'))
GENPOP_REM_RVO2VOLUME

NR.CD_REM_RVO2VOLUME <- rma(RVO2VOLUME_MD, RVO2VOLUME_VAR, data=RVO2VOLUME_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_RVO2VOLUME

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2VOLUME_ESCALC$RVO2VOLUME_VAR)
SGDIFF <- metagen(RVO2VOLUME_MD, RVO2SEI, studlab=RVO2VOLUME$study, data=RVO2VOLUME_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup, 
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2VOLUME, ylim = c(-1, 84), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(79:66, 61:30, 25:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2VOLUME$mean.pre,
                         RVO2VOLUME$sd.pre,
                         RVO2VOLUME$mean.post,
                         RVO2VOLUME$sd.post,
                         RVO2VOLUME$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(FITNESS_REM_RVO2VOLUME, row=64.5, cex=0.4, col=11, mlab="")
text(-55.8, 64.5, cex=0.38, bquote(paste("RE Model for 40-89 Minutes/Week: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(FITNESS_REM_RVO2VOLUME$I2, digits = 1, format="f")), "%")))

addpoly(CMETAB_REM_RVO2VOLUME, row=28.5, cex=0.4, col=11, mlab="")
text(-55, 28.5, cex=0.38, bquote(paste("RE Model for 90-149 Minutes/Week: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(CMETAB_REM_RVO2VOLUME$I2, digits = 1, format="f")), "%")))

addpoly(GENPOP_REM_RVO2VOLUME, row=0.5, cex=0.4, col=11, mlab="")
text(-55.5, 0.7, cex=0.38, bquote(paste("RE Model for " >="150 Minutes/Week: p < 0.002",
                                        "; ", I^2, " = ", .(formatC(GENPOP_REM_RVO2VOLUME$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 83, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 84, c('Pre', 'Post'), cex=0.3)
text(38.5, 83, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2VOLUME$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2VOLUME$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2VOLUME$I2, digits=1, format="f")), "%)"))) 

text(-51, 80.5, cex=0.4, bquote(paste(bolditalic("SCI-Specific Exercise Guidelines for Fitness (40-89 Min/Week)"))))
text(-42.5, 62.5, cex=0.4, bquote(paste(bolditalic("SCI-Specific Exercise Guidelines for Cardiometabolic Health (90-149 Min/Week)"))))
text(-48.3, 26.5, cex=0.4, bquote(paste(bolditalic("Achieving General Population Exercise Guidelines (">= "150 Min/Week)"))))

text(-55.7, -2, cex=0.38, "Test for Subgroup Differences: Q = 0.51, df = 2, p = 0.77")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(RVO2VOLUME$subgroup,c("fitness", "cmetab", "genpop"))]
funnelplotdata <- funnel(REM_RVO2VOLUME, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("40 - 89 min/wk","90 - 149 min/wk","150 min/wk or more"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))

