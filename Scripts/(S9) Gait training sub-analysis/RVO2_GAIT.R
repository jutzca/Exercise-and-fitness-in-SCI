install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/Secondary meta-analyses/'
wdir <- '~/Documents/R meta-analysis/GAIT_DATA/'
setwd(data_dir)

RVO2GAIT <- read.xlsx("GAIT_DATA.xlsx", sheet = 2)

RVO2GAIT$n.post <- as.numeric(RVO2GAIT$n.post)
RVO2GAIT$mean.post <- as.numeric(RVO2GAIT$mean.post)
RVO2GAIT$sd.post <- as.numeric(RVO2GAIT$sd.post)
RVO2GAIT$n.pre <- as.numeric(RVO2GAIT$n.pre)
RVO2GAIT$mean.pre <- as.numeric(RVO2GAIT$mean.pre)
RVO2GAIT$sd.pre <- as.numeric(RVO2GAIT$sd.pre)

# Naming subgroups
RVO2GAIT$subgroup <- as.factor(RVO2GAIT$subgroup)

# Effect size calculation
RVO2GAIT_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=RVO2GAIT,var.names=c('RVO2GAIT_MD','RVO2GAIT_VAR'))

# Overall random effects model
REM_RVO2GAIT <- rma(RVO2GAIT_MD, RVO2GAIT_VAR, data=RVO2GAIT_ESCALC, digits=3, 
                    slab=study)
REM_RVO2GAIT

# Subgroups random effects models
ACE_REM_RVO2GAIT <- rma(RVO2GAIT_MD, RVO2GAIT_VAR, data=RVO2GAIT_ESCALC, digits=3, 
                        slab=study, subset=(subgroup=='ace'))
ACE_REM_RVO2GAIT

TREADMILL_REM_RVO2GAIT <- rma(RVO2GAIT_MD, RVO2GAIT_VAR, data=RVO2GAIT_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='treadmill'))
TREADMILL_REM_RVO2GAIT

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2GAIT_ESCALC$RVO2GAIT_VAR)
SGDIFF <- metagen(RVO2GAIT_MD, RVO2SEI, studlab=RVO2GAIT$study, data=RVO2GAIT_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2GAIT, ylim = c(-2,18), xlim = c(-85,60), at = c(-10, 0, 10, 20), rows = c(14:11, 7:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2GAIT$mean.pre, RVO2GAIT$sd.pre, RVO2GAIT$mean.post, RVO2GAIT$sd.post, RVO2GAIT$n.post), digits = 2),
       ilab.xpos = c(-44, -38, -32, -26, -20),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (mL/kg/min)', col=10)

addpoly(ACE_REM_RVO2GAIT, row=9.5, cex=0.4, col=11, mlab="")
text(-58.2, 9.5, cex=0.4, bquote(paste("RE Model for ACE CPET: p = 0.83",
                                       "; ", I^2, " = ", .(formatC(ACE_REM_RVO2GAIT$I2, digits = 1, format="f")), "%")))

addpoly(TREADMILL_REM_RVO2GAIT, row=0.5, cex=0.4, col=11, mlab="")
text(-56, 0.7, cex=0.4, bquote(paste("RE Model for Treadmill CPET: p = 0.12",
                                       "; ", I^2, " = ", .(formatC(TREADMILL_REM_RVO2GAIT$I2, digits = 1, format="f")), "%")))

text(c(-44, -38, -32, -26, -20), 17, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-41, -29), 18, c('Pre', 'Post'), cex=0.35)
text(36, 17, cex=0.4, bquote(paste(bold("Weight"))))

text(-41.9, -0.9, cex=0.4, bquote(paste("for All Studies (p = 0.40", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2GAIT$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2GAIT$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2GAIT$I2, digits=1, format="f")), "%)"))) 

text(-75.5, 15.25, cex=0.44, bquote(paste(bolditalic("ACE CPET"))))
text(-73.5, 8.25, cex=0.44, bquote(paste(bolditalic("Treadmill CPET"))))

text(-52, -1.8, cex=0.4, "Test for Subgroup Differences: Q = 0.81, df = 1, p = 0.37")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(AVO2TSI$subgroup, c("acute (< 1 year)", "chronic (> 1 year)", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_AVO2TSI, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
