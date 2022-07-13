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

RVO2LENGTH <- read.xlsx("RVO2_DATA.xlsx", sheet = 7)

RVO2LENGTH$n.post <- as.numeric(RVO2LENGTH$n.post)
RVO2LENGTH$mean.post <- as.numeric(RVO2LENGTH$mean.post)
RVO2LENGTH$sd.post <- as.numeric(RVO2LENGTH$sd.post)
RVO2LENGTH$n.pre <- as.numeric(RVO2LENGTH$n.pre)
RVO2LENGTH$mean.pre <- as.numeric(RVO2LENGTH$mean.pre)
RVO2LENGTH$sd.pre <- as.numeric(RVO2LENGTH$sd.pre)

# Naming subgroups
RVO2LENGTH$subgroup <- as.factor(RVO2LENGTH$subgroup)

# Effect size calculation
RVO2LENGTH_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=RVO2LENGTH,var.names=c('RVO2LENGTH_MD','RVO2LENGTH_VAR'))

# Overall random effects model
REM_RVO2LENGTH <- rma(RVO2LENGTH_MD, RVO2LENGTH_VAR, data=RVO2LENGTH_ESCALC, digits=3, 
                     slab=study)
REM_RVO2LENGTH

# Subgroups random effects models
SIX_REM_RVO2LENGTH <- rma(RVO2LENGTH_MD, RVO2LENGTH_VAR, data=RVO2LENGTH_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='6weeks'))
SIX_REM_RVO2LENGTH

SIXTWELVE_REM_RVO2LENGTH <- rma(RVO2LENGTH_MD, RVO2LENGTH_VAR, data=RVO2LENGTH_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='6-12weeks'))
SIXTWELVE_REM_RVO2LENGTH

TWELVE_REM_RVO2LENGTH <- rma(RVO2LENGTH_MD, RVO2LENGTH_VAR, data=RVO2LENGTH_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='12weeks'))
TWELVE_REM_RVO2LENGTH

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2LENGTH_ESCALC$RVO2LENGTH_VAR)
SGDIFF <-  metagen(RVO2LENGTH_MD, RVO2SEI, studlab=RVO2LENGTH$study, data=RVO2LENGTH_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2LENGTH, ylim = c(-1, 88), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(83:61, 56:21, 16:2), digits = 1,
       ilab=format(cbind(RVO2LENGTH$mean.pre,
                         RVO2LENGTH$sd.pre,
                         RVO2LENGTH$mean.post,
                         RVO2LENGTH$sd.post,
                         RVO2LENGTH$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(SIX_REM_RVO2LENGTH, row=59.5, cex=0.4, col=11, mlab="")
text(-61, 59.5, cex=0.38, bquote(paste("RE Model for "<="6 weeks",
                                         ": p < 0.001; ", I^2, " = ", .(formatC(SIX_REM_RVO2LENGTH$I2, digits = 1, format="f")), "%")))

addpoly(SIXTWELVE_REM_RVO2LENGTH, row=19.5, cex=0.4, col=11, mlab="")
text(-56.6, 19.5, cex=0.38, bquote(paste("RE Model for > 6 to "<="12 weeks",
                                         ": p < 0.001; ", I^2, " = ", .(formatC(SIXTWELVE_REM_RVO2LENGTH$I2, digits = 1, format="f")), "%")))

addpoly(TWELVE_REM_RVO2LENGTH, row=0.5, cex=0.4, col=11, mlab="")
text(-60.7, 0.5, cex=0.38, bquote(paste("RE Model for > 12 weeks: p < 0.001",
                                         "; ", I^2, " = ", .(formatC(TWELVE_REM_RVO2LENGTH$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 87, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 88, c('Pre', 'Post'), cex=0.3)
text(38.5, 87, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2LENGTH$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2LENGTH$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2LENGTH$I2, digits=1, format="f")), "%)"))) 

text(-77.5, 84.5, cex=0.4, (bquote(paste(bolditalic(" "<= "6 weeks")))))
text(-73, 57.5, cex=0.4, (bquote(paste(bolditalic("> 6 to "<="12 weeks")))))
text(-76, 17.5, cex=0.4, bquote(paste(bolditalic("> 12 weeks"))))

text(-55.7, -2, cex=0.38, "Test for Subgroup Differences: Q = 6.02, df = 2, p = 0.05")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(RVO2LENGTH$subgroup,c("6weeks", "6-12weeks", "12weeks"))]
funnelplotdata <- funnel(REM_RVO2LENGTH, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("6 weeks or less","7 to 12 weeks","More than 12 weeks"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))

