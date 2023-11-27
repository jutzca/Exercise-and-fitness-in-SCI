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

AVO2LENGTH <- read.xlsx("AVO2_DATA_mixedout.xlsx", sheet = 7)

AVO2LENGTH$n.post <- as.numeric(AVO2LENGTH$n.post)
AVO2LENGTH$mean.post <- as.numeric(AVO2LENGTH$mean.post)
AVO2LENGTH$sd.post <- as.numeric(AVO2LENGTH$sd.post)
AVO2LENGTH$n.pre <- as.numeric(AVO2LENGTH$n.pre)
AVO2LENGTH$mean.pre <- as.numeric(AVO2LENGTH$mean.pre)
AVO2LENGTH$sd.pre <- as.numeric(AVO2LENGTH$sd.pre)

# Naming subgroups
AVO2LENGTH$subgroup <- as.factor(AVO2LENGTH$subgroup)

# Effect size calculation
AVO2LENGTH_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=AVO2LENGTH,var.names=c('AVO2LENGTH_MD','AVO2LENGTH_VAR'))

# Overall random effects model
REM_AVO2LENGTH <- rma(AVO2LENGTH_MD, AVO2LENGTH_VAR, data=AVO2LENGTH_ESCALC, digits=3, 
                      slab=study)
REM_AVO2LENGTH

# Subgroups random effects models
SIX_REM_AVO2LENGTH <- rma(AVO2LENGTH_MD, AVO2LENGTH_VAR, data=AVO2LENGTH_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='6weeks'))
SIX_REM_AVO2LENGTH

SIXTWELVE_REM_AVO2LENGTH <- rma(AVO2LENGTH_MD, AVO2LENGTH_VAR, data=AVO2LENGTH_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='6-12weeks'))
SIXTWELVE_REM_AVO2LENGTH

TWELVE_REM_AVO2LENGTH <- rma(AVO2LENGTH_MD, AVO2LENGTH_VAR, data=AVO2LENGTH_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='12weeks'))
TWELVE_REM_AVO2LENGTH

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2LENGTH_ESCALC$AVO2LENGTH_VAR)
SGDIFF <- metagen(AVO2LENGTH_MD, AVO2SEI, studlab=AVO2LENGTH$study, data=AVO2LENGTH_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2LENGTH, ylim = c(-1,88), xlim = c(-11.5,7), alim = c(-2,2), rows = c(83:73, 68:34, 29:2), cex = 0.4,
       ilab=format(cbind(AVO2LENGTH$mean.pre, AVO2LENGTH$sd.pre, AVO2LENGTH$mean.post, AVO2LENGTH$sd.post, AVO2LENGTH$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(SIX_REM_AVO2LENGTH, row=71.5, cex=0.4, col=11, mlab="")
text(-8.35, 71.5, cex=0.4, bquote(paste("RE Model for "<="6 weeks: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(SIX_REM_AVO2LENGTH$I2, digits = 1, format="f")), "%")))

addpoly(SIXTWELVE_REM_AVO2LENGTH, row=32.5, cex=0.4, col=11, mlab="")
text(-7.92, 32.5, cex=0.4, bquote(paste("RE Model for > 6 to "<="12 weeks: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(SIXTWELVE_REM_AVO2LENGTH$I2, digits = 1, format="f")), "%")))

addpoly(TWELVE_REM_AVO2LENGTH, row=0.5, cex=0.4, col=11, mlab="")
text(-8.3, 0.7, cex=0.4, bquote(paste("RE Model for > 12 weeks: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(TWELVE_REM_AVO2LENGTH$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 87, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 88, c('Pre', 'Post'), cex=0.3)
text(3.9, 87, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2LENGTH$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2LENGTH$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2LENGTH$I2, digits=1, format="f")), "%)")))

text(-10.55, 84.5, cex=0.44, bquote(paste(bolditalic(" "<= "6 weeks"))))
text(-10, 69.5, cex=0.44, bquote(paste(bolditalic("> 6 to "<="12 weeks"))))
text(-10.4, 30.5, cex=0.44, bquote(paste(bolditalic("> 12 weeks"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 1.24, df = 2, p = 0.54")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(AVO2LENGTH$subgroup,c("6weeks", "6-12weeks", "12weeks"))]
funnelplotdata <- funnel(REM_AVO2LENGTH, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("6 weeks or less","7 to 12 weeks ","More than 12 weeks"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))

