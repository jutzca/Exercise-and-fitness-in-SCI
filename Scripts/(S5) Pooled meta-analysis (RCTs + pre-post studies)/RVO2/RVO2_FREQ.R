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

RVO2FREQ <- read.xlsx("RVO2_DATA_mixedout.xlsx", sheet = 8)

RVO2FREQ$n.post <- as.integer(RVO2FREQ$n.post)
RVO2FREQ$mean.post <- as.numeric(RVO2FREQ$mean.post)
RVO2FREQ$sd.post <- as.numeric(RVO2FREQ$sd.post)
RVO2FREQ$n.pre <- as.integer(RVO2FREQ$n.pre)
RVO2FREQ$mean.pre <- as.numeric(RVO2FREQ$mean.pre)
RVO2FREQ$sd.pre <- as.numeric(RVO2FREQ$sd.pre)

# Naming subgroups
RVO2FREQ$subgroup <- as.factor(RVO2FREQ$subgroup)

# Effect size calculation
RVO2FREQ_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=RVO2FREQ,var.names=c('RVO2FREQ_MD','RVO2FREQ_VAR'))

# Overall random effects model
REM_RVO2FREQ <- rma(RVO2FREQ_MD, RVO2FREQ_VAR, data=RVO2FREQ_ESCALC, digits=3, 
                     slab=study)
REM_RVO2FREQ

# Subgroups random effects models
THREE_REM_RVO2FREQ <- rma(RVO2FREQ_MD, RVO2FREQ_VAR, data=RVO2FREQ_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='three'))
THREE_REM_RVO2FREQ

THREEFIVE_REM_RVO2FREQ <- rma(RVO2FREQ_MD, RVO2FREQ_VAR, data=RVO2FREQ_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='threefive'), control=list(stepadj=0.5, maxiter=10000))
THREEFIVE_REM_RVO2FREQ

FIVE_REM_RVO2FREQ <- rma(RVO2FREQ_MD, RVO2FREQ_VAR, data=RVO2FREQ_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='five'))
FIVE_REM_RVO2FREQ

NR.CD_REM_RVO2FREQ <- rma(RVO2FREQ_MD, RVO2FREQ_VAR, data=RVO2FREQ_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_RVO2FREQ

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2FREQ_ESCALC$RVO2FREQ_VAR)
SGDIFF <- metagen(RVO2FREQ_MD, RVO2SEI, studlab=RVO2FREQ$study, data=RVO2FREQ_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2FREQ, ylim = c(-1, 89), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(84:73, 68:17, 12:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2FREQ$mean.pre,
                         RVO2FREQ$sd.pre,
                         RVO2FREQ$mean.post,
                         RVO2FREQ$sd.post,
                         RVO2FREQ$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(THREE_REM_RVO2FREQ, row=71.5, cex=0.4, col=11, mlab="")
text(-57, 71.5, cex=0.38, bquote(paste("RE Model for <3 Sessions/Week: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(THREE_REM_RVO2FREQ$I2, digits = 1, format="f")), "%")))

addpoly(THREEFIVE_REM_RVO2FREQ, row=15.5, cex=0.4, col=11, mlab="")
text(-53, 15.5, cex=0.38, bquote(paste("RE Model for ">="3 to < 5 Sessions/Week: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(THREEFIVE_REM_RVO2FREQ$I2, digits = 1, format="f")), "%")))

addpoly(FIVE_REM_RVO2FREQ, row=0.5, cex=0.4, col=11, mlab="")
text(-56, 0.7, cex=0.38, bquote(paste("RE Model for ">="5 Sessions/Week: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(FIVE_REM_RVO2FREQ$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 88, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 89, c('Pre', 'Post'), cex=0.3)
text(38.5, 88, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2FREQ$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2FREQ$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2FREQ$I2, digits=1, format="f")), "%)"))) 

text(-72.5, 85.5, cex=0.44, bquote(paste(bolditalic("< 3 Sessions/Week"))))
text(-69.9, 69.5, cex=0.44, bquote(paste(bolditalic(" ">="3 to < 5 Sessions/Week"))))
text(-73.2, 13.5, cex=0.44, bquote(paste(bolditalic(" ">="5 Sessions/Week"))))

text(-55.7, -2, cex=0.38, "Test for Subgroup Differences: Q = 1.64, df = 2, p = 0.44")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(RVO2FREQ$subgroup,c("three", "threefive", "five"))]
funnelplotdata <- funnel(REM_RVO2FREQ, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Less than 3 sessions/wk","3 to 4 sessions/wk","5 or more sessions/wk"),
       cex=0.48, pch=20, col=c("purple","blue","orange"))

