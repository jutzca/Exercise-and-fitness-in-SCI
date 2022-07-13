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

AVO2FREQ <- read.xlsx("AVO2_DATA.xlsx", sheet = 8)

AVO2FREQ$n.post <- as.numeric(AVO2FREQ$n.post)
AVO2FREQ$mean.post <- as.numeric(AVO2FREQ$mean.post)
AVO2FREQ$sd.post <- as.numeric(AVO2FREQ$sd.post)
AVO2FREQ$n.pre <- as.numeric(AVO2FREQ$n.pre)
AVO2FREQ$mean.pre <- as.numeric(AVO2FREQ$mean.pre)
AVO2FREQ$sd.pre <- as.numeric(AVO2FREQ$sd.pre)

# Naming subgroups
AVO2FREQ$subgroup <- as.factor(AVO2FREQ$subgroup)

# Effect size calculation
AVO2FREQ_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=AVO2FREQ,var.names=c('AVO2FREQ_MD','AVO2FREQ_VAR'))

# Overall random effects model
REM_AVO2FREQ <- rma(AVO2FREQ_MD, AVO2FREQ_VAR, data=AVO2FREQ_ESCALC, digits=3, 
                    slab=study)
REM_AVO2FREQ

# Subgroups random effects models
THREE_REM_AVO2FREQ <- rma(AVO2FREQ_MD, AVO2FREQ_VAR, data=AVO2FREQ_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='three'))
THREE_REM_AVO2FREQ

THREEFIVE_REM_AVO2FREQ <- rma(AVO2FREQ_MD, AVO2FREQ_VAR, data=AVO2FREQ_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='threefive'))
THREEFIVE_REM_AVO2FREQ

FIVE_REM_AVO2FREQ <- rma(AVO2FREQ_MD, AVO2FREQ_VAR, data=AVO2FREQ_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='five'))
FIVE_REM_AVO2FREQ

NR.CD_REM_AVO2FREQ <- rma(AVO2FREQ_MD, AVO2FREQ_VAR, data=AVO2FREQ_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_AVO2FREQ

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2FREQ_ESCALC$AVO2FREQ_VAR)
SGDIFF <- metagen(AVO2FREQ_MD, AVO2SEI, studlab=AVO2FREQ$study, data=AVO2FREQ_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2FREQ, ylim = c(-1,87), xlim = c(-11.5,7), alim = c(-2,2), rows = c(82:64, 59:25, 20:10, 5:2), cex = 0.4,
       ilab=format(cbind(AVO2FREQ$mean.pre, AVO2FREQ$sd.pre, AVO2FREQ$mean.post, AVO2FREQ$sd.post, AVO2FREQ$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(THREE_REM_AVO2FREQ, row=62.5, cex=0.4, col=11, mlab="")
text(-7.85, 62.5, cex=0.4, bquote(paste("RE Model for < 3 Sessions/Week: p < 0.002",
                                        "; ", I^2, " = ", .(formatC(THREE_REM_AVO2FREQ$I2, digits = 1, format="f")), "%")))

addpoly(THREEFIVE_REM_AVO2FREQ, row=23.5, cex=0.4, col=11, mlab="")
text(-7.45, 23.5, cex=0.4, bquote(paste("RE Model for ">="3 to < 5 Sessions/Week: p < 0.002",
                                        "; ", I^2, " = ", .(formatC(THREEFIVE_REM_AVO2FREQ$I2, digits = 1, format="f")), "%")))

addpoly(FIVE_REM_AVO2FREQ, row=8.5, cex=0.4, col=11, mlab="")
text(-7.9, 8.5, cex=0.4, bquote(paste("RE Model for ">="5 Sessions/Week: p < 0.002",
                                        "; ", I^2, " = ", .(formatC(FIVE_REM_AVO2FREQ$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_AVO2FREQ, row=0.5, cex=0.4, col=11, mlab="")
text(-8.7, 0.5, cex=0.4, bquote(paste("RE Model for NR/CD: p = 0.15",
                                       "; ", I^2, " = ", .(formatC(NR.CD_REM_AVO2FREQ$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 86, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 87, c('Pre', 'Post'), cex=0.3)
text(3.9, 86, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2FREQ$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2FREQ$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2FREQ$I2, digits=1, format="f")), "%)")))

text(-9.9, 83.5, cex=0.44, bquote(paste(bolditalic("< 3 Sessions/Week"))))
text(-9.55, 60.5, cex=0.44, bquote(paste(bolditalic(" ">="3 to < 5 Sessions/Week"))))
text(-10, 21.5, cex=0.44, bquote(paste(bolditalic(" ">="5 Sessions/Week"))))
text(-9.13, 6.5, cex=0.44, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 5.40, df = 3, p = 0.14")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(AVO2FREQ$subgroup,c("three", "threefive", "five", "not reported/cannot determine"))]
funnelplotdata <- funnel(REM_AVO2FREQ, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Less than 3 sessions/wk","3 to 4 sessions/wk","5 or more sessions/wk","NR/CD"),
       cex=0.5, pch=20, col=c("purple","blue","orange","red"))



