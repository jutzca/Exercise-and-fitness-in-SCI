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

AVO2TSI <- read.xlsx("AVO2_DATA_mixedout.xlsx", sheet = 1)

AVO2TSI$n.post <- as.numeric(AVO2TSI$n.post)
AVO2TSI$mean.post <- as.numeric(AVO2TSI$mean.post)
AVO2TSI$sd.post <- as.numeric(AVO2TSI$sd.post)
AVO2TSI$n.pre <- as.numeric(AVO2TSI$n.pre)
AVO2TSI$mean.pre <- as.numeric(AVO2TSI$mean.pre)
AVO2TSI$sd.pre <- as.numeric(AVO2TSI$sd.pre)

# Naming subgroups
AVO2TSI$subgroup <- as.factor(AVO2TSI$subgroup)

# Effect size calculation
AVO2TSI_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2TSI,var.names=c('AVO2TSI_MD','AVO2TSI_VAR'))

# Overall random effects model
REM_AVO2TSI <- rma(AVO2TSI_MD, AVO2TSI_VAR, data=AVO2TSI_ESCALC, digits=3, 
                   slab=study)
REM_AVO2TSI

devtools::install_github("MathiasHarrer/dmetar")
library(dmetar)
library(devtools)

m.cont <- metacont(n.e = n.post,
                   mean.e = mean.post,
                   sd.e = sd.post,
                   n.c = n.pre,
                   mean.c = mean.pre,
                   sd.c = sd.pre,
                   studlab = study,
                   data = AVO2TSI,
                   sm = "MD",
                   fixed = FALSE,
                   random = TRUE,
                   method.tau = "REML",
                   hakn = TRUE)
summary(m.cont)

# Subgroups random effects models
ACUTE_REM_AVO2TSI <- rma(AVO2TSI_MD, AVO2TSI_VAR, data=AVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='acute (< 1 year)'))
ACUTE_REM_AVO2TSI

CHRONIC_REM_AVO2TSI <- rma(AVO2TSI_MD, AVO2TSI_VAR, data=AVO2TSI_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='chronic (> 1 year)'))
CHRONIC_REM_AVO2TSI

MIXED_REM_AVO2TSI <- rma(AVO2TSI_MD, AVO2TSI_VAR, data=AVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_AVO2TSI

NR.CD_REM_AVO2TSI <- rma(AVO2TSI_MD, AVO2TSI_VAR, data=AVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_AVO2TSI

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2TSI_ESCALC$AVO2TSI_VAR)
SGDIFF <- metagen(AVO2TSI_MD, AVO2SEI, studlab=AVO2TSI$study, data=AVO2TSI_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2TSI, ylim = c(-1,67), xlim = c(-11.5,7), alim = c(-2,2), rows = c(62:55, 50:2), cex=0.4,
       ilab=format(cbind(AVO2TSI$mean.pre, AVO2TSI$sd.pre, AVO2TSI$mean.post, AVO2TSI$sd.post, AVO2TSI$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(ACUTE_REM_AVO2TSI, row=53.5, cex=0.4, col=11, mlab="")
text(-7.8, 53.5, cex=0.4, bquote(paste("RE Model for Acute TSI (<1 year): p < 0.001",
                                       "; ", I^2, " = ", .(formatC(ACUTE_REM_AVO2TSI$I2, digits = 1, format="f")), "%")))

addpoly(CHRONIC_REM_AVO2TSI, row=0.5, cex=0.4, col=11, mlab="")
text(-7.7, 0.7, cex=0.4, bquote(paste("RE Model for Chronic TSI (>1 year): p < 0.001",
                                       "; ", I^2, " = ", .(formatC(CHRONIC_REM_AVO2TSI$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 66, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 67, c('Pre', 'Post'), cex=0.3)
text(3.9, 66, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2TSI$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2TSI$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2TSI$I2, digits=1, format="f")), "%)"))) 

text(-10.2, 63.5, cex=0.44, bquote(paste(bolditalic("Acute (<1 year)"))))
text(-10.05, 51.5, cex=0.44, bquote(paste(bolditalic("Chronic (>1 year)"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 0.18, df = 1, p = 0.67")

# Egger's test
regtest(REM_AVO2TSI, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))
my_colours <- c("purple","blue","orange","red")[match(AVO2TSI$subgroup, c("acute (< 1 year)", "chronic (> 1 year)", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_AVO2TSI, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic"),cex=0.8, pch=20, col=c("purple","blue"))
