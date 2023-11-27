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

RVO2TSI <- read.xlsx("RVO2_DATA_mixedout.xlsx", sheet = 1)

RVO2TSI$n.post <- as.integer(RVO2TSI$n.post, 0)
RVO2TSI$mean.post <- as.numeric(RVO2TSI$mean.post)
RVO2TSI$sd.post <- as.numeric(RVO2TSI$sd.post)
RVO2TSI$n.pre <- as.integer(RVO2TSI$n.pre)
RVO2TSI$mean.pre <- as.numeric(RVO2TSI$mean.pre)
RVO2TSI$sd.pre <- as.numeric(RVO2TSI$sd.pre)

# Naming subgroups
RVO2TSI$subgroup <- as.factor(RVO2TSI$subgroup)

# Effect size calculation
RVO2TSI_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=RVO2TSI,var.names=c('RVO2TSI_MD','RVO2TSI_VAR'))

# Overall random effects model
REM_RVO2TSI <- rma(RVO2TSI_MD, RVO2TSI_VAR, data=RVO2TSI_ESCALC, digits=3, 
                   slab=study)
REM_RVO2TSI

# Subgroups random effects models
ACUTE_REM_RVO2TSI <- rma(RVO2TSI_MD, RVO2TSI_VAR, data=RVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='acute (<1-year)'))
ACUTE_REM_RVO2TSI

CHRONIC_REM_RVO2TSI <- rma(RVO2TSI_MD, RVO2TSI_VAR, data=RVO2TSI_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='chronic (>1-year)'))
CHRONIC_REM_RVO2TSI

MIXED_REM_RVO2TSI <- rma(RVO2TSI_MD, RVO2TSI_VAR, data=RVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_RVO2TSI

NR.CD_REM_RVO2TSI <- rma(RVO2TSI_MD, RVO2TSI_VAR, data=RVO2TSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_RVO2TSI

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2TSI_ESCALC$RVO2TSI_VAR)
SGDIFF <- metagen(RVO2TSI_MD, RVO2SEI, data=RVO2TSI_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2TSI, ylim = c(-1, 70), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(65:58, 53:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2TSI$mean.pre,
                         RVO2TSI$sd.pre,
                         RVO2TSI$mean.post,
                         RVO2TSI$sd.post,
                         RVO2TSI$n.post), digits = 1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(ACUTE_REM_RVO2TSI, row=56.5, cex=0.4, col=11, mlab="")
text(-56.1, 56.5, cex=0.38, bquote(paste("RE Model for Acute TSI (<1 year): p = 0.002",
                                        "; ", I^2, " = ", .(formatC(ACUTE_REM_RVO2TSI$I2, digits = 1, format="f")), "%")))

addpoly(CHRONIC_REM_RVO2TSI, row=0.5, cex=0.4, col=11, mlab="")
text(-55, 0.7, cex=0.38, bquote(paste("RE Model for Chronic TSI (>1 year): p < 0.003",
                                         "; ", I^2, " = ", .(formatC(CHRONIC_REM_RVO2TSI$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 69, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 70, c('Pre', 'Post'), cex=0.3)
text(38.5, 69, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                             "; ",Tau^2, " = ",.(formatC(REM_RVO2TSI$tau2, digits = 2, format="f")),
                                             "; ",Z, " = ",.(formatC(REM_RVO2TSI$zval, digits = 2, format="f")),
                                             "; ",I^2, " = ",.(formatC(REM_RVO2TSI$I2, digits=1, format="f")), "%)"))) 

text(-74.4, 66.5, cex=0.4, bquote(paste(bolditalic("Acute (<1 year)"))))
text(-73.7, 54.5, cex=0.4, bquote(paste(bolditalic("Chronic (>1 year)"))))

text(-55.7, -2, cex=0.38, "Test for Subgroup Differences: Q = 0.57, df = 3, p = 0.45")

# Egger's test
regtest(REM_RVO2TSI, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(RVO2TSI$subgroup, c("acute (<1-year)", "chronic (>1-year)"))] 
funnelplotdata <- funnel(REM_RVO2TSI, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic"),cex=0.6, pch=20, col=c("purple","blue"))

# Leave one out analysis
influence(model = REM_RVO2TSI, progbar = TRUE)

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
                   data = RVO2TSI,
                   sm = "MD",
                   fixed = FALSE,
                   random = TRUE,
                   method.tau = "REML",
                   hakn = TRUE)
summary(m.cont)
