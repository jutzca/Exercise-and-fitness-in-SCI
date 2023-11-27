install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/PPO DATA/'
wdir <- '~/Documents/R meta-analysis/PPO_DATA/'
setwd(data_dir)

PPOTSI <- read.xlsx("PPO_DATA_mixedout.xlsx", sheet = 1)

PPOTSI$n.post <- as.integer(PPOTSI$n.post)
PPOTSI$mean.post <- as.numeric(PPOTSI$mean.post)
PPOTSI$sd.post <- as.numeric(PPOTSI$sd.post)
PPOTSI$n.pre <- as.integer(PPOTSI$n.pre)
PPOTSI$mean.pre <- as.numeric(PPOTSI$mean.pre)
PPOTSI$sd.pre <- as.numeric(PPOTSI$sd.pre)

# Naming subgroups
PPOTSI$subgroup <- as.factor(PPOTSI$subgroup)

# Effect size calculation
PPOTSI_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=PPOTSI,var.names=c('PPOTSI_MD','PPOTSI_VAR'))

# Overall random effects model
REM_PPOTSI <- rma(PPOTSI_MD, PPOTSI_VAR, data=PPOTSI_ESCALC, digits=3, 
                   slab=study)
REM_PPOTSI

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
                   data = PPOTSI,
                   sm = "MD",
                   fixed = FALSE,
                   random = TRUE,
                   method.tau = "REML",
                   hakn = TRUE)
summary(m.cont)

influence(model = REM_PPOTSI)

# Subgroups random effects models
ACUTE_REM_PPOTSI <- rma(PPOTSI_MD, PPOTSI_VAR, data=PPOTSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='acute (<1-year)'))
ACUTE_REM_PPOTSI

CHRONIC_REM_PPOTSI<- rma(PPOTSI_MD, PPOTSI_VAR, data=PPOTSI_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='chronic (>1-year)'))
CHRONIC_REM_PPOTSI

MIXED_REM_PPOTSI <- rma(PPOTSI_MD, PPOTSI_VAR, data=PPOTSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_PPOTSI

NR.CD_REM_PPOTSI <- rma(PPOTSI_MD, PPOTSI_VAR, data=PPOTSI_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_PPOTSI

# Test for subgroup differences
PPOSEI <- sqrt(PPOTSI_ESCALC$PPOTSI_VAR)
SGDIFF <- metagen(PPOTSI_MD, PPOSEI, studlab=PPOTSI$study, data=PPOTSI_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOTSI, ylim = c(-1, 60), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(55:46, 41:2), digits = c(0,0), cex=0.5,
       ilab=format(cbind(PPOTSI$mean.pre,
                         PPOTSI$sd.pre,
                         PPOTSI$mean.post,
                         PPOTSI$sd.post,
                         PPOTSI$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(ACUTE_REM_PPOTSI, row=44.5, cex=0.5, col=11, mlab="")
text(-106, 44.5, cex=0.5, bquote(paste("RE Model for Acute TSI (<1 year): p < 0.001",
                                         "; ", I^2, " = ", .(formatC(ACUTE_REM_PPOTSI$I2, digits = 1, format="f")), "%")))

addpoly(CHRONIC_REM_PPOTSI, row=0.5, cex=0.5, col=11, mlab="")
text(-104.1, 0.7, cex=0.5, bquote(paste("RE Model for Chronic TSI (>1 year): p < 0.001",
                                         "; ", I^2, " = ", .(formatC(CHRONIC_REM_PPOTSI$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 59, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 60, c('Pre', 'Post'), cex=0.38)
text(84, 59.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_PPOTSI$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_PPOTSI$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_PPOTSI$I2, digits=1, format="f")), "%)"))) 

text(-153.5, 56.5, cex=0.5, bquote(paste(bolditalic("Acute (<1 year)"))))
text(-152, 42.5, cex=0.5, bquote(paste(bolditalic("Chronic (>1 year)"))))

text(-104, -2, cex=0.48, "Test for Subgroup Differences: Q = 6.56, df = 1, p = 0.01")

# Funnel plot
par(mar=c(4,4,1,2))
regtest(REM_PPOTSI, model = "rma", ret.fit = FALSE, digits = 2)

my_colours <- c("purple","blue","orange","red")[match(PPOTSI$subgroup, c("acute (<1-year)", "chronic (>1-year)"))] 
funnelplotdata <- funnel(REM_PPOTSI, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic"),cex=0.8, pch=20, col=c("purple","blue"))
