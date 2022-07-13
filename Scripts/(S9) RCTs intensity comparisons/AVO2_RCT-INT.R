install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/RCT INT/'
wdir <- '~/Documents/R meta-analysis/RCTS_INTENSITY_DATA/'
setwd(data_dir)

AVO2RCTINT <- read.xlsx("RCTS_INTENSITY_DATA.xlsx", sheet = 1)

AVO2RCTINT$n.mod <- as.numeric(AVO2RCTINT$n.mod)
AVO2RCTINT$mean.mod <- as.numeric(AVO2RCTINT$mean.mod)
AVO2RCTINT$sd.mod <- as.numeric(AVO2RCTINT$sd.mod)
AVO2RCTINT$n.vig <- as.numeric(AVO2RCTINT$n.vig)
AVO2RCTINT$mean.vig <- as.numeric(AVO2RCTINT$mean.vig)
AVO2RCTINT$sd.vig <- as.numeric(AVO2RCTINT$sd.vig)

# Naming subgroups
AVO2RCTINT$subgroup <- as.factor(AVO2RCTINT$subgroup)

# Effect size calculation
AVO2RCTINT_ESCALC <- escalc(measure="MD",m1i=mean.mod,sd1i=sd.mod,n1i=n.mod,
                         m2i=mean.vig,sd2i=sd.vig,n2i=n.vig,
                         data=AVO2RCTINT,var.names=c('AVO2RCTINT_MD','AVO2RCTINT_VAR'))

# Overall random effects model
REM_AVO2RCTINT <- rma(AVO2RCTINT_MD, AVO2RCTINT_VAR, data=AVO2RCTINT_ESCALC, digits=3, 
                   slab=study)
REM_AVO2RCTINT

# Subgroups random effects models
MATCH_REM_AVO2RCTINT <- rma(AVO2RCTINT_MD, AVO2RCTINT_VAR, data=AVO2RCTINT_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='match'))
MATCH_REM_AVO2RCTINT

UNMATCH_REM_AVO2RCTINT <- rma(AVO2RCTINT_MD, AVO2RCTINT_VAR, data=AVO2RCTINT_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_AVO2RCTINT

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2RCTINT_ESCALC$AVO2RCTINT_VAR)
SGDIFF <- metagen(AVO2RCTINT_MD, AVO2SEI, studlab=AVO2RCTINT$study, data=AVO2RCTINT_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
forest(REM_AVO2RCTINT, ylim = c(-2, 10), xlim = c(-1.3, 0.8), alim = c(-0.4,0.4), rows = c(6:6, 2:2), cex=0.4,
       ilab=format(cbind(AVO2RCTINT$mean.mod, AVO2RCTINT$sd.mod, AVO2RCTINT$n.mod, AVO2RCTINT$mean.vig, AVO2RCTINT$sd.vig, AVO2RCTINT$n.vig), digits = 2),
       ilab.xpos = c(-1, -0.9, -0.8, -0.7, -0.6, -0.5),
       showweights = TRUE, slab = study, xlab = 'Favours Moderate                             Favours Vigorous', col=10)

addpoly(MATCH_REM_AVO2RCTINT, row=4.5, cex=0.4, col=11, mlab="")
text(-0.93, 4.5, cex=0.4, bquote(paste("RE Model for Matched Exercise Volume: p = 0.59",
                                       "; ", I^2, " = ", .(formatC(MATCH_REM_AVO2RCTINT$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_AVO2RCTINT, row=0.5, cex=0.4, col=11, mlab="")
text(-0.93, 0.7, cex=0.4, bquote(paste("RE Model for Different Exercise Volume: p = 0.16",
                                       "; ", I^2, " = ", .(formatC(UNMATCH_REM_AVO2RCTINT$I2, digits = 1, format="f")), "%")))

text(c(-1, -0.9, -0.8, -0.7, -0.6, -0.5), 8.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-0.9, -0.6), 9.5, c('Moderate', 'Vigorous'), cex=0.35)
text(0.69, 8.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-1.23, 8.5, cex=0.4, bquote(paste(bold("Study"))))
text(0.49, 8.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-0.843, -0.94, cex=0.4, bquote(paste("for All Studies: p = 0.67", 
                                         "; ",Tau^2, " = ",.(formatC(REM_AVO2RCTINT$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_AVO2RCTINT$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_AVO2RCTINT$I2, digits=1, format="f")), "%"))) 

text(-1.107, 7.2, cex=0.44, bquote(paste(bolditalic("Matched Exercise Volume"))))
text(-1.109, 3.2, cex=0.44, bquote(paste(bolditalic("Different Exercise Volume"))))

text(-0.95, -1.8, cex=0.4, "Test for Subgroup Differences: Q = 2.13, df = 1, p = 0.14")

# Egger's test
regtest(REM_AVO2RCTINT, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(AVO2RCTINT$subgroup, c("match", "unmatch"))]
funnelplotdata <- funnel(REM_AVO2RCTINT, ylim = c(0,0.2), digits = c(2,2), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Matched","Unmatched"),cex=0.8, pch=20, col=c("purple","blue"))

