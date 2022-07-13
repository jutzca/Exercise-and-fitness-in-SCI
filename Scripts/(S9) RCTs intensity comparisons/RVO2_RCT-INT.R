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

RVO2RCTINT <- read.xlsx("RCTS_INTENSITY_DATA.xlsx", sheet = 2)

RVO2RCTINT$n.mod <- as.numeric(RVO2RCTINT$n.mod)
RVO2RCTINT$mean.mod <- as.numeric(RVO2RCTINT$mean.mod)
RVO2RCTINT$sd.mod <- as.numeric(RVO2RCTINT$sd.mod)
RVO2RCTINT$n.vig <- as.numeric(RVO2RCTINT$n.vig)
RVO2RCTINT$mean.vig <- as.numeric(RVO2RCTINT$mean.vig)
RVO2RCTINT$sd.vig <- as.numeric(RVO2RCTINT$sd.vig)

# Naming subgroups
RVO2RCTINT$subgroup <- as.factor(RVO2RCTINT$subgroup)

# Effect size calculation
RVO2RCTINT_ESCALC <- escalc(measure="MD",m1i=mean.mod,sd1i=sd.mod,n1i=n.mod,
                            m2i=mean.vig,sd2i=sd.vig,n2i=n.vig,
                            data=RVO2RCTINT,var.names=c('RVO2RCTINT_MD','RVO2RCTINT_VAR'))

# Overall random effects model
REM_RVO2RCTINT<- rma(RVO2RCTINT_MD, RVO2RCTINT_VAR, data=RVO2RCTINT_ESCALC, digits=3, 
                      slab=study, method = "FE")
REM_RVO2RCTINT

# Subgroups random effects models
MATCH_REM_RVO2RCTINT <- rma(RVO2RCTINT_MD, RVO2RCTINT_VAR, data=RVO2RCTINT_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='match'))
MATCH_REM_RVO2RCTINT

UNMATCH_REM_RVO2RCTINT <- rma(RVO2RCTINT_MD, RVO2RCTINT_VAR, data=RVO2RCTINT_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_RVO2RCTINT

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2RCTINT_ESCALC$RVO2RCTINT_VAR)
SGDIFF <- metagen(RVO2RCTINT_MD, RVO2SEI, studlab=RVO2RCTINT$study, data=RVO2RCTINT_ESCALC, fixed=TRUE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2RCTINT, ylim = c(-2, 14), xlim = c(-23, 14), at = c(-5, -2.5, 0, 2.5, 5), rows = c(10:7, 3:2), cex=0.4, digits = 1,
       ilab=format(cbind(RVO2RCTINT$mean.mod, RVO2RCTINT$sd.mod, RVO2RCTINT$n.mod, RVO2RCTINT$mean.vig, RVO2RCTINT$sd.vig, RVO2RCTINT$n.vig), digits = 1),
       ilab.xpos = c(-17, -15, -13, -11, -9, -7),
       showweights = TRUE, slab = study, col=10, xlab = 'Favours Moderate                Favours Vigorous')

addpoly(MATCH_REM_RVO2RCTINT, row=5.5, cex=0.4, col=11, mlab="")
text(-16.5, 5.5, cex=0.4, bquote(paste("RE Model for Matched Exercise Volume: p = 0.48",
                                        "; ", I^2, " = ", .(formatC(MATCH_REM_RVO2RCTINT$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_RVO2RCTINT, row=0.5, cex=0.4, col=11, mlab="")
text(-16.5, 0.7, cex=0.4, bquote(paste("RE Model for Different Exercise Volume: p = 0.15",
                                       "; ", I^2, " = ", .(formatC(UNMATCH_REM_RVO2RCTINT$I2, digits = 1, format="f")), "%")))

text(c(-17, -15, -13, -11, -9, -7), 12.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-15, -9), 13.5, c('Moderate', 'Vigorous'), cex=0.35)
text(12.05, 12.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-21.75, 12.5, cex=0.4, bquote(paste(bold("Study"))))
text(9.4, 12.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-15.99, -0.94, cex=0.4, bquote(paste("for All Studies: p = 0.88", 
                                          "; ",Z, " = ",.(formatC(REM_RVO2RCTINT$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_RVO2RCTINT$I2, digits=1, format="f")), "%"))) 

text(-19.7, 11.2, cex=0.44, bquote(paste(bolditalic("Matched Exercise Volume"))))
text(-19.7, 4.2, cex=0.44, bquote(paste(bolditalic("Different Exercise Volume"))))

text(-16.9, -1.8, cex=0.4, "Test for Subgroup Differences: Q = 2.58, df = 1, p = 0.11")

# Egger's test
regtest(REM_RVO2RCTINT, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(RVO2RCTINT$subgroup, c("match", "unmatch"))]
funnelplotdata <- funnel(REM_RVO2RCTINT, ylim = c(0,2.4), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Matched","Unmatched"),cex=0.8, pch=20, col=c("purple","blue"))


