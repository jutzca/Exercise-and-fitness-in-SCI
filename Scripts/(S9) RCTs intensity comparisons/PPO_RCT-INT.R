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

PPORCTINT <- read.xlsx("RCTS_INTENSITY_DATA.xlsx", sheet = 3)

PPORCTINT$n.mod <- as.numeric(PPORCTINT$n.mod)
PPORCTINT$mean.mod <- as.numeric(PPORCTINT$mean.mod)
PPORCTINT$sd.mod <- as.numeric(PPORCTINT$sd.mod)
PPORCTINT$n.vig <- as.numeric(PPORCTINT$n.vig)
PPORCTINT$mean.vig <- as.numeric(PPORCTINT$mean.vig)
PPORCTINT$sd.vig <- as.numeric(PPORCTINT$sd.vig)

# Naming subgroups
PPORCTINT$subgroup <- as.factor(PPORCTINT$subgroup)

# Effect size calculation
PPORCTINT_ESCALC <- escalc(measure="MD",m1i=mean.mod,sd1i=sd.mod,n1i=n.mod,
                            m2i=mean.vig,sd2i=sd.vig,n2i=n.vig,
                            data=PPORCTINT,var.names=c('PPORCTINT_MD','PPORCTINT_VAR'))

# Overall random effects model
REM_PPORCTINT<- rma(PPORCTINT_MD, PPORCTINT_VAR, data=PPORCTINT_ESCALC, digits=3, 
                     slab=study, method = "FE")
REM_PPORCTINT

# Subgroups random effects models
MATCH_REM_PPORCTINT <- rma(PPORCTINT_MD, PPORCTINT_VAR, data=PPORCTINT_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='match'))
MATCH_REM_PPORCTINT

UNMATCH_REM_PPORCTINT <- rma(PPORCTINT_MD, PPORCTINT_VAR, data=PPORCTINT_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_PPORCTINT

# Test for subgroup differences
PPOSEI <- sqrt(PPORCTINT_ESCALC$PPORCTINT_VAR)
SGDIFF <- metagen(PPORCTINT_MD, PPOSEI, studlab=PPORCTINT$study, data=PPORCTINT_ESCALC, fixed=TRUE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPORCTINT, ylim = c(-2, 12), xlim = c(-65, 37), at = c(-20, -10, 0, 10, 20), rows = c(8:6, 2:2), cex=0.4, digits = 0,
       ilab=format(cbind(PPORCTINT$mean.mod, PPORCTINT$sd.mod, PPORCTINT$n.mod, PPORCTINT$mean.vig, PPORCTINT$sd.vig, PPORCTINT$n.vig), digits = 0),
       ilab.xpos = c(-50, -45, -40, -35, -30, -25),
       showweights = TRUE, col=10, slab=study, xlab = 'Favours Moderate                                  Favours Vigorous')

addpoly(MATCH_REM_PPORCTINT, row=4.5, cex=0.4, col=11, mlab="")
text(-47.2, 4.5, cex=0.4, bquote(paste("RE Model for Matched Exercise Volume: p = 0.60",
                                       "; ", I^2, " = ", .(formatC(MATCH_REM_PPORCTINT$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_PPORCTINT, row=0.5, cex=0.4, col=11, mlab="")
text(-47.2, 0.7, cex=0.4, bquote(paste("RE Model for Different Exercise Volume: p = 0.19",
                                       "; ", I^2, " = ", .(formatC(UNMATCH_REM_PPORCTINT$I2, digits = 1, format="f")), "%")))

text(c(-50, -45, -40, -35, -30, -25), 10.5, c('Mean', 'SD', 'Total N', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-45, -30), 11.5, c('Moderate', 'Vigorous'), cex=0.38)
text(32, 10.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-61.8, 10.5, cex=0.4, bquote(paste(bold("Study"))))
text(26, 10.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-45.7, -0.94, cex=0.4, bquote(paste("for All Studies: p = 0.62", 
                                          "; ",Z, " = ",.(formatC(REM_PPORCTINT$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPORCTINT$I2, digits=1, format="f")), "%"))) 

text(-55.9, 9.2, cex=0.44, bquote(paste(bolditalic("Matched Exercise Volume"))))
text(-55.9, 3.2, cex=0.44, bquote(paste(bolditalic("Different Exercise Volume"))))

text(-48, -1.8, cex=0.4, "Test for Subgroup Differences: Q = 1.72, df = 1, p = 0.19")

# Egger's test
regtest(REM_PPORCTINT, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(PPORCTINT$subgroup, c("match", "unmatch"))]
funnelplotdata <- funnel(REM_PPORCTINT, ylim = c(0,16), digits = c(1,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Matched","Unmatched"),cex=0.8, pch=20, col=c("purple","blue"))


