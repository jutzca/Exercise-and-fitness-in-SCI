install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/OBSV DATA/'
wdir <- '~/Documents/R meta-analysis/OBSV_DATA/'
setwd(data_dir)

RVO2OBSV <- read.xlsx("OBSV_DATA.xlsx", sheet = 2)

RVO2OBSV$n.post <- as.numeric(RVO2OBSV$n.post)
RVO2OBSV$mean.post <- as.numeric(RVO2OBSV$mean.post)
RVO2OBSV$sd.post <- as.numeric(RVO2OBSV$sd.post)
RVO2OBSV$n.pre <- as.numeric(RVO2OBSV$n.pre)
RVO2OBSV$mean.pre <- as.numeric(RVO2OBSV$mean.pre)
RVO2OBSV$sd.pre <- as.numeric(RVO2OBSV$sd.pre)

# Naming subgroups
RVO2OBSV$subgroup <- as.factor(RVO2OBSV$subgroup)

# Effect size calculation
RVO2OBSV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=RVO2OBSV,var.names=c('RVO2OBSV_MD','RVO2OBSV_VAR'))

# Overall random effects model
REM_RVO2OBSV <- rma(RVO2OBSV_MD, RVO2OBSV_VAR, data=RVO2OBSV_ESCALC, digits=3, 
                    slab=study, method = "REML")
REM_RVO2OBSV

# Subgroups random effects models
REHAB_REM_RVO2OBSV <- rma(RVO2OBSV_MD, RVO2OBSV_VAR, data=RVO2OBSV_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='rehab'))
REHAB_REM_RVO2OBSV

FREE_REM_RVO2OBSV <- rma(RVO2OBSV_MD, RVO2OBSV_VAR, data=RVO2OBSV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='free'))
FREE_REM_RVO2OBSV

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2OBSV_ESCALC$RVO2OBSV_VAR)
SGDIFF <- metagen(RVO2OBSV_MD, RVO2SEI, studlab=RVO2OBSV$study, data=RVO2OBSV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2OBSV, ylim = c(-2, 12), xlim = c(-40, 20),at = c(-6, -3, 0, 3, 6), rows = c(8:8, 4:2), cex=0.4, digits = c(1,0),
       ilab=format(cbind(RVO2OBSV$mean.pre, RVO2OBSV$sd.pre, RVO2OBSV$mean.post, RVO2OBSV$sd.post, RVO2OBSV$n.post), digits = 2),
       ilab.xpos = c(-28, -24, -20, -16, -12),
       showweights = TRUE, slab = study, xlab = 'Mean Difference (mL/kg/min)', col=10)
       
addpoly(REHAB_REM_RVO2OBSV, row=6.5, cex=0.4, col=11, mlab="")
text(-26.4, 6.5, cex=0.4, bquote(paste("RE Model for Inpatient Rehabilitation: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(REHAB_REM_RVO2OBSV$I2, digits = 1, format="f")), "%")))

addpoly(FREE_REM_RVO2OBSV, row=0.5, cex=0.4, col=11, mlab="")
text(-28.8, 0.7, cex=0.4, bquote(paste("RE Model for Free-Living: p = 0.64",
                                       "; ", I^2, " = ", .(formatC(FREE_REM_RVO2OBSV$I2, digits = 1, format="f")), "%")))

text(c(-28, -24, -20, -16, -12), 10.5, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-26, -18), 11.5, c('Pre', 'Post'), cex=0.3)
text(16, 10.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-37.1, 10.5, cex=0.4, bquote(paste(bold("Study"))))
text(10, 10.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-22.2, -0.95, cex=0.4, bquote(paste("for All Studies: p = 0.43", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2OBSV$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2OBSV$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2OBSV$I2, digits=1, format="f")), "%"))) 

text(-33.2, 9.2, cex=0.44, bquote(paste(bolditalic("Inpatient Rehabilitation"))))
text(-35.9, 5.2, cex=0.44, bquote(paste(bolditalic("Free-Living"))))

# Egger's test
regtest(REM_RVO2OBSV, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(RVO2OBSV$subgroup, c("rehab", "free"))]
funnelplotdata <- funnel(REM_RVO2OBSV, ylim = c(0,3), digits = c(1,2), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Rehab","Free-Living"),cex=0.8, pch=20, col=c("purple","blue"))

