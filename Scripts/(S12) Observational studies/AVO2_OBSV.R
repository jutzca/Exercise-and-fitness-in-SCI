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

AVO2OBSV <- read.xlsx("OBSV_DATA.xlsx", sheet = 1)

AVO2OBSV$n.post <- as.numeric(AVO2OBSV$n.post)
AVO2OBSV$mean.post <- as.numeric(AVO2OBSV$mean.post)
AVO2OBSV$sd.post <- as.numeric(AVO2OBSV$sd.post)
AVO2OBSV$n.pre <- as.numeric(AVO2OBSV$n.pre)
AVO2OBSV$mean.pre <- as.numeric(AVO2OBSV$mean.pre)
AVO2OBSV$sd.pre <- as.numeric(AVO2OBSV$sd.pre)

# Naming subgroups
AVO2OBSV$subgroup <- as.factor(AVO2OBSV$subgroup)

# Effect size calculation
AVO2OBSV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2OBSV,var.names=c('AVO2OBSV_MD','AVO2OBSV_VAR'))

# Overall random effects model
REM_AVO2OBSV <- rma(AVO2OBSV_MD, AVO2OBSV_VAR, data=AVO2OBSV_ESCALC, digits=3, 
                   slab=study, method = "REML")
REM_AVO2OBSV

# Subgroups random effects models
REHAB_REM_AVO2OBSV <- rma(AVO2OBSV_MD, AVO2OBSV_VAR, data=AVO2OBSV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='rehab'))
REHAB_REM_AVO2OBSV

FREE_REM_AVO2OBSV <- rma(AVO2OBSV_MD, AVO2OBSV_VAR, data=AVO2OBSV_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='free'))
FREE_REM_AVO2OBSV

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2OBSV_ESCALC$AVO2OBSV_VAR)
SGDIFF <- metagen(AVO2OBSV_MD, AVO2SEI, studlab=AVO2OBSV$study, data=AVO2OBSV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2OBSV, ylim = c(-2, 14), xlim = c(-7,4), alim = c(-1,1), rows = c(10:8, 4:2), digits = c(2,0), cex=0.4,
       ilab=format(cbind(AVO2OBSV$mean.pre, AVO2OBSV$sd.pre, AVO2OBSV$mean.post, AVO2OBSV$sd.post, AVO2OBSV$n.post), digits = 2),
       ilab.xpos = c(-5, -4.25, -3.5, -2.75, -2),
       showweights = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(REHAB_REM_AVO2OBSV, row=6.5, cex=0.4, col=11, mlab="")
text(-4.45, 6.5, cex=0.4, bquote(paste("RE Model for Inpatient Rehabilitation: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(REHAB_REM_AVO2OBSV$I2, digits = 1, format="f")), "%")))

addpoly(FREE_REM_AVO2OBSV, row=0.5, cex=0.4, col=11, mlab="")
text(-4.95, 0.7, cex=0.4, bquote(paste("RE Model for Free-Living: p = 0.06",
                                    "; ", I^2, " = ", .(formatC(FREE_REM_AVO2OBSV$I2, digits = 1, format="f")), "%")))

text(c(-5, -4.25, -3.5, -2.75, -2), 12.5, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-4.625, -3.125), 13.5, c('Pre', 'Post'), cex=0.3)
text(3.25, 12.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-7.5, 12.5, cex=0.4, bquote(paste(bold("Study"))))
text(1.9, 12.5, cex=0.4, bquote(paste(bold("Weight"))))


text(-4.19, -0.95, cex=0.4, bquote(paste("for All Studies: p < 0.001", 
                                         "; ",Z, " = ",.(formatC(REM_AVO2OBSV$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_AVO2OBSV$I2, digits=1, format="f")), "%"))) 

text(-5.75, 11.2, cex=0.44, bquote(paste(bolditalic("Inpatient Rehabilitation"))))
text(-6.25, 5.2, cex=0.44, bquote(paste(bolditalic("Free-Living"))))

# Egger's test
regtest(REM_AVO2OBSV, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(AVO2OBSV$subgroup, c("rehab", "free"))]
funnelplotdata <- funnel(REM_AVO2OBSV, ylim = c(0,0.2), digits = c(2,2), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Rehab","Free-Living"),cex=0.8, pch=20, col=c("purple","blue"))

