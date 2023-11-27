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

PPOOBSV <- read.xlsx("OBSV_DATA.xlsx", sheet = 3)

PPOOBSV$n.post <- as.numeric(PPOOBSV$n.post)
PPOOBSV$mean.post <- as.numeric(PPOOBSV$mean.post)
PPOOBSV$sd.post <- as.numeric(PPOOBSV$sd.post)
PPOOBSV$n.pre <- as.numeric(PPOOBSV$n.pre)
PPOOBSV$mean.pre <- as.numeric(PPOOBSV$mean.pre)
PPOOBSV$sd.pre <- as.numeric(PPOOBSV$sd.pre)

# Naming subgroups
PPOOBSV$subgroup <- as.factor(PPOOBSV$subgroup)

# Effect size calculation
PPOOBSV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=PPOOBSV,var.names=c('PPOOBSV_MD','PPOOBSV_VAR'))

# Overall random effects model
REM_PPOOBSV <- rma(PPOOBSV_MD, PPOOBSV_VAR, data=PPOOBSV_ESCALC, digits=3, 
                    slab=study, method = "REML")
REM_PPOOBSV

help("rma")

# Subgroups random effects models
REHAB_REM_PPOOBSV <- rma(PPOOBSV_MD, PPOOBSV_VAR, data=PPOOBSV_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='rehab'))
REHAB_REM_PPOOBSV

FREE_REM_PPOOBSV <- rma(PPOOBSV_MD, PPOOBSV_VAR, data=PPOOBSV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='free'))
FREE_REM_PPOOBSV

# Test for subgroup differences
PPOSEI <- sqrt(PPOOBSV_ESCALC$PPOOBSV_VAR)
SGDIFF <- metagen(PPOOBSV_MD, PPOSEI, studlab=PPOOBSV$study, data=PPOOBSV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOOBSV, ylim = c(-2, 14), xlim = c(-70, 40),at = c(-20, -10, 0, 10, 20), rows = c(10:8, 4:2), cex=0.4, digits = c(0,0),
       ilab=format(cbind(PPOOBSV$mean.pre, PPOOBSV$sd.pre, PPOOBSV$mean.post, PPOOBSV$sd.post, PPOOBSV$n.post), digits = 0),
       ilab.xpos = c(-49, -43, -37, -31, -25),
       showweights = TRUE, slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(REHAB_REM_PPOOBSV, row=6.5, cex=0.4, col=11, mlab="")
text(-45.3, 6.5, cex=0.4, bquote(paste("RE Model for Inpatient Rehabilitation: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(REHAB_REM_PPOOBSV$I2, digits = 1, format="f")), "%")))

addpoly(FREE_REM_PPOOBSV, row=0.5, cex=0.4, col=11, mlab="")
text(-49, 0.7, cex=0.4, bquote(paste("RE Model for Free-Living: p = 0.006",
                                       "; ", I^2, " = ", .(formatC(FREE_REM_PPOOBSV$I2, digits = 1, format="f")), "%")))

text(c(-49, -43, -37, -31, -25), 12.5, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-46, -34), 13.5, c('Pre', 'Post'), cex=0.3)
text(33, 12.5, cex=0.4, bquote(paste(bold("MD [95% CI]"))))
text(-65, 12.5, cex=0.4, bquote(paste(bold("Study"))))
text(24, 12.5, cex=0.4, bquote(paste(bold("Weight"))))

text(-41.8, -0.95, cex=0.4, bquote(paste("for All Studies: p < 0.001", 
                                         "; ",Z, " = ",.(formatC(REM_PPOOBSV$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_PPOOBSV$I2, digits=1, format="f")), "%"))) 

text(-57.7, 11.2, cex=0.44, bquote(paste(bolditalic("Inpatient Rehabilitation"))))
text(-62.4, 5.2, cex=0.44, bquote(paste(bolditalic("Free-Living"))))

# Egger's test
regtest(REM_PPOOBSV, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue")[match(PPOOBSV$subgroup, c("rehab", "free"))]
funnelplotdata <- funnel(REM_PPOOBSV, ylim = c(0,10), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Rehab","Free-Living"),cex=0.8, pch=20, col=c("purple","blue"))

