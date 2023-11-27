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

AVO2EXMOD <- read.xlsx("AVO2_DATA_mixedout.xlsx", sheet = 4)

AVO2EXMOD$n.post <- as.numeric(AVO2EXMOD$n.post)
AVO2EXMOD$mean.post <- as.numeric(AVO2EXMOD$mean.post)
AVO2EXMOD$sd.post <- as.numeric(AVO2EXMOD$sd.post)
AVO2EXMOD$n.pre <- as.numeric(AVO2EXMOD$n.pre)
AVO2EXMOD$mean.pre <- as.numeric(AVO2EXMOD$mean.pre)
AVO2EXMOD$sd.pre <- as.numeric(AVO2EXMOD$sd.pre)

# Naming subgroups
AVO2EXMOD$subgroup <- as.factor(AVO2EXMOD$subgroup)

# Effect size calculation
AVO2EXMOD_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=AVO2EXMOD,var.names=c('AVO2EXMOD_MD','AVO2EXMOD_VAR'))

# Overall random effects model
REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                     slab=study)
REM_AVO2EXMOD

# Subgroups random effects models
AEROBIC_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='aerobic'))
AEROBIC_REM_AVO2EXMOD

RESISTANCE_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='resistance'))
RESISTANCE_REM_AVO2EXMOD

FES_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='fes'))
FES_REM_AVO2EXMOD

GAIT_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='gait'))
GAIT_REM_AVO2EXMOD

MIXED_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='mixed'))
MIXED_REM_AVO2EXMOD

BEVC_REM_AVO2EXMOD <- rma(AVO2EXMOD_MD, AVO2EXMOD_VAR, data=AVO2EXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='bchange'))
BEVC_REM_AVO2EXMOD

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2EXMOD_ESCALC$AVO2EXMOD_VAR)
SGDIFF <- metagen(AVO2EXMOD_MD, AVO2SEI, studlab=AVO2EXMOD$study, data=AVO2EXMOD_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup,
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2EXMOD, ylim = c(-1,100), xlim = c(-11.5,7), alim = c(-2,2), rows = c(95:70, 65:62, 57:41, 36:27, 22:9, 4:2), cex=0.4,
       ilab=format(cbind(AVO2EXMOD$mean.pre, AVO2EXMOD$sd.pre, AVO2EXMOD$mean.post, AVO2EXMOD$sd.post, AVO2EXMOD$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(AEROBIC_REM_AVO2EXMOD, row=68.5, cex=0.4, col=11, mlab="")
text(-7.72, 68.5, cex=0.4, bquote(paste("RE Model for Aerobic, Upper-Body: p < 0.004",
                                     "; ", I^2, " = ", .(formatC(AEROBIC_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(RESISTANCE_REM_AVO2EXMOD, row=60.5, cex=0.4, col=11, mlab="")
text(-8.4, 60.5, cex=0.4, bquote(paste("RE Model for Resistance: p = 0.003",
                                     "; ", I^2, " = ", .(formatC(RESISTANCE_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(FES_REM_AVO2EXMOD, row=39.5, cex=0.4, col=11, mlab="")
text(-8.7, 39.5, cex=0.4, bquote(paste("RE Model for FES: p < 0.004",
                                     "; ", I^2, " = ", .(formatC(FES_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(GAIT_REM_AVO2EXMOD, row=25.5, cex=0.4, col=11, mlab="")
text(-8.38, 25.5, cex=0.4, bquote(paste("RE Model for Gait Training: p = 0.14",
                                     "; ", I^2, " = ", .(formatC(GAIT_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_AVO2EXMOD, row=7.5, cex=0.4, col=11, mlab="")
text(-8.6, 7.5, cex=0.4, bquote(paste("RE Model for Mixed: p < 0.004",
                                     "; ", I^2, " = ", .(formatC(MIXED_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(BEVC_REM_AVO2EXMOD, row=0.5, cex=0.4, col=11, mlab="")
text(-8.05, 0.7, cex=0.4, bquote(paste("RE Model for Behaviour Change: p = 0.10",
                                    "; ", I^2, " = ", .(formatC(BEVC_REM_AVO2EXMOD$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 99, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 100, c('Pre', 'Post'), cex=0.3)
text(3.9, 99, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2EXMOD$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2EXMOD$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2EXMOD$I2, digits=1, format="f")), "%)"))) 

text(-9.12, 96.5, cex=0.44, bquote(paste(bolditalic("Aerobic, Volitional Upper-Body"))))
text(-9.87, 66.5, cex=0.44, bquote(paste(bolditalic("Resistance Training"))))
text(-9.08, 58.5, cex=0.44, bquote(paste(bolditalic("Functional Electrical Stimulation"))))
text(-10.32, 37.5, cex=0.44, bquote(paste(bolditalic("Gait Training"))))
text(-9.49, 23.5, cex=0.44, bquote(paste(bolditalic("Hybrid, Mixed, Multimodal"))))
text(-9.95, 5.5, cex=0.44, bquote(paste(bolditalic("Behaviour Change"))))

text(-7.7, -2, cex=0.4, "Test for Subgroup Differences: Q = 10.12, df = 5, p = 0.07")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green","black")[match(AVO2EXMOD$subgroup, c("aerobic", "resistance", "fes", "gait", "mixed", "behaviour change"))] 
funnelplotdata <- funnel(REM_AVO2EXMOD, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Aerobic","Resistance","FES","Gait","Mixed/hybrid","Behaviour change"),cex=0.6, pch=20, col=c("purple","blue","orange","red","green","black"))

