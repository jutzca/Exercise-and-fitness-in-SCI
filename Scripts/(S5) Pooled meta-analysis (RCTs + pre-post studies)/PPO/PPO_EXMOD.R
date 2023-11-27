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

PPOEXMOD <- read.xlsx("PPO_DATA_SEPT.xlsx", sheet = 4)

PPOEXMOD$n.post <- as.numeric(PPOEXMOD$n.post)
PPOEXMOD$mean.post <- as.numeric(PPOEXMOD$mean.post)
PPOEXMOD$sd.post <- as.numeric(PPOEXMOD$sd.post)
PPOEXMOD$n.pre <- as.numeric(PPOEXMOD$n.pre)
PPOEXMOD$mean.pre <- as.numeric(PPOEXMOD$mean.pre)
PPOEXMOD$sd.pre <- as.numeric(PPOEXMOD$sd.pre)

# Naming subgroups
PPOEXMOD$subgroup <- as.factor(PPOEXMOD$subgroup)

# Effect size calculation
PPOEXMOD_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=PPOEXMOD,var.names=c('PPOEXMOD_MD','PPOEXMOD_VAR'))

# Overall random effects model
REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                     slab=study)
REM_PPOEXMOD

# Subgroups random effects models
AEROBIC_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='aerobic'))
AEROBIC_REM_PPOEXMOD

RESISTANCE_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='resistance'))
RESISTANCE_REM_PPOEXMOD

FES_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='fes'))
FES_REM_PPOEXMOD

GAIT_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='gait'))
GAIT_REM_PPOEXMOD

MIXED_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='mixed'))
MIXED_REM_PPOEXMOD

BEVC_REM_PPOEXMOD <- rma(PPOEXMOD_MD, PPOEXMOD_VAR, data=PPOEXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='behaviour change'))
BEVC_REM_PPOEXMOD

# Test for subgroup differences
PPOSEI <- sqrt(PPOEXMOD_ESCALC$PPOEXMOD_VAR)
SGDIFF <- metagen(PPOEXMOD_MD, PPOSEI, studlab=PPOEXMOD$study, data=PPOEXMOD_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOEXMOD, ylim = c(-1, 91), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(86:61, 56:54, 49:36, 31:30, 25:11, 6:2), digits = c(0,0), cex=0.5,
       ilab=format(cbind(PPOEXMOD$mean.pre,
                         PPOEXMOD$sd.pre,
                         PPOEXMOD$mean.post,
                         PPOEXMOD$sd.post,
                         PPOEXMOD$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col = 10)

addpoly(AEROBIC_REM_PPOEXMOD, row=59.5, cex=0.5, col=11, mlab="")
text(-104, 59.5, cex=0.5, bquote(paste("RE Model for Aerobic, Upper-Body: p < 0.003",
                                       "; ", I^2, " = ", .(formatC(AEROBIC_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

addpoly(RESISTANCE_REM_PPOEXMOD, row=52.5, cex=0.5, col=11, mlab="")
text(-117.8, 52.5, cex=0.5, bquote(paste("RE Model for Resistance: p < 0.003",
                                       "; ", I^2, " = ", .(formatC(RESISTANCE_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

addpoly(FES_REM_PPOEXMOD, row=34.5, cex=0.5, col=11, mlab="")
text(-124, 34.5, cex=0.5, bquote(paste("RE Model for FES: p < 0.003",
                                       "; ", I^2, " = ", .(formatC(FES_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

addpoly(GAIT_REM_PPOEXMOD, row=28.5, cex=0.5, col=11, mlab="")
text(-117.4, 28.5, cex=0.5, bquote(paste("RE Model for Gait Training: p = 0.54",
                                       "; ", I^2, " = ", .(formatC(GAIT_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_PPOEXMOD, row=9.5, cex=0.5, col=11, mlab="")
text(-122, 9.5, cex=0.5, bquote(paste("RE Model for Mixed: p < 0.003",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

addpoly(BEVC_REM_PPOEXMOD, row=0.5, cex=0.5, col=11, mlab="")
text(-108.5, 0.7, cex=0.5, bquote(paste("RE Model for Behaviour Change: p = 0.14",
                                      "; ", I^2, " = ", .(formatC(BEVC_REM_PPOEXMOD$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 90, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 91, c('Pre', 'Post'), cex=0.38)
text(84, 90.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOEXMOD$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOEXMOD$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOEXMOD$I2, digits=1, format="f")), "%)"))) 

text(-133, 87.5, cex=0.5, bquote(paste(bolditalic("Aerobic, Volitional Upper-Body"))))
text(-148, 57.5, cex=0.5, bquote(paste(bolditalic("Resistance Training"))))
text(-131, 50.5, cex=0.5, bquote(paste(bolditalic("Functional Electrical Stimulation"))))
text(-157, 32.5, cex=0.5, bquote(paste(bolditalic("Gait Training"))))
text(-140, 26.5, cex=0.5, bquote(paste(bolditalic("Hybrid, Mixed, Multimodal"))))
text(-149.7, 7.5, cex=0.5, bquote(paste(bolditalic("Behaviour Change"))))

text(-102, -2, cex=0.48, "Test for Subgroup Differences: Q = 18.99, df = 5, p = 0.002")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green","black")[match(PPOEXMOD$subgroup, c("aerobic", "resistance", "fes", "gait", "mixed", "behaviour change"))] 
funnelplotdata <- funnel(REM_PPOEXMOD, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Aerobic","Resistance","FES","Gait","Mixed/hybrid","Behaviour change"),cex=0.5, pch=20, col=c("purple","blue","orange","red","green","black"))