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

RVO2EXMOD <- read.xlsx("RVO2_DATA.xlsx", sheet = 4)

RVO2EXMOD$n.post <- as.integer(RVO2EXMOD$n.post)
RVO2EXMOD$mean.post <- as.numeric(RVO2EXMOD$mean.post)
RVO2EXMOD$sd.post <- as.numeric(RVO2EXMOD$sd.post)
RVO2EXMOD$n.pre <- as.integer(RVO2EXMOD$n.pre)
RVO2EXMOD$mean.pre <- as.numeric(RVO2EXMOD$mean.pre)
RVO2EXMOD$sd.pre <- as.numeric(RVO2EXMOD$sd.pre)

# Naming subgroups
RVO2EXMOD$subgroup <- as.factor(RVO2EXMOD$subgroup)

# Effect size calculation
RVO2EXMOD_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=RVO2EXMOD,var.names=c('RVO2EXMOD_MD','RVO2EXMOD_VAR'))

# Overall random effects model
REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                   slab=study)
REM_RVO2EXMOD

# Subgroups random effects models
AEROBIC_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                        slab=study, subset=(subgroup=='aerobic'))
AEROBIC_REM_RVO2EXMOD

RESISTANCE_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='resistance'))
RESISTANCE_REM_RVO2EXMOD

FES_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='fes'))
FES_REM_RVO2EXMOD

GAIT_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='gait'))
GAIT_REM_RVO2EXMOD

MIXED_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='mixed'))
MIXED_REM_RVO2EXMOD

BEVC_REM_RVO2EXMOD <- rma(RVO2EXMOD_MD, RVO2EXMOD_VAR, data=RVO2EXMOD_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='behaviour change'))
BEVC_REM_RVO2EXMOD

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2EXMOD_ESCALC$RVO2EXMOD_VAR)
SGDIFF <- metagen(RVO2EXMOD_MD, RVO2SEI, data=RVO2EXMOD_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2EXMOD, ylim = c(-1, 100), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(95:63, 58:56, 51:44, 39:30, 25:8, 3:2), digits = 1,
       ilab=format(cbind(RVO2EXMOD$mean.pre,
                         RVO2EXMOD$sd.pre,
                         RVO2EXMOD$mean.post,
                         RVO2EXMOD$sd.post,
                         RVO2EXMOD$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(AEROBIC_REM_RVO2EXMOD, row=61.5, cex=0.4, col=11, mlab="")
text(-55.3, 61.5, cex=0.38, bquote(paste("RE Model for Aerobic, Upper-Body: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(AEROBIC_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(RESISTANCE_REM_RVO2EXMOD, row=54.5, cex=0.4, col=11, mlab="")
text(-60.3, 54.5, cex=0.38, bquote(paste("RE Model for Resistance: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(RESISTANCE_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(FES_REM_RVO2EXMOD, row=42.5, cex=0.4, col=11, mlab="")
text(-63.3, 42.5, cex=0.38, bquote(paste("RE Model for FES: p = 0.006",
                                         "; ", I^2, " = ", .(formatC(FES_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(GAIT_REM_RVO2EXMOD, row=28.5, cex=0.4, col=11, mlab="")
text(-60.6, 28.5, cex=0.38, bquote(paste("RE Model for Gait Training: p = 0.40",
                                         "; ", I^2, " = ", .(formatC(GAIT_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_RVO2EXMOD, row=6.5, cex=0.4, col=11, mlab="")
text(-63.6, 6.5, cex=0.38, bquote(paste("RE Model for Mixed: p = 0.35",
                                         "; ", I^2, " = ", .(formatC(MIXED_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))

addpoly(BEVC_REM_RVO2EXMOD, row=0.5, cex=0.4, col=11, mlab="")
text(-57.5, 0.5, cex=0.38, bquote(paste("RE Model for Behaviour Change: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(BEVC_REM_RVO2EXMOD$I2, digits = 1, format="f")), "%")))


text(c(-40, -35, -30, -25, -20), 99, cex=0.38, c("Mean", "SD", "Mean", "SD", "Total N"))
text(c(-37.5, -27.5), 100, cex=0.38, c("Pre", "Post"))
text(38.5, 99, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2EXMOD$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2EXMOD$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2EXMOD$I2, digits=1, format="f")), "%)"))) 

text(-66.5, 96.5, cex=0.4, bquote(paste(bolditalic("Aerobic, Volitional Upper-Body"))))
text(-72.1, 59.5, cex=0.4, bquote(paste(bolditalic("Resistance Training"))))
text(-65.8, 52.5, cex=0.4, bquote(paste(bolditalic("Functional Electrical Stimulation"))))
text(-75.8, 40.5, cex=0.4, bquote(paste(bolditalic("Gait Training"))))
text(-69.3, 26.5, cex=0.4, bquote(paste(bolditalic("Hybrid, Mixed, Multimodal"))))
text(-72.8, 4.5, cex=0.4, bquote(paste(bolditalic("Behaviour Change"))))

text(-55.5, -2, cex=0.38, "Test for Subgroup Differences: Q = 12.95, df = 5, p = 0.02")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green","black")[match(RVO2EXMOD$subgroup, c("aerobic", "resistance", "fes", "gait", "mixed", "behaviour change"))] 
funnelplotdata <- funnel(REM_RVO2EXMOD, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Aerobic","Resistance","FES","Gait","Mixed/hybrid","Behaviour change"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green","black"))
