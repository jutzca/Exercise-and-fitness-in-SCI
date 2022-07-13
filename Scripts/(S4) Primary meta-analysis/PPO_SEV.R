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

PPOSEV <- read.xlsx("PPO_DATA.xlsx", sheet = 3)

PPOSEV$n.post <- as.numeric(PPOSEV$n.post)
PPOSEV$mean.post <- as.numeric(PPOSEV$mean.post)
PPOSEV$sd.post <- as.numeric(PPOSEV$sd.post)
PPOSEV$n.pre <- as.numeric(PPOSEV$n.pre)
PPOSEV$mean.pre <- as.numeric(PPOSEV$mean.pre)
PPOSEV$sd.pre <- as.numeric(PPOSEV$sd.pre)

# Naming subgroups
PPOSEV$subgroup <- as.factor(PPOSEV$subgroup)

# Effect size calculation
PPOSEV_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=PPOSEV,var.names=c('PPOSEV_MD','PPOSEV_VAR'))

# Overall random effects model
REM_PPOSEV <- rma(PPOSEV_MD, PPOSEV_VAR, data=PPOSEV_ESCALC, digits=3, 
                   slab=study)
REM_PPOSEV

# Subgroups random effects models
COMP_REM_PPOSEV <- rma(PPOSEV_MD, PPOSEV_VAR, data=PPOSEV_ESCALC, digits=3, 
                        slab=study, subset=(subgroup=='comp'))
COMP_REM_PPOSEV

INCOMP_REM_PPOSEV<- rma(PPOSEV_MD, PPOSEV_VAR, data=PPOSEV_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='incomp'))
INCOMP_REM_PPOSEV

MIXED_REM_PPOSEV <- rma(PPOSEV_MD, PPOSEV_VAR, data=PPOSEV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_PPOSEV

NR.CD_REM_PPOSEV <- rma(PPOSEV_MD, PPOSEV_VAR, data=PPOSEV_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_PPOSEV

# Test for subgroup differences
PPOSEI <- sqrt(PPOSEV_ESCALC$PPOSEV_VAR)
SGDIFF <- metagen(PPOSEV_MD, PPOSEI, studlab=PPOSEV$study, data=PPOSEV_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOSEV, ylim = c(-1, 79), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(74:51, 46:45, 40:16, 11:2), digits = c(0,0),
       ilab=format(cbind(PPOSEV$mean.pre,
                         PPOSEV$sd.pre,
                         PPOSEV$mean.post,
                         PPOSEV$sd.post,
                         PPOSEV$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col = 10)

addpoly(COMP_REM_PPOSEV, row=49.5, cex=0.5, col=11, mlab="")
text(-109.5, 49.5, cex=0.5, bquote(paste("RE Model for Motor-Complete: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(COMP_REM_PPOSEV$I2, digits = 1, format="f")), "%")))

addpoly(INCOMP_REM_PPOSEV, row=43.5, cex=0.5, col=11, mlab="")
text(-110.8, 43.5, cex=0.5, bquote(paste("RE Model for Motor-Incomplete: p = 0.25",
                                       "; ", I^2, " = ", .(formatC(INCOMP_REM_PPOSEV$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_PPOSEV, row=14.5, cex=0.5, col=11, mlab="")
text(-122.1, 14.5, cex=0.5, bquote(paste("RE Model for Mixed: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_PPOSEV$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_PPOSEV, row=0.5, cex=0.5, col=11, mlab="")
text(-121.5, 0.5, cex=0.5, bquote(paste("RE Model for NR/CD: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(NR.CD_REM_PPOSEV$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 78, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 79, c('Pre', 'Post'), cex=0.38)
text(84, 78.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOSEV$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOSEV$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOSEV$I2, digits=1, format="f")), "%)"))) 

text(-138.8, 75.5, cex=0.5, bquote(paste(bolditalic("Motor-Complete (AIS A-B)"))))
text(-136.8, 47.5, cex=0.5, bquote(paste(bolditalic("Motor-Incomplete (AIS C-D)"))))
text(-152.4, 41.5, cex=0.5, bquote(paste(bolditalic("Mixed (AIS A-D)"))))
text(-134, 12.5, cex=0.5, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-104.5, -2, cex=0.48, "Test for Subgroup Differences: Q = 2.73, df = 3, p = 0.43")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(PPOSEV$subgroup, c("comp", "incomp", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_PPOSEV, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Motor-complete","Motor-incomplete","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
