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

PPOCLASS <- read.xlsx("PPO_DATA.xlsx", sheet = 2)

PPOCLASS$n.post <- as.numeric(PPOCLASS$n.post)
PPOCLASS$mean.post <- as.numeric(PPOCLASS$mean.post)
PPOCLASS$sd.post <- as.numeric(PPOCLASS$sd.post)
PPOCLASS$n.pre <- as.numeric(PPOCLASS$n.pre)
PPOCLASS$mean.pre <- as.numeric(PPOCLASS$mean.pre)
PPOCLASS$sd.pre <- as.numeric(PPOCLASS$sd.pre)

# Naming subgroups
PPOCLASS$subgroup <- as.factor(PPOCLASS$subgroup)

# Effect size calculation
PPOCLASS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=PPOCLASS,var.names=c('PPOCLASS_MD','PPOCLASS_VAR'))

# Overall random effects model
REM_PPOCLASS <- rma(PPOCLASS_MD, PPOCLASS_VAR, data=PPOCLASS_ESCALC, digits=3, 
                     slab=study)
REM_PPOCLASS

# Subgroups random effects models
TETRA_REM_PPOCLASS <- rma(PPOCLASS_MD, PPOCLASS_VAR, data=PPOCLASS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='tetraplegia'))
TETRA_REM_PPOCLASS

PARA_REM_PPOCLASS <- rma(PPOCLASS_MD, PPOCLASS_VAR, data=PPOCLASS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='paraplegia'))
PARA_REM_PPOCLASS

MIXED_REM_PPOCLASS <- rma(PPOCLASS_MD, PPOCLASS_VAR, data=PPOCLASS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='mixed'))
MIXED_REM_PPOCLASS

NR.CD_REM_PPOCLASS <- rma(PPOCLASS_MD, PPOCLASS_VAR, data=PPOCLASS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_PPOCLASS

# Test for subgroup differences
PPOSEI <- sqrt(PPOCLASS_ESCALC$PPOCLASS_VAR)
SGDIFF <- metagen(PPOCLASS_MD, PPOSEI, studlab=PPOCLASS$study, data=PPOCLASS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOCLASS, ylim = c(-1, 79), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(74:72, 67:46, 41:8, 3:2), digits = c(0,0),
       ilab=format(cbind(PPOCLASS$mean.pre,
                         PPOCLASS$sd.pre,
                         PPOCLASS$mean.post,
                         PPOCLASS$sd.post,
                         PPOCLASS$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(TETRA_REM_PPOCLASS, row=70.5, cex=0.5, col=11, mlab="")
text(-118, 70.5, cex=0.5, bquote(paste("RE Model for Tetraplegia: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(TETRA_REM_PPOCLASS$I2, digits = 1, format="f")), "%")))

addpoly(PARA_REM_PPOCLASS, row=44.5, cex=0.5, col=11, mlab="")
text(-117.1, 44.5, cex=0.5, bquote(paste("RE Model for Paraplegia: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(PARA_REM_PPOCLASS$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_PPOCLASS, row=6.5, cex=0.5, col=11, mlab="")
text(-122.2, 6.5, cex=0.5, bquote(paste("RE Model for Mixed: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_PPOCLASS$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_PPOCLASS, row=0.5, cex=0.5, col=11, mlab="")
text(-122.4, 0.5, cex=0.5, bquote(paste("RE Model for NR/CD: p = 0.001",
                                      "; ", I^2, " = ", .(formatC(NR.CD_REM_PPOCLASS$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 78, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 79, c('Pre', 'Post'), cex=0.38)
text(84, 78.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOCLASS$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOCLASS$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOCLASS$I2, digits=1, format="f")), "%)"))) 

text(-159.5, 75.5, cex=0.5, bquote(paste(bolditalic("Tetraplegia"))))
text(-160, 68.5, cex=0.5, bquote(paste(bolditalic("Paraplegia"))))
text(-166, 42.5, cex=0.5, bquote(paste(bolditalic("Mixed"))))
text(-134, 4.5, cex=0.5, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-102, -2, cex=0.48, "Test for Subgroup Differences: Q = 24.06, df = 3, p < 0.001")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(PPOCLASS$subgroup, c("tetraplegia", "paraplegia", "mixed", "not reported/cannot determine"))]
funnelplotdata <- funnel(REM_PPOCLASS, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Tetraplegia","Paraplegia","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
