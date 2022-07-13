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

RVO2CLASS <- read.xlsx("RVO2_DATA.xlsx", sheet = 2)

# Set as numeric
RVO2CLASS$n.post <- as.integer(RVO2CLASS$n.post)
RVO2CLASS$mean.post <- as.numeric(RVO2CLASS$mean.post)
RVO2CLASS$sd.post <- as.numeric(RVO2CLASS$sd.post)
RVO2CLASS$n.pre <- as.integer(RVO2CLASS$n.pre)
RVO2CLASS$mean.pre <- as.numeric(RVO2CLASS$mean.pre)
RVO2CLASS$sd.pre <- as.numeric(RVO2CLASS$sd.pre)

# Naming subgroups
RVO2CLASS$subgroup <- as.factor(RVO2CLASS$subgroup)

# Effect size calculation
RVO2CLASS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=RVO2CLASS,var.names=c('RVO2CLASS_MD','RVO2CLASS_VAR'))

# Overall random effects model
REM_RVO2CLASS <- rma(RVO2CLASS_MD, RVO2CLASS_VAR, data=RVO2CLASS_ESCALC, digits=3, 
                   slab=study)
REM_RVO2CLASS

# Subgroups random effects models
TETRA_REM_RVO2CLASS <- rma(RVO2CLASS_MD, RVO2CLASS_VAR, data=RVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='tetraplegia'))
TETRA_REM_RVO2CLASS

PARA_REM_RVO2CLASS <- rma(RVO2CLASS_MD, RVO2CLASS_VAR, data=RVO2CLASS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='paraplegia'))
PARA_REM_RVO2CLASS

MIXED_REM_RVO2CLASS <- rma(RVO2CLASS_MD, RVO2CLASS_VAR, data=RVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_RVO2CLASS

NR.CD_REM_RVO2CLASS <- rma(RVO2CLASS_MD, RVO2CLASS_VAR, data=RVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_RVO2CLASS

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2CLASS_ESCALC$RVO2CLASS_VAR)
SGDIFF <- metagen(RVO2CLASS_MD, RVO2SEI, data=RVO2CLASS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

# Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2CLASS, ylim = c(-1, 92), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(87:85, 80:53, 48:8, 3:2), digits = 1,
       ilab=format(cbind(RVO2CLASS$mean.pre,
                         RVO2CLASS$sd.pre,
                         RVO2CLASS$mean.post,
                         RVO2CLASS$sd.post,
                         RVO2CLASS$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)
        
addpoly(TETRA_REM_RVO2CLASS, row=83.5, cex=0.35, col=11, mlab="")
text(-60.95, 83.5, cex=0.38, bquote(paste("RE Model for Tetraplegia: p = 0.04",
                                         "; ", I^2, " = ", .(formatC(TETRA_REM_RVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(PARA_REM_RVO2CLASS, row=51.5, cex=0.4, col=11, mlab="") 
text(-60.3, 51.5, cex=0.38, bquote(paste("RE Model for Paraplegia: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(PARA_REM_RVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_RVO2CLASS, row=6.5, cex=0.4, col=11, mlab="")
text(-63, 6.5, cex=0.38, bquote(paste("RE Model for Mixed: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(MIXED_REM_RVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_RVO2CLASS, row=0.5, cex=0.4, col=11, mlab="")
text(-63, 0.5, cex=0.38, bquote(paste("RE Model for NR/CD: p = 0.02",
                                        "; ", I^2, " = ", .(formatC(MIXED_REM_RVO2CLASS$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 91, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 92, c('Pre', 'Post'), cex=0.3)
text(38.5, 91, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2CLASS$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2CLASS$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2CLASS$I2, digits=1, format="f")), "%)"))) 

text(-77, 88.5, cex=0.4, bquote(paste(bolditalic("Tetraplegia"))))
text(-77, 81.5, cex=0.4, bquote(paste(bolditalic("Paraplegia"))))
text(-79, 49.5, cex=0.4, bquote(paste(bolditalic("Mixed"))))
text(-66.8, 4.5, cex=0.4, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-55.5, -2, cex=0.38, "Test for Subgroup Differences: Q = 3.36, df = 3, p = 0.34")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(RVO2CLASS$subgroup, c("tetraplegia", "paraplegia", "mixed", "not reported/cannot determine"))]
funnelplotdata <- funnel(REM_RVO2CLASS, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Tetraplegia","Paraplegia","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
