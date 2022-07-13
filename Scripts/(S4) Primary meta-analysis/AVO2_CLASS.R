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

AVO2CLASS <- read.xlsx("AVO2_DATA.xlsx", sheet = 2)

AVO2CLASS$mean.post <- as.numeric(AVO2CLASS$mean.post)
AVO2CLASS$sd.post <- as.numeric(AVO2CLASS$sd.post)
AVO2CLASS$n.post <- as.numeric(AVO2CLASS$n.post)
AVO2CLASS$mean.pre <- as.numeric(AVO2CLASS$mean.pre)
AVO2CLASS$sd.pre <- as.numeric(AVO2CLASS$sd.pre)
AVO2CLASS$n.pre <- as.numeric(AVO2CLASS$n.pre)


# Naming subgroups
AVO2CLASS$subgroup <- as.factor(AVO2CLASS$subgroup)

# Effect size calculation
AVO2CLASS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2CLASS,var.names=c('AVO2CLASS_MD','AVO2CLASS_VAR'))

# Overall random effects model
REM_AVO2CLASS <- rma(AVO2CLASS_MD, AVO2CLASS_VAR, data=AVO2CLASS_ESCALC, digits=3, 
                   slab=study)
REM_AVO2CLASS

# Subgroups random effects models
TETRA_REM_AVO2CLASS <- rma(AVO2CLASS_MD, AVO2CLASS_VAR, data=AVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='tetra'))
TETRA_REM_AVO2CLASS

PARA_REM_AVO2CLASS <- rma(AVO2CLASS_MD, AVO2CLASS_VAR, data=AVO2CLASS_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='para'))
PARA_REM_AVO2CLASS

MIXED_REM_AVO2CLASS <- rma(AVO2CLASS_MD, AVO2CLASS_VAR, data=AVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='mixed'))
MIXED_REM_AVO2CLASS

NR.CD_REM_AVO2CLASS <- rma(AVO2CLASS_MD, AVO2CLASS_VAR, data=AVO2CLASS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_AVO2CLASS

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2CLASS_ESCALC$AVO2CLASS_VAR)
SGDIFF <- metagen(AVO2CLASS_MD, AVO2SEI, studlab=AVO2CLASS$study, data=AVO2CLASS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2CLASS, ylim = c(-1,87), xlim = c(-11.5,7), alim = c(-2,2), rows = c(82:81, 76:57, 52:9, 4:2),
       ilab=format(cbind(AVO2CLASS$mean.pre, AVO2CLASS$sd.pre, AVO2CLASS$mean.post, AVO2CLASS$sd.post, AVO2CLASS$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(TETRA_REM_AVO2CLASS, row=79.5, cex=0.4, col=11, mlab="")
text(-8.4, 79.5, cex=0.4, bquote(paste("RE Model for Tetraplegia: p = 0.23",
                                       "; ", I^2, " = ", .(formatC(TETRA_REM_AVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(PARA_REM_AVO2CLASS, row=55.5, cex=0.4, col=11, mlab="")
text(-8.35, 55.5, cex=0.4, bquote(paste("RE Model for Paraplegia: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(PARA_REM_AVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_AVO2CLASS, row=7.5, cex=0.4, col=11, mlab="")
text(-8.6, 7.5, cex=0.4, bquote(paste("RE Model for Mixed: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_AVO2CLASS$I2, digits = 1, format="f")), "%")))

addpoly(NR.CD_REM_AVO2CLASS, row=0.5, cex=0.4, col=11, mlab="")
text(-8.52, 0.7, cex=0.4, bquote(paste("RE Model for NR/CD: p < 0.002",
                                      "; ", I^2, " = ", .(formatC(NR.CD_REM_AVO2CLASS$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 86, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 87, c('Pre', 'Post'), cex=0.3)
text(3.9, 86, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2CLASS$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2CLASS$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2CLASS$I2, digits=1, format="f")), "%)"))) 

text(-10.45, 83.5, cex=0.44, bquote(paste(bolditalic("Tetraplegia"))))
text(-10.45, 77.5, cex=0.44, bquote(paste(bolditalic("Paraplegia"))))
text(-10.7, 53.5, cex=0.44, bquote(paste(bolditalic("Mixed"))))
text(-9.19, 5.5, cex=0.44, bquote(paste(bolditalic("Not reported/Cannot determine"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 1.64, df = 3, p = 0.65")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(AVO2CLASS$subgroup, c("tetra", "para", "mixed", "not reported/cannot determine"))]
funnelplotdata <- funnel(REM_AVO2CLASS, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Tetraplegia","Paraplegia","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))

