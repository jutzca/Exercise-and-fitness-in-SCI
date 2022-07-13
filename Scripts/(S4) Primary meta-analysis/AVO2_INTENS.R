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

AVO2INTENS <- read.xlsx("AVO2_DATA.xlsx", sheet = 5)

AVO2INTENS$n.post <- as.numeric(AVO2INTENS$n.post)
AVO2INTENS$mean.post <- as.numeric(AVO2INTENS$mean.post)
AVO2INTENS$sd.post <- as.numeric(AVO2INTENS$sd.post)
AVO2INTENS$n.pre <- as.numeric(AVO2INTENS$n.pre)
AVO2INTENS$mean.pre <- as.numeric(AVO2INTENS$mean.pre)
AVO2INTENS$sd.pre <- as.numeric(AVO2INTENS$sd.pre)

# Naming subgroups
AVO2INTENS$subgroup <- as.factor(AVO2INTENS$subgroup)

# Effect size calculation
AVO2INTENS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=AVO2INTENS,var.names=c('AVO2INTENS_MD','AVO2INTENS_VAR'))

# Overall random effects model
REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                      slab=study)
REM_AVO2INTENS

# Subgroups random effects models
LIGHT_REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='light'))
LIGHT_REM_AVO2INTENS

MOD_REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='mod'))
MOD_REM_AVO2INTENS

MODVIG_REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='modvig'))
MODVIG_REM_AVO2INTENS

VIG_REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='vig'), control=list(stepadj=0.5, maxiter=10000))
VIG_REM_AVO2INTENS

MIXED_REM_AVO2INTENS <- rma(AVO2INTENS_MD, AVO2INTENS_VAR, data=AVO2INTENS_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='mixed/cd'))
MIXED_REM_AVO2INTENS

AVO2SEI <- sqrt(AVO2INTENS_ESCALC$AVO2INTENS_VAR)
SGDIFF <- metagen(AVO2INTENS_MD, AVO2SEI, studlab=AVO2INTENS$study, data=AVO2INTENS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup,
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2INTENS, ylim = c(-1,91), xlim = c(-11.5,7), alim = c(-2,2), rows = c(86:86, 81:74, 69:49, 44:31, 26:2), cex = 0.4,
       ilab=format(cbind(AVO2INTENS$mean.pre, AVO2INTENS$sd.pre, AVO2INTENS$mean.post, AVO2INTENS$sd.post, AVO2INTENS$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(LIGHT_REM_AVO2INTENS, row=84.5, cex=0.4, col=11, mlab="")
text(-8.8, 84.5, cex=0.4, bquote(paste("RE Model for Light: p = 0.85",
                                        "; ", I^2, " = ", .(formatC(LIGHT_REM_AVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(MOD_REM_AVO2INTENS, row=72.5, cex=0.4, col=11, mlab="")
text(-8.5, 72.5, cex=0.4, bquote(paste("RE Model for Moderate: p = 0.01",
                                        "; ", I^2, " = ", .(formatC(MOD_REM_AVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(MODVIG_REM_AVO2INTENS, row=47.5, cex=0.4, col=11, mlab="")
text(-7.62, 47.5, cex=0.4, bquote(paste("RE Model for Moderate-to-Vigorous: p < 0.003",
                                        "; ", I^2, " = ", .(formatC(MODVIG_REM_AVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(VIG_REM_AVO2INTENS, row=29.5, cex=0.4, col=11, mlab="")
text(-8.52, 29.5, cex=0.4, bquote(paste("RE Model for Vigorous: p < 0.003",
                                        "; ", I^2, " = ", .(formatC(VIG_REM_AVO2INTENS$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_AVO2INTENS, row=0.5, cex=0.4, col=11, mlab="")
text(-7.5, 0.7, cex=0.4, bquote(paste("RE Model for Mixed/Cannot Determine: p < 0.003",
                                        "; ", I^2, " = ", .(formatC(MIXED_REM_AVO2INTENS$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 90, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 91, c('Pre', 'Post'), cex=0.3)
text(3.9, 90, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2INTENS$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2INTENS$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2INTENS$I2, digits=1, format="f")), "%)"))) 

text(-10.75, 87.5, cex=0.44, bquote(paste(bolditalic("Light"))))
text(-10.5, 82.5, cex=0.44, bquote(paste(bolditalic("Moderate"))))
text(-9.65, 70.5, cex=0.44, bquote(paste(bolditalic("Moderate-to-Vigorous"))))
text(-10.57, 45.5, cex=0.44, bquote(paste(bolditalic("Vigorous"))))
text(-9.55, 27.5, cex=0.44, bquote(paste(bolditalic("Mixed/Cannot Determine"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 2.12, df = 4, p = 0.71")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green")[match(AVO2INTENS$subgroup,c("light", "mod", "modvig", "vig", "mixed/cd"))]
funnelplotdata <- funnel(REM_AVO2INTENS, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Light","Moderate","Mod-to-Vig","Vigorous","Mixed/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green"))
