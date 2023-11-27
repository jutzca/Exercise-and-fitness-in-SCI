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

PPOINTENS <- read.xlsx("PPO_DATA_mixedout.xlsx", sheet = 5)

PPOINTENS$n.post <- as.numeric(PPOINTENS$n.post)
PPOINTENS$mean.post <- as.numeric(PPOINTENS$mean.post)
PPOINTENS$sd.post <- as.numeric(PPOINTENS$sd.post)
PPOINTENS$n.pre <- as.numeric(PPOINTENS$n.pre)
PPOINTENS$mean.pre <- as.numeric(PPOINTENS$mean.pre)
PPOINTENS$sd.pre <- as.numeric(PPOINTENS$sd.pre)

# Naming subgroups
PPOINTENS$subgroup <- as.factor(PPOINTENS$subgroup)

# Effect size calculation
PPOINTENS_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=PPOINTENS,var.names=c('PPOINTENS_MD','PPOINTENS_VAR'))

# Overall random effects model
REM_PPOINTENS<- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                      slab=study)
REM_PPOINTENS

# Subgroups random effects models
LIGHT_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='light'))
LIGHT_REM_PPOINTENS

MOD_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='mod'))
MOD_REM_PPOINTENS

MODVIG_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='modvig'))
MODVIG_REM_PPOINTENS

VIG_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='vig'))
VIG_REM_PPOINTENS

SUPRA_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='supra'))
SUPRA_REM_PPOINTENS

MIXED_REM_PPOINTENS <- rma(PPOINTENS_MD, PPOINTENS_VAR, data=PPOINTENS_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='mixed/cd'))
MIXED_REM_PPOINTENS

# Test for subgroup differences
PPOSEI <- sqrt(PPOINTENS_ESCALC$PPOINTENS_VAR)
SGDIFF <- metagen(PPOINTENS_MD, PPOSEI, studlab=PPOINTENS$study, data=PPOINTENS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup,
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOINTENS, ylim = c(-1, 63), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(58:58, 53:44, 39:22, 17:7, 2:2), digits = c(0,0), cex=0.5,
       ilab=format(cbind(PPOINTENS$mean.pre,
                         PPOINTENS$sd.pre,
                         PPOINTENS$mean.post,
                         PPOINTENS$sd.post,
                         PPOINTENS$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(LIGHT_REM_PPOINTENS, row=56.5, cex=0.5, col=11, mlab="")
text(-126.2, 56.5, cex=0.5, bquote(paste("RE Model for Light: p = 0.92",
                                         "; ", I^2, " = ", .(formatC(LIGHT_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(MOD_REM_PPOINTENS, row=42.5, cex=0.5, col=11, mlab="")
text(-118, 42.5, cex=0.5, bquote(paste("RE Model for Moderate: p = 0.009",
                                         "; ", I^2, " = ", .(formatC(MOD_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(MODVIG_REM_PPOINTENS, row=20.5, cex=0.5, col=11, mlab="")
text(-101.2, 20.5, cex=0.5, bquote(paste("RE Model for Moderate-to-Vigorous: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(MODVIG_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(VIG_REM_PPOINTENS, row=5.5, cex=0.5, col=11, mlab="")
text(-118.3, 5.5, cex=0.5, bquote(paste("RE Model for Vigorous: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(VIG_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(SUPRA_REM_PPOINTENS, row=0.5, cex=0.5, col=11, mlab="")
text(-115, 0.7, cex=0.5, bquote(paste("RE Model for Supramaximal: p = 0.50",
                                         "; ", I^2, " = ", .(formatC(SUPRA_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 62, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 63, c('Pre', 'Post'), cex=0.38)
text(84, 62.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOINTENS$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOINTENS$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOINTENS$I2, digits=1, format="f")), "%)"))) 

text(-166.3, 59.5, cex=0.5, bquote(paste(bolditalic("Light"))))
text(-161.7, 54.5, cex=0.5, bquote(paste(bolditalic("Moderate"))))
text(-144, 40.5, cex=0.5, bquote(paste(bolditalic("Moderate-to-Vigorous"))))
text(-161.8, 18.5, cex=0.5, bquote(paste(bolditalic("Vigorous"))))
text(-154.8, 3.5, cex=0.5, bquote(paste(bolditalic("Supramaximal"))))

text(-104, -2, cex=0.48, "Test for Subgroup Differences: Q = 6.87, df = 4, p = 0.14")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green")[match(PPOINTENS$subgroup,c("light", "mod", "modvig", "vig", "supra"))]
funnelplotdata <- funnel(REM_PPOINTENS, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Light","Moderate","Mod-to-Vig","Vigorous","Supramaximal"),cex=0.5, pch=20, col=c("purple","blue","orange","red","green"))

