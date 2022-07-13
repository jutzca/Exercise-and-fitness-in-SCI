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

PPOINTENS <- read.xlsx("PPO_DATA.xlsx", sheet = 5)

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
SGDIFF <- metagen(PPOINTENS_MD, PPOSEI, studlab=PPOINTENS$study, data=PPOINTENS_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOINTENS, ylim = c(-1, 87), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(82:82, 77:68, 63:48, 43:33, 28:28, 23:2), digits = c(0,0),
       ilab=format(cbind(PPOINTENS$mean.pre,
                         PPOINTENS$sd.pre,
                         PPOINTENS$mean.post,
                         PPOINTENS$sd.post,
                         PPOINTENS$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(LIGHT_REM_PPOINTENS, row=80.5, cex=0.5, col=11, mlab="")
text(-126.2, 80.5, cex=0.5, bquote(paste("RE Model for Light: p = 0.92",
                                         "; ", I^2, " = ", .(formatC(LIGHT_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(MOD_REM_PPOINTENS, row=66.5, cex=0.5, col=11, mlab="")
text(-118, 66.5, cex=0.5, bquote(paste("RE Model for Moderate: p = 0.009",
                                         "; ", I^2, " = ", .(formatC(MOD_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(MODVIG_REM_PPOINTENS, row=46.5, cex=0.5, col=11, mlab="")
text(-101.2, 46.5, cex=0.5, bquote(paste("RE Model for Moderate-to-Vigorous: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(MODVIG_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(VIG_REM_PPOINTENS, row=31.5, cex=0.5, col=11, mlab="")
text(-118.3, 31.5, cex=0.5, bquote(paste("RE Model for Vigorous: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(VIG_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(SUPRA_REM_PPOINTENS, row=26.5, cex=0.5, col=11, mlab="")
text(-115, 26.5, cex=0.5, bquote(paste("RE Model for Supramaximal: p = 0.50",
                                         "; ", I^2, " = ", .(formatC(SUPRA_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_PPOINTENS, row=0.5, cex=0.5, col=11, mlab="")
text(-99.5, 0.7, cex=0.5, bquote(paste("RE Model for Mixed/Cannot Determine: p < 0.004",
                                         "; ", I^2, " = ", .(formatC(MIXED_REM_PPOINTENS$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 86, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 87, c('Pre', 'Post'), cex=0.38)
text(84, 86.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOINTENS$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOINTENS$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOINTENS$I2, digits=1, format="f")), "%)"))) 

text(-166.3, 83.5, cex=0.5, bquote(paste(bolditalic("Light"))))
text(-161.7, 78.5, cex=0.5, bquote(paste(bolditalic("Moderate"))))
text(-144, 64.5, cex=0.5, bquote(paste(bolditalic("Moderate-to-Vigorous"))))
text(-161.8, 44.5, cex=0.5, bquote(paste(bolditalic("Vigorous"))))
text(-154.8, 29.5, cex=0.5, bquote(paste(bolditalic("Supramaximal"))))
text(-141.5, 24.5, cex=0.5, bquote(paste(bolditalic("Mixed/Cannot Determine"))))

text(-102, -2, cex=0.48, "Test for Subgroup Differences: Q = 17.99, df = 5, p = 0.003")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green","black")[match(PPOINTENS$subgroup,c("light", "mod", "modvig", "vig", "supra", "mixed/cd"))]
funnelplotdata <- funnel(REM_PPOINTENS, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Light","Moderate","Mod-to-Vig","Vigorous","Supramaximal","Mixed/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green","black"))