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

PPOVOLUME <- read.xlsx("PPO_DATA_mixedout.xlsx", sheet = 9)

PPOVOLUME$n.post <- as.numeric(PPOVOLUME$n.post)
PPOVOLUME$mean.post <- as.numeric(PPOVOLUME$mean.post)
PPOVOLUME$sd.post <- as.numeric(PPOVOLUME$sd.post)
PPOVOLUME$n.pre <- as.numeric(PPOVOLUME$n.pre)
PPOVOLUME$mean.pre <- as.numeric(PPOVOLUME$mean.pre)
PPOVOLUME$sd.pre <- as.numeric(PPOVOLUME$sd.pre)

# Naming subgroups
PPOVOLUME$subgroup <- as.factor(PPOVOLUME$subgroup)

# Effect size calculation
PPOVOLUME_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=PPOVOLUME,var.names=c('PPOVOLUME_MD','PPOVOLUME_VAR'))

# Overall random effects model
REM_PPOVOLUME <- rma(PPOVOLUME_MD, PPOVOLUME_VAR, data=PPOVOLUME_ESCALC, digits=3, 
                      slab=study)
REM_PPOVOLUME

# Subgroups random effects models
FITNESS_REM_PPOVOLUME <- rma(PPOVOLUME_MD, PPOVOLUME_VAR, data=PPOVOLUME_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='fitness'))
FITNESS_REM_PPOVOLUME

CMETAB_REM_PPOVOLUME <- rma(PPOVOLUME_MD, PPOVOLUME_VAR, data=PPOVOLUME_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='cmetab'))
CMETAB_REM_PPOVOLUME

GENPOP_REM_PPOVOLUME <- rma(PPOVOLUME_MD, PPOVOLUME_VAR, data=PPOVOLUME_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='genpop'))
GENPOP_REM_PPOVOLUME

NR.CD_REM_PPOVOLUME <- rma(PPOVOLUME_MD, PPOVOLUME_VAR, data=PPOVOLUME_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='cd'))
NR.CD_REM_PPOVOLUME

# Test for subgroup differences
PPOSEI <- sqrt(PPOVOLUME_ESCALC$PPOVOLUME_VAR)
SGDIFF <- metagen(PPOVOLUME_MD, PPOSEI, studlab=PPOVOLUME$study, data=PPOVOLUME_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOVOLUME, ylim = c(-1, 70), xlim = c(-190, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(65:50, 45:19, 14:2), digits = c(0,0), cex=0.5,
       ilab=format(cbind(PPOVOLUME$mean.pre,
                         PPOVOLUME$sd.pre,
                         PPOVOLUME$mean.post,
                         PPOVOLUME$sd.post,
                         PPOVOLUME$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(FITNESS_REM_PPOVOLUME, row=48.5, cex=0.5, col=11, mlab="")
text(-111.4, 48.5, cex=0.5, bquote(paste("RE Model for 40-89 Minutes/Week: p < 0.002",
                                         "; ", I^2, " = ", .(formatC(FITNESS_REM_PPOVOLUME$I2, digits = 1, format="f")), "%")))

addpoly(CMETAB_REM_PPOVOLUME, row=17.5, cex=0.5, col=11, mlab="")
text(-110, 17.5, cex=0.5, bquote(paste("RE Model for 90-149 Minutes/Week: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(CMETAB_REM_PPOVOLUME$I2, digits = 1, format="f")), "%")))

addpoly(GENPOP_REM_PPOVOLUME, row=0.5, cex=0.5, col=11, mlab="")
text(-112, 0.7, cex=0.5, bquote(paste("RE Model for " >="150 Minutes/Week: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(GENPOP_REM_PPOVOLUME$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 69, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 70, c('Pre', 'Post'), cex=0.38)
text(84, 69.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-83.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOVOLUME$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOVOLUME$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOVOLUME$I2, digits=1, format="f")), "%)"))) 

text(-115, 66.5, cex=0.47, bquote(paste(bolditalic("SCI-Specific Guidelines for Fitness (40-89 Min/Wk)"))))
text(-92.5, 46.5, cex=0.47, bquote(paste(bolditalic("SCI-Specific Guidelines for Cardiometabolic Health (90-149 Min/Wk)"))))
text(-107.5, 15.5, cex=0.47, bquote(paste(bolditalic("Achieving General Population Guidelines (">= "150 Min/Wk)"))))

text(-113, -2, cex=0.48, "Test for Subgroup Differences: Q = 3.47, df = 2, p = 0.18")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(PPOVOLUME$subgroup,c("fitness", "cmetab", "genpop"))]
funnelplotdata <- funnel(REM_PPOVOLUME, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("40 - 89 min/wk","90 - 149 min/wk","150 min/wk or more"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))


