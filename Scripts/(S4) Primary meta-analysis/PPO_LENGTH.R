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

PPOLENGTH <- read.xlsx("PPO_DATA.xlsx", sheet = 7)

PPOLENGTH$n.post <- as.numeric(PPOLENGTH$n.post)
PPOLENGTH$mean.post <- as.numeric(PPOLENGTH$mean.post)
PPOLENGTH$sd.post <- as.numeric(PPOLENGTH$sd.post)
PPOLENGTH$n.pre <- as.numeric(PPOLENGTH$n.pre)
PPOLENGTH$mean.pre <- as.numeric(PPOLENGTH$mean.pre)
PPOLENGTH$sd.pre <- as.numeric(PPOLENGTH$sd.pre)

# Naming subgroups
PPOLENGTH$subgroup <- as.factor(PPOLENGTH$subgroup)

# Effect size calculation
PPOLENGTH_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                            m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                            data=PPOLENGTH,var.names=c('PPOLENGTH_MD','PPOLENGTH_VAR'))

# Overall random effects model
REM_PPOLENGTH <- rma(PPOLENGTH_MD, PPOLENGTH_VAR, data=PPOLENGTH_ESCALC, digits=3, 
                      slab=study)
REM_PPOLENGTH

# Subgroups random effects models
SIX_REM_PPOLENGTH <- rma(PPOLENGTH_MD, PPOLENGTH_VAR, data=PPOLENGTH_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='6weeks'))
SIX_REM_PPOLENGTH

SIXTWELVE_REM_PPOLENGTH <- rma(PPOLENGTH_MD, PPOLENGTH_VAR, data=PPOLENGTH_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='6-12weeks'))
SIXTWELVE_REM_PPOLENGTH

TWELVE_REM_PPOLENGTH <- rma(PPOLENGTH_MD, PPOLENGTH_VAR, data=PPOLENGTH_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='12weeks'))
TWELVE_REM_PPOLENGTH

# Test for subgroup differences
PPOSEI <- sqrt(PPOLENGTH_ESCALC$PPOLENGTH_VAR)
SGDIFF <- metagen(PPOLENGTH_MD, PPOSEI, studlab=PPOLENGTH$study, data=PPOLENGTH_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOLENGTH, ylim = c(-1, 75), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(70:54, 49:28, 23:2), digits = c(0,0),
       ilab=format(cbind(PPOLENGTH$mean.pre,
                         PPOLENGTH$sd.pre,
                         PPOLENGTH$mean.post,
                         PPOLENGTH$sd.post,
                         PPOLENGTH$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(SIX_REM_PPOLENGTH, row=52.5, cex=0.5, col=11, mlab="")
text(-117, 52.5, cex=0.5, bquote(paste("RE Model for "<="6 weeks",
                                       ": p < 0.001; ", I^2, " = ", .(formatC(SIX_REM_PPOLENGTH$I2, digits = 1, format="f")), "%")))

addpoly(SIXTWELVE_REM_PPOLENGTH, row=26.5, cex=0.5, col=11, mlab="")
text(-107.9, 26.5, cex=0.5, bquote(paste("RE Model for > 6 to "<="12 weeks",
                                         ": p < 0.001; ", I^2, " = ", .(formatC(SIXTWELVE_REM_PPOLENGTH$I2, digits = 1, format="f")), "%")))

addpoly(TWELVE_REM_PPOLENGTH, row=0.5, cex=0.5, col=11, mlab="")
text(-115.4, 0.7, cex=0.5, bquote(paste("RE Model for > 12 weeks: p < 0.001",
                                        "; ", I^2, " = ", .(formatC(TWELVE_REM_PPOLENGTH$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 74, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 75, c('Pre', 'Post'), cex=0.38)
text(84, 74.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOLENGTH$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOLENGTH$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOLENGTH$I2, digits=1, format="f")), "%)"))) 

text(-161, 71.5, cex=0.53, bquote(paste(bolditalic(" "<= "6 Weeks"))))
text(-150, 50.5, cex=0.53, bquote(paste(bolditalic("> 6 to "<="12 Weeks"))))
text(-158.4, 24.5, cex=0.53, bquote(paste(bolditalic("> 12 Weeks"))))

text(-104.5, -2, cex=0.48, "Test for Subgroup Differences: Q = 1.43, df = 2, p = 0.49")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(PPOLENGTH$subgroup,c("6weeks", "6-12weeks", "12weeks"))]
funnelplotdata <- funnel(REM_PPOLENGTH, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("6 weeks or less","7 to 12 weeks","More than 12 weeks"), 
       cex=0.5, pch=20, col=c("purple","blue","orange"))

