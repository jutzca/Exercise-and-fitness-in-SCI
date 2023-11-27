install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/Secondary meta-analyses/'
wdir <- '~/Documents/R meta-analysis/GAIT_DATA/'
setwd(data_dir)

PPOGAIT <- read.xlsx("GAIT_DATA.xlsx", sheet = 3)

PPOGAIT$n.post <- as.numeric(PPOGAIT$n.post)
PPOGAIT$mean.post <- as.numeric(PPOGAIT$mean.post)
PPOGAIT$sd.post <- as.numeric(PPOGAIT$sd.post)
PPOGAIT$n.pre <- as.numeric(PPOGAIT$n.pre)
PPOGAIT$mean.pre <- as.numeric(PPOGAIT$mean.pre)
PPOGAIT$sd.pre <- as.numeric(PPOGAIT$sd.pre)

# Naming subgroups
PPOGAIT$subgroup <- as.factor(PPOGAIT$subgroup)

# Effect size calculation
PPOGAIT_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=PPOGAIT,var.names=c('PPOGAIT_MD','PPOGAIT_VAR'))

# Overall random effects model
REM_PPOGAIT <- rma(PPOGAIT_MD, PPOGAIT_VAR, data=PPOGAIT_ESCALC, digits=3, 
                    slab=study)
REM_PPOGAIT

# Subgroups random effects models
ACE_REM_PPOGAIT <- rma(PPOGAIT_MD, PPOGAIT_VAR, data=PPOGAIT_ESCALC, digits=3, 
                        slab=study, subset=(subgroup=='ace'))
ACE_REM_PPOGAIT

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOGAIT, ylim = c(-2,7), xlim = c(-95,60), at = c(-20, -10, 0, 10, 20), rows = c(3:2), digits = 1, cex=0.4,
       ilab=format(cbind(PPOGAIT$mean.pre, PPOGAIT$sd.pre, PPOGAIT$mean.post, PPOGAIT$sd.post, PPOGAIT$n.post), digits = 2),
       ilab.xpos = c(-54, -48, -42, -36, -30),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(ACE_REM_PPOGAIT, row=0.5, cex=0.4, col=11, mlab="")
text(-66.6, 0.5, cex=0.4, bquote(paste("RE Model for ACE CPET: p = 0.68",
                                       "; ", I^2, " = ", .(formatC(ACE_REM_PPOGAIT$I2, digits = 1, format="f")), "%")))

text(c(-54, -48, -42, -36, -30), 6, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.35)
text(c(-51, -39), 7, c('Pre', 'Post'), cex=0.35)
text(36, 6, cex=0.4, bquote(paste(bold("Weight"))))

text(-49.2, -0.93, cex=0.4, bquote(paste("for All Studies (p = 0.68", 
                                        "; ",Tau^2, " = ",.(formatC(REM_PPOGAIT$tau2, digits = 2, format="f")),
                                        "; ",Z, " = ",.(formatC(REM_PPOGAIT$zval, digits = 2, format="f")),
                                        "; ",I^2, " = ",.(formatC(REM_PPOGAIT$I2, digits=1, format="f")), "%)"))) 

text(-85, 4.25, cex=0.44, bquote(paste(bolditalic("ACE CPET"))))
