install.packages("tidyverse")
install.packages("meta")
install.packages("metafor")
install.packages("openxlsx")

library(tidyverse)
library(meta)
library(metafor)
library(openxlsx)

data_dir <- '~/Documents/R meta-analysis/Secondary meta-analyses/'
wdir <- '~/Documents/R meta-analysis/CPET_DATA/'
setwd(data_dir)

AVO2CPET <- read.xlsx("CPET_DATA.xlsx", sheet = 1)

AVO2CPET$n.post <- as.numeric(AVO2CPET$n.post)
AVO2CPET$mean.post <- as.numeric(AVO2CPET$mean.post)
AVO2CPET$sd.post <- as.numeric(AVO2CPET$sd.post)
AVO2CPET$n.pre <- as.numeric(AVO2CPET$n.pre)
AVO2CPET$mean.pre <- as.numeric(AVO2CPET$mean.pre)
AVO2CPET$sd.pre <- as.numeric(AVO2CPET$sd.pre)

# Naming subgroups
AVO2CPET$subgroup <- as.factor(AVO2CPET$subgroup)

# Effect size calculation
AVO2CPET_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=AVO2CPET,var.names=c('AVO2CPET_MD','AVO2CPET_VAR'))

# Overall random effects model
REM_AVO2CPET <- rma(AVO2CPET_MD, AVO2CPET_VAR, data=AVO2CPET_ESCALC, digits=3, 
                   slab=study)
REM_AVO2CPET

# Subgroups random effects models
MATCH_REM_AVO2CPET <- rma(AVO2CPET_MD, AVO2CPET_VAR, data=AVO2CPET_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='match'))
MATCH_REM_AVO2CPET

UNMATCH_REM_AVO2CPET <- rma(AVO2CPET_MD, AVO2CPET_VAR, data=AVO2CPET_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_AVO2CPET

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2CPET_ESCALC$AVO2CPET_VAR)
SGDIFF <- metagen(AVO2CPET_MD, AVO2SEI, studlab=AVO2CPET$study, data=AVO2CPET_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Save as 475 x 825
## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2CPET, ylim = c(-1,84), xlim = c(-11.5,7), alim = c(-2,2), rows = c(79:24, 19:2), cex=0.4,
       ilab=format(cbind(AVO2CPET$mean.pre, AVO2CPET$sd.pre, AVO2CPET$mean.post, AVO2CPET$sd.post, AVO2CPET$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(MATCH_REM_AVO2CPET, row=22.5, cex=0.4, col=11, mlab="")
text(-7.9, 22.5, cex=0.4, bquote(paste("RE Model for matched CPET: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(MATCH_REM_AVO2CPET$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_AVO2CPET, row=0.5, cex=0.4, col=11, mlab="")
text(-7.9, 0.7, cex=0.4, bquote(paste("RE Model for different CPET: p = 0.005",
                                       "; ", I^2, " = ", .(formatC(UNMATCH_REM_AVO2CPET$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 83, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 84, c('Pre', 'Post'), cex=0.3)
text(3.8, 83, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.2, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2CPET$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2CPET$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2CPET$I2, digits=1, format="f")), "%)"))) 

text(-8.05, 80.5, cex=0.44, bquote(paste(bolditalic("CPET Modality Matches Intervention Modality"))))
text(-8, 20.5, cex=0.44, bquote(paste(bolditalic("CPET Modality Differs to Intervention Modality"))))

text(-7.55, -2, cex=0.4, "Test for Subgroup Differences: Q = 3.87, df = 1, p = 0.049")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(AVO2TSI$subgroup, c("acute (< 1 year)", "chronic (> 1 year)", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_AVO2TSI, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
