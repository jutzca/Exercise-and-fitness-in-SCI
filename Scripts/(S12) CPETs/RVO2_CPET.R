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

RVO2CPET <- read.xlsx("CPET_DATA.xlsx", sheet = 2)

RVO2CPET$n.post <- as.integer(RVO2CPET$n.post)
RVO2CPET$mean.post <- as.numeric(RVO2CPET$mean.post)
RVO2CPET$sd.post <- as.numeric(RVO2CPET$sd.post)
RVO2CPET$n.pre <- as.integer(RVO2CPET$n.pre)
RVO2CPET$mean.pre <- as.numeric(RVO2CPET$mean.pre)
RVO2CPET$sd.pre <- as.numeric(RVO2CPET$sd.pre)

# Naming subgroups
RVO2CPET$subgroup <- as.factor(RVO2CPET$subgroup)

# Effect size calculation
RVO2CPET_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                         m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                         data=RVO2CPET,var.names=c('RVO2CPET_MD','RVO2CPET_VAR'))

# Overall random effects model
REM_RVO2CPET <- rma(RVO2CPET_MD, RVO2CPET_VAR, data=RVO2CPET_ESCALC, digits=3, 
                   slab=study)
REM_RVO2CPET

# Subgroups random effects models
MATCH_REM_RVO2CPET <- rma(RVO2CPET_MD, RVO2CPET_VAR, data=RVO2CPET_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='match'))
MATCH_REM_RVO2CPET

UNMATCH_REM_RVO2CPET <- rma(RVO2CPET_MD, RVO2CPET_VAR, data=RVO2CPET_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_RVO2CPET

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2CPET_ESCALC$RVO2CPET_VAR)
SGDIFF <- metagen(RVO2CPET_MD, RVO2SEI, data=RVO2CPET_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2CPET, ylim = c(-1, 84), xlim = c(-80, 60), at = c(-10, 0, 10, 20), rows = c(79:25, 20:2), digits = 1, cex=0.4,
       ilab=format(cbind(RVO2CPET$mean.pre,
                         RVO2CPET$sd.pre,
                         RVO2CPET$mean.post,
                         RVO2CPET$sd.post,
                         RVO2CPET$n.post), digits = 1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(MATCH_REM_RVO2CPET, row=23.5, cex=0.4, col=11, mlab="")
text(-50, 23.5, cex=0.38, bquote(paste("RE Model for Matched CPET Modality: p < 0.001",
                                         "; ", I^2, " = ", .(formatC(MATCH_REM_RVO2CPET$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_RVO2CPET, row=0.5, cex=0.4, col=11, mlab="")
text(-50, 0.7, cex=0.38, bquote(paste("RE Model for Different CPET Modality: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(UNMATCH_REM_RVO2CPET$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 83, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 84, c('Pre', 'Post'), cex=0.3)
text(38.5, 83, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-42, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2CPET$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2CPET$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2CPET$I2, digits=1, format="f")), "%)"))) 

text(-55.5, 80.5, cex=0.4, bquote(paste(bolditalic("CPET Modality Matches Intervention Modality"))))
text(-55, 21.5, cex=0.4, bquote(paste(bolditalic("CPET Modality Differs to Intervention Modality"))))

text(-52, -2, cex=0.38, "Test for Subgroup Differences: Q = 5.58, df = 1, p = 0.02")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(RVO2TSI$subgroup, c("acute (<1-year)", "chronic (>1-year)", "mixed", "not reported/cannot determine"))] 
funnelplotdata <- funnel(REM_RVO2TSI, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Acute","Chronic","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
