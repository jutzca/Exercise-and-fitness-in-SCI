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

RVO2INTPRESC <- read.xlsx("RVO2_DATA.xlsx", sheet = 6)

RVO2INTPRESC$n.post <- as.integer(RVO2INTPRESC$n.post)
RVO2INTPRESC$mean.post <- as.numeric(RVO2INTPRESC$mean.post)
RVO2INTPRESC$sd.post <- as.numeric(RVO2INTPRESC$sd.post)
RVO2INTPRESC$n.pre <- as.integer(RVO2INTPRESC$n.pre)
RVO2INTPRESC$mean.pre <- as.numeric(RVO2INTPRESC$mean.pre)
RVO2INTPRESC$sd.pre <- as.numeric(RVO2INTPRESC$sd.pre)

# Naming subgroups
RVO2INTPRESC$subgroup <- as.factor(RVO2INTPRESC$subgroup)

# Effect size calculation
RVO2INTPRESC_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                           m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                           data=RVO2INTPRESC,var.names=c('RVO2INTPRESC_MD','RVO2INTPRESC_VAR'))

# Overall random effects model
REM_RVO2INTPRESC <- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                     slab=study)
REM_RVO2INTPRESC

# Subgroups random effects models
VO2_REM_RVO2INTPRESC <- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                             slab=study, subset=(subgroup=='vo2'), control=list(stepadj = 0.5, maxiter=10000))
VO2_REM_RVO2INTPRESC

HR_REM_RVO2INTPRESC <- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                                slab=study, subset=(subgroup=='hr'))
HR_REM_RVO2INTPRESC

RPE_REM_RVO2INTPRESC<- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='rpe'))
RPE_REM_RVO2INTPRESC

WLOAD_REM_RVO2INTPRESC <- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='wload'), control=list(stepadj=0.5, maxiter=10000))
WLOAD_REM_RVO2INTPRESC

MIXED_REM_RVO2INTPRESC <- rma(RVO2INTPRESC_MD, RVO2INTPRESC_VAR, data=RVO2INTPRESC_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='mixed/cd'))
MIXED_REM_RVO2INTPRESC

# Test for subgroup differences
RVO2SEI <- sqrt(RVO2INTPRESC_ESCALC$RVO2INTPRESC_VAR)
SGDIFF <- metagen(RVO2INTPRESC_MD, RVO2SEI, data=RVO2INTPRESC_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup, 
                  control=list(stepadj=0.5, maxiter=1000))
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_RVO2INTPRESC, ylim = c(-1, 96), xlim = c(-85, 60), at = c(-10, 0, 10, 20), rows = c(91:80, 75:50, 45:38, 33:28, 23:2), digits = 1,
       ilab=format(cbind(RVO2INTPRESC$mean.pre,
                         RVO2INTPRESC$sd.pre,
                         RVO2INTPRESC$mean.post,
                         RVO2INTPRESC$sd.post,
                         RVO2INTPRESC$n.post), digits=1),
       ilab.xpos = c(-40, -35, -30, -25, -20),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (mL/kg/min)', col = 10)

addpoly(VO2_REM_RVO2INTPRESC, row=78.5, cex=0.4, col=11, mlab="")
text(-55.8, 78.5, cex=0.38, bquote(paste("RE Model for Oxygen Consumption: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(VO2_REM_RVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(HR_REM_RVO2INTPRESC, row=48.5, cex=0.4, col=11, mlab="")
text(-60.1, 48.5, cex=0.38, bquote(paste("RE Model for Heart Rate: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(HR_REM_RVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(RPE_REM_RVO2INTPRESC, row=36.5, cex=0.4, col=11, mlab="")
text(-63, 36.5, cex=0.38, bquote(paste("RE Model for RPE: p = 0.002",
                                         "; ", I^2, " = ", .(formatC(RPE_REM_RVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(WLOAD_REM_RVO2INTPRESC, row=26.5, cex=0.4, col=11, mlab="")
text(-61.4, 26.5, cex=0.38, bquote(paste("RE Model for Workload: p = 0.002",
                                         "; ", I^2, " = ", .(formatC(WLOAD_REM_RVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_RVO2INTPRESC, row=0.5, cex=0.4, col=11, mlab="")
text(-53.6, 0.5, cex=0.38, bquote(paste("RE Model for Mixed/Cannot Determine: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(MIXED_REM_RVO2INTPRESC$I2, digits = 1, format="f")), "%")))

text(c(-40, -35, -30, -25, -20), 95, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-37.5, -27.5), 96, c('Pre', 'Post'), cex=0.3)
text(38.5, 95, cex=0.4, bquote(paste(bold("Weight (%)"))))

text(-45.8, -0.9, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                         "; ",Tau^2, " = ",.(formatC(REM_RVO2INTPRESC$tau2, digits = 2, format="f")),
                                         "; ",Z, " = ",.(formatC(REM_RVO2INTPRESC$zval, digits = 2, format="f")),
                                         "; ",I^2, " = ",.(formatC(REM_RVO2INTPRESC$I2, digits=1, format="f")), "%)"))) 

text(-71.1, 92.5, cex=0.4, bquote(paste(bolditalic("Oxygen Consumption"))))
text(-76.9, 76.5, cex=0.4, bquote(paste(bolditalic("Heart Rate"))))
text(-67.9, 46.5, cex=0.4, bquote(paste(bolditalic("Rating of Percieved Exertion"))))
text(-77.7, 34.5, cex=0.4, bquote(paste(bolditalic("Workload"))))
text(-69.6, 24.5, cex=0.4, bquote(paste(bolditalic("Mixed/Cannot Determine"))))

text(-55.7, -2, cex=0.38, "Test for Subgroup Differences: Q = 2.16, df = 4, p = 0.71")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green")[match(RVO2INTPRESC$subgroup,c("vo2", "hr", "rpe", "wload", "mixed/cd"))]
funnelplotdata <- funnel(REM_RVO2INTPRESC, ylim = c(0,8), digits = c(0,0), cex = 0.5, xlab = 'Mean Difference (mL/kg/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Oxygen consumption","Heart rate","RPE","Workload","Mixed/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green"))



