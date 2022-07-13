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

AVO2INTPRESC <- read.xlsx("AVO2_INTPRESC.xlsx", sheet = 6)

AVO2INTPRESC$n.post <- as.numeric(AVO2INTPRESC$n.post)
AVO2INTPRESC$mean.post <- as.numeric(AVO2INTPRESC$mean.post)
AVO2INTPRESC$sd.post <- as.numeric(AVO2INTPRESC$sd.post)
AVO2INTPRESC$n.pre <- as.numeric(AVO2INTPRESC$n.pre)
AVO2INTPRESC$mean.pre <- as.numeric(AVO2INTPRESC$mean.pre)
AVO2INTPRESC$sd.pre <- as.numeric(AVO2INTPRESC$sd.pre)

# Naming subgroups
AVO2INTPRESC$subgroup <- as.factor(AVO2INTPRESC$subgroup)

# Effect size calculation
AVO2INTPRESC_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                              m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                              data=AVO2INTPRESC,var.names=c('AVO2INTPRESC_MD','AVO2INTPRESC_VAR'))

# Overall random effects model
REM_AVO2INTPRESC <- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                        slab=study)
REM_AVO2INTPRESC

# Subgroups random effects models
VO2_REM_AVO2INTPRESC<- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='vo2'))
VO2_REM_AVO2INTPRESC

HR_REM_AVO2INTPRESC <- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='hr'))
HR_REM_AVO2INTPRESC

RPE_REM_AVO2INTPRESC <- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='rpe'))
RPE_REM_AVO2INTPRESC

WLOAD_REM_AVO2INTPRESC <- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='wload'))
WLOAD_REM_AVO2INTPRESC

MIXED_REM_AVO2INTPRESC <- rma(AVO2INTPRESC_MD, AVO2INTPRESC_VAR, data=AVO2INTPRESC_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='mixed/cd'))
MIXED_REM_AVO2INTPRESC

# Test for subgroup differences
AVO2SEI <- sqrt(AVO2INTPRESC_ESCALC$AVO2INTPRESC_VAR)
SGDIFF <- metagen(AVO2INTPRESC_MD, AVO2SEI, studlab=AVO2INTPRESC$study, data=AVO2INTPRESC_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_AVO2INTPRESC, ylim = c(-1,91), xlim = c(-11.5,7), alim = c(-2,2), rows = c(86:79, 74:59, 54:46, 41:33, 28:2), cex = 0.4,
       ilab=format(cbind(AVO2INTPRESC$mean.pre, AVO2INTPRESC$sd.pre, AVO2INTPRESC$mean.post, AVO2INTPRESC$sd.post, AVO2INTPRESC$n.post), digits = 2),
       ilab.xpos = c(-7, -6, -5, -4, -3),
       showweights = TRUE, header = TRUE, slab = study, xlab = 'Mean Difference (L/min)', col=10)

addpoly(VO2_REM_AVO2INTPRESC, row=77.5, cex=0.4, col=11, mlab="")
text(-7.7, 77.5, cex=0.4, bquote(paste("RE Model for Oxygen Consumption: p = 0.003",
                                       "; ", I^2, " = ", .(formatC(VO2_REM_AVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(HR_REM_AVO2INTPRESC, row=57.5, cex=0.4, col=11, mlab="")
text(-8.33, 57.5, cex=0.4, bquote(paste("RE Model for Heart Rate: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(HR_REM_AVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(RPE_REM_AVO2INTPRESC, row=44.5, cex=0.4, col=11, mlab="")
text(-8.8, 44.5, cex=0.4, bquote(paste("RE Model for RPE: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(RPE_REM_AVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(WLOAD_REM_AVO2INTPRESC, row=31.5, cex=0.4, col=11, mlab="")
text(-8.48, 31.5, cex=0.4, bquote(paste("RE Model for Workload: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(WLOAD_REM_AVO2INTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_AVO2INTPRESC, row=0.5, cex=0.4, col=11, mlab="")
text(-7.5, 0.5, cex=0.4, bquote(paste("RE Model for Mixed/Cannot Determine: p < 0.002",
                                       "; ", I^2, " = ", .(formatC(MIXED_REM_AVO2INTPRESC$I2, digits = 1, format="f")), "%")))

text(c(-7, -6, -5, -4, -3), 90, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-6.5, -4.5), 91, c('Pre', 'Post'), cex=0.3)
text(3.9, 90, cex=0.4, bquote(paste(bold("Weight"))))

text(-6.5, -0.9, cex=0.4, bquote(paste("for All Studies (p < 0.001", 
                                       "; ",Tau^2, " = ",.(formatC(REM_AVO2INTPRESC$tau2, digits = 2, format="f")),
                                       "; ",Z, " = ",.(formatC(REM_AVO2INTPRESC$zval, digits = 2, format="f")),
                                       "; ",I^2, " = ",.(formatC(REM_AVO2INTPRESC$I2, digits=1, format="f")), "%)"))) 

text(-9.75, 87.5, cex=0.44, bquote(paste(bolditalic("Oxygen Consumption"))))
text(-10.45, 75.5, cex=0.44, bquote(paste(bolditalic("Heart Rate"))))
text(-9.3, 55.5, cex=0.44, bquote(paste(bolditalic("Rating of Percieved Exertion"))))
text(-10.55, 42.5, cex=0.44, bquote(paste(bolditalic("Workload"))))
text(-9.52, 29.5, cex=0.44, bquote(paste(bolditalic("Mixed/Cannot Determine"))))

text(-7.75, -2, cex=0.4, "Test for Subgroup Differences: Q = 2.08, df = 4, p = 0.72")

# Egger's test
regtest(REM_AVO2INTPRESC, model = "rma", ret.fit = FALSE, digits = 2)

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green")[match(AVO2INTPRESC$subgroup,c("vo2", "hr", "rpe", "wload", "mixed/cd"))]
funnelplotdata <- funnel(REM_AVO2INTPRESC, ylim = c(0,2), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (L/min)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Oxygen consumption","Heart rate","RPE","Workload","Mixed/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green"))
