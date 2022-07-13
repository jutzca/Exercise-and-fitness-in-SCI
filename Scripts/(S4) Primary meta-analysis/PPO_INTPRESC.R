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

PPOINTPRESC <- read.xlsx("PPO_DATA.xlsx", sheet = 6)

PPOINTPRESC$n.post <- as.numeric(PPOINTPRESC$n.post)
PPOINTPRESC$mean.post <- as.numeric(PPOINTPRESC$mean.post)
PPOINTPRESC$sd.post <- as.numeric(PPOINTPRESC$sd.post)
PPOINTPRESC$n.pre <- as.numeric(PPOINTPRESC$n.pre)
PPOINTPRESC$mean.pre <- as.numeric(PPOINTPRESC$mean.pre)
PPOINTPRESC$sd.pre <- as.numeric(PPOINTPRESC$sd.pre)

# Naming subgroups
PPOINTPRESC$subgroup <- as.factor(PPOINTPRESC$subgroup)

# Effect size calculation
PPOINTPRESC_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                              m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                              data=PPOINTPRESC,var.names=c('PPOINTPRESC_MD','PPOINTPRESC_VAR'))

# Overall random effects model
REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                        slab=study)
REM_PPOINTPRESC

# Subgroups random effects models
VO2_REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                            slab=study, subset=(subgroup=='vo2'))
VO2_REM_PPOINTPRESC

HR_REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='hr'))
HR_REM_PPOINTPRESC

RPE_REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                           slab=study, subset=(subgroup=='rpe'))
RPE_REM_PPOINTPRESC

WLOAD_REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='wload'))
WLOAD_REM_PPOINTPRESC

MIXED_REM_PPOINTPRESC <- rma(PPOINTPRESC_MD, PPOINTPRESC_VAR, data=PPOINTPRESC_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='mixed/cd'))
MIXED_REM_PPOINTPRESC

# Test for subgroup differences
PPOSEI <- sqrt(PPOINTPRESC_ESCALC$PPOINTPRESC_VAR)
SGDIFF <- metagen(PPOINTPRESC_MD, PPOSEI, studlab=PPOINTPRESC$study, data=PPOINTPRESC_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOINTPRESC, ylim = c(-1, 83), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(78:70, 65:52, 47:42, 37:31, 26:2), digits = c(0,0),
       ilab=format(cbind(PPOINTPRESC$mean.pre,
                         PPOINTPRESC$sd.pre,
                         PPOINTPRESC$mean.post,
                         PPOINTPRESC$sd.post,
                         PPOINTPRESC$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header =TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(VO2_REM_PPOINTPRESC, row=68.5, cex=0.5, col=11, mlab="")
text(-103.5, 68.5, cex=0.5, bquote(paste("RE Model for Oxygen Consumption: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(VO2_REM_PPOINTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(HR_REM_PPOINTPRESC, row=50.5, cex=0.5, col=11, mlab="")
text(-116, 50.5, cex=0.5, bquote(paste("RE Model for Heart Rate: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(HR_REM_PPOINTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(RPE_REM_PPOINTPRESC, row=40.5, cex=0.5, col=11, mlab="")
text(-127, 40.5, cex=0.5, bquote(paste("RE Model for RPE: p = 0.03",
                                         "; ", I^2, " = ", .(formatC(RPE_REM_PPOINTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(WLOAD_REM_PPOINTPRESC, row=29.5, cex=0.5, col=11, mlab="")
text(-119.2, 29.5, cex=0.5, bquote(paste("RE Model for Workload: p = 0.01",
                                         "; ", I^2, " = ", .(formatC(WLOAD_REM_PPOINTPRESC$I2, digits = 1, format="f")), "%")))

addpoly(MIXED_REM_PPOINTPRESC, row=0.5, cex=0.5, col=11, mlab="")
text(-99.7, 0.7, cex=0.5, bquote(paste("RE Model for Mixed/Cannot Determine: p < 0.003",
                                         "; ", I^2, " = ", .(formatC(MIXED_REM_PPOINTPRESC$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 82, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 83, c('Pre', 'Post'), cex=0.38)
text(84, 82.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOINTPRESC$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOINTPRESC$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOINTPRESC$I2, digits=1, format="f")), "%)"))) 

text(-145.5, 79.5, cex=0.5, bquote(paste(bolditalic("Oxygen Consumption"))))
text(-160, 66.5, cex=0.5, bquote(paste(bolditalic("Heart Rate"))))
text(-136.5, 48.5, cex=0.5, bquote(paste(bolditalic("Rating of Percieved Exertion"))))
text(-161.6, 38.5, cex=0.5, bquote(paste(bolditalic("Workload"))))
text(-141.6, 27.5, cex=0.5, bquote(paste(bolditalic("Mixed/Cannot Determine"))))

text(-102, -2, cex=0.48, "Test for Subgroup Differences: Q = 20.01, df = 4, p < 0.001")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red","green")[match(PPOINTPRESC$subgroup,c("vo2", "hr", "rpe", "wload", "mixed/cd"))]
funnelplotdata <- funnel(REM_PPOINTPRESC, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Oxygen consumption","Heart rate","RPE","Workload","Mixed/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red","green"))