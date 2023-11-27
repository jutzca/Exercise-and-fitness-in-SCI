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

PPOCPET <- read.xlsx("CPET_DATA.xlsx", sheet = 3)

PPOCPET$n.post <- as.numeric(PPOCPET$n.post)
PPOCPET$mean.post <- as.numeric(PPOCPET$mean.post)
PPOCPET$sd.post <- as.numeric(PPOCPET$sd.post)
PPOCPET$n.pre <- as.numeric(PPOCPET$n.pre)
PPOCPET$mean.pre <- as.numeric(PPOCPET$mean.pre)
PPOCPET$sd.pre <- as.numeric(PPOCPET$sd.pre)

# Naming subgroups
PPOCPET$subgroup <- as.factor(PPOCPET$subgroup)

# Effect size calculation
PPOCPET_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=PPOCPET,var.names=c('PPOCPET_MD','PPOCPET_VAR'))

# Overall random effects model
REM_PPOCPET <- rma(PPOCPET_MD, PPOCPET_VAR, data=PPOCPET_ESCALC, digits=3, 
                    slab=study)
REM_PPOCPET

# Subgroups random effects models
MATCH_REM_PPOCPET <- rma(PPOCPET_MD, PPOCPET_VAR, data=PPOCPET_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='match'))
MATCH_REM_PPOCPET

UNMATCH_REM_PPOCPET <- rma(PPOCPET_MD, PPOCPET_VAR, data=PPOCPET_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='unmatch'))
UNMATCH_REM_PPOCPET

# Test for subgroup differences
PPOSEI <- sqrt(PPOCPET_ESCALC$PPOCPET_VAR)
SGDIFF <- metagen(PPOCPET_MD, PPOSEI, studlab=PPOCPET$study, data=PPOCPET_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOCPET, ylim = c(-1, 75), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(70:23, 18:2), digits = c(0,0),cex=0.4,
       ilab=format(cbind(PPOCPET$mean.pre,
                         PPOCPET$sd.pre,
                         PPOCPET$mean.post,
                         PPOCPET$sd.post,
                         PPOCPET$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(MATCH_REM_PPOCPET, row=21.5, cex=0.4, col=11, mlab="")
text(-108, 21.5, cex=0.38, bquote(paste("RE Model for Matched CPET Modality: p < 0.001",
                                       "; ", I^2, " = ", .(formatC(MATCH_REM_PPOCPET$I2, digits = 1, format="f")), "%")))

addpoly(UNMATCH_REM_PPOCPET, row=0.5, cex=0.4, col=11, mlab="")
text(-108.5, 0.7, cex=0.38, bquote(paste("RE Model for Different CPET Modality: p < 0.001",
                                         "; ", I^2, " = ", .(formatC(UNMATCH_REM_PPOCPET$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 74, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.3)
text(c(-88.5, -66.5), 75, c('Pre', 'Post'), cex=0.3)
text(87.5, 74, cex=0.4, bquote(paste(bold("Weight"))))

text(-87.9, -0.87, cex=0.38, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOCPET$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOCPET$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOCPET$I2, digits=1, format="f")), "%)"))) 

text(-121, 71.5, cex=0.4, bquote(paste(bolditalic("CPET Modality Matches Intervention Modality"))))
text(-120, 19.5, cex=0.4, bquote(paste(bolditalic("CPET Modality Differs to Intervention Modality"))))

text(-112.8, -2, cex=0.38, "Test for Subgroup Differences: Q = 0.61, df = 1, p = 0.44")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange","red")[match(PPOCLASS$subgroup, c("tetraplegia", "paraplegia", "mixed", "not reported/cannot determine"))]
funnelplotdata <- funnel(REM_PPOCLASS, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Tetraplegia","Paraplegia","Mixed","NR/CD"),cex=0.8, pch=20, col=c("purple","blue","orange","red"))
