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

PPOFREQ <- read.xlsx("PPO_DATA_mixedout.xlsx", sheet = 8)

PPOFREQ$n.post <- as.numeric(PPOFREQ$n.post)
PPOFREQ$mean.post <- as.numeric(PPOFREQ$mean.post)
PPOFREQ$sd.post <- as.numeric(PPOFREQ$sd.post)
PPOFREQ$n.pre <- as.numeric(PPOFREQ$n.pre)
PPOFREQ$mean.pre <- as.numeric(PPOFREQ$mean.pre)
PPOFREQ$sd.pre <- as.numeric(PPOFREQ$sd.pre)

# Naming subgroups
PPOFREQ$subgroup <- as.factor(PPOFREQ$subgroup)

# Effect size calculation
PPOFREQ_ESCALC <- escalc(measure="MD",m1i=mean.post,sd1i=sd.post,n1i=n.post,
                          m2i=mean.pre,sd2i=sd.pre,n2i=n.pre,
                          data=PPOFREQ,var.names=c('PPOFREQ_MD','PPOFREQ_VAR'))

# Overall random effects model
REM_PPOFREQ <- rma(PPOFREQ_MD, PPOFREQ_VAR, data=PPOFREQ_ESCALC, digits=3, 
                    slab=study)
REM_PPOFREQ

# Subgroups random effects models
THREE_REM_PPOFREQ <- rma(PPOFREQ_MD, PPOFREQ_VAR, data=PPOFREQ_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='three'))
THREE_REM_PPOFREQ

THREEFIVE_REM_PPOFREQ <- rma(PPOFREQ_MD, PPOFREQ_VAR, data=PPOFREQ_ESCALC, digits=3, 
                              slab=study, subset=(subgroup=='threefive'))
THREEFIVE_REM_PPOFREQ

FIVE_REM_PPOFREQ <- rma(PPOFREQ_MD, PPOFREQ_VAR, data=PPOFREQ_ESCALC, digits=3, 
                         slab=study, subset=(subgroup=='five'))
FIVE_REM_PPOFREQ

NR.CD_REM_PPOFREQ <- rma(PPOFREQ_MD, PPOFREQ_VAR, data=PPOFREQ_ESCALC, digits=3, 
                          slab=study, subset=(subgroup=='not reported/cannot determine'))
NR.CD_REM_PPOFREQ

# Test for subgroup differences
PPOSEI <- sqrt(PPOFREQ_ESCALC$PPOFREQ_VAR)
SGDIFF <- metagen(PPOFREQ_MD, PPOSEI, studlab=PPOFREQ$study, data=PPOFREQ_ESCALC, fixed=FALSE, method.tau="REML", subgroup=subgroup)
SGDIFF

## Forest plot
par(mar=c(4,4,0,2))
forest(REM_PPOFREQ, ylim = c(-1, 72), xlim = c(-180, 135), at = c(-40, -20, 0, 20, 40, 60), rows = c(67:55, 50:11, 6:2), digits = c(0,0), cex=0.5,
       ilab=format(cbind(PPOFREQ$mean.pre,
                         PPOFREQ$sd.pre,
                         PPOFREQ$mean.post,
                         PPOFREQ$sd.post,
                         PPOFREQ$n.post), digits=0),
       ilab.xpos = c(-94, -83, -72, -61, -50),
       showweights = TRUE, header = TRUE,
       slab = study, xlab = 'Mean Difference (W)', col=10)

addpoly(THREE_REM_PPOFREQ, row=53.5, cex=0.5, col=11, mlab="")
text(-107.5, 53.5, cex=0.5, bquote(paste("RE Model for < 3 Sessions/Week: p = 0.003",
                                       "; ", I^2, " = ", .(formatC(THREE_REM_PPOFREQ$I2, digits = 1, format="f")), "%")))

addpoly(THREEFIVE_REM_PPOFREQ, row=9.5, cex=0.5, col=11, mlab="")
text(-98.3, 9.5, cex=0.5, bquote(paste("RE Model for ">="3 to < 5 Sessions/Week: p < 0.004",
                                       "; ", I^2, " = ", .(formatC(THREEFIVE_REM_PPOFREQ$I2, digits = 1, format="f")), "%")))

addpoly(FIVE_REM_PPOFREQ, row=0.5, cex=0.5, col=11, mlab="")
text(-107.5, 0.7, cex=0.5, bquote(paste("RE Model for ">="5 Sessions/Week: p = 0.10",
                                       "; ", I^2, " = ", .(formatC(FIVE_REM_PPOFREQ$I2, digits = 1, format="f")), "%")))

text(c(-94, -83, -72, -61, -48), 71, c('Mean', 'SD', 'Mean', 'SD', 'Total N'), cex=0.38)
text(c(-88.5, -66.5), 72, c('Pre', 'Post'), cex=0.38)
text(84, 71.05, cex=0.5, bquote(paste(bold("Weight"))))

text(-77.5, -0.88, cex=0.48, bquote(paste("for All Studies (p < 0.001", 
                                          "; ",Tau^2, " = ",.(formatC(REM_PPOFREQ$tau2, digits = 2, format="f")),
                                          "; ",Z, " = ",.(formatC(REM_PPOFREQ$zval, digits = 2, format="f")),
                                          "; ",I^2, " = ",.(formatC(REM_PPOFREQ$I2, digits=1, format="f")), "%)"))) 

text(-149.1, 68.5, cex=0.5, bquote(paste(bolditalic("< 3 Sessions/Week"))))
text(-142.3, 51.5, cex=0.5, bquote(paste(bolditalic(" ">="3 to < 5 Sessions/Week"))))
text(-150.5, 7.5, cex=0.5, bquote(paste(bolditalic(" ">="5 Sessions/Week"))))

text(-102, -2, cex=0.48, "Test for Subgroup Differences: Q = 25.06, df = 2, p < 0.001")

# Funnel plot
par(mar=c(4,4,1,2))

my_colours <- c("purple","blue","orange")[match(PPOFREQ$subgroup,c("three", "threefive", "five"))]
funnelplotdata <- funnel(REM_PPOFREQ, ylim = c(0,50), digits = c(0,1), cex = 0.5, xlab = 'Mean Difference (W)')
with(funnelplotdata, points(x,y, col = my_colours, pch = 20))

legend("topright", c("Less than 3 sessions/wk","3 to 4 sessions/wk","5 or more sessions/wk"),
       cex=0.5, pch=20, col=c("purple","blue","orange"))

