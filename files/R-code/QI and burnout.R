#This file is best used within R Studio
#---------------------------------
version
citation(package = "base", lib.loc = NULL, auto = NULL)
KUBlue = "#0022B4"
SkyBlue = "#6DC6E7"
#windows(600, 600, pointsize = 12) # Size 600x600
setwd(dirname(rstudioapi::getSourceEditorContext()$path))
#setwd("../plots")
getwd()
par(mar=c(5.1,4.1,4.1,2.1), mfrow=c(1,1))
old.par <- par(no.readonly=T)
plot.new()
xmin <- par("usr")[1] + strwidth("A")
xmax <- par("usr")[2] - strwidth("A")
ymin <- par("usr")[3] + strheight("A")
ymax <- par("usr")[4] - strheight("A")
current.date <- as.character(strftime (Sys.time(), format="%Y-%m-%d", tz="", usetz=FALSE))
#---------------------------------
# Mixing fonts in text
# mtext(expression(paste(bold("Note: "),"trials in green region showed significant benefit.")), side = 1, line = 4.2,adj=0,cex=1)
#---------------------------------
`%notin%` <- Negate(`%in%`)
p.value <- sprintf(p.value, fmt='%#.3f')
#---------------------------------
# Mixing fonts in text
# mtext(expression(paste(bold("Note: "),"trials in green region showed significant benefit.")), side = 1, line = 4.2,adj=0,cex=1)library(stringr) # str_trim, str_remove
library(plyr) # rename columns
library(metafor)
library (meta) # metamean
library(openxlsx) # read.xlsx
library(xlsx)     # read.xlsx2
#library(Rcmdr)
library(boot) #inv.logit
library(moments) # skewness
library(tseries) # jarque.bera.tes
library(grid)
#library(ggplot2)
#library(barplot3d)
#library(rgl) # text3d
#library("rqdatatable") # natural_join (replaces NA values in a dataframe merge)
#---------------------------------
# For troubleshooting
options(error=NULL)
library(tcltk) 
# msgbox for troubleshooting: 
#tk_messageBox(type = "ok", paste(current.state,': ', nrow(data.state.current),sepo=''), caption = "Troubleshooting")
#browser()
# Finish, quit is c or Q
# enter f press enter and then enter Q
#---------------------------------
# Grab data
# co <- read.table("https://data.princeton.edu/eco572/datasets/cohhpop.dat", col.names=c("age","pop"), header=FALSE)
file.filter <- matrix(c("Spreadsheets","*.csv;*.xls;*.xlsx","All","*.*"),byrow=TRUE,ncol=2)
filename     = choose.files(filters = file.filter,caption = "Select data file",index = 1,multi=FALSE)
data.import   <- read.table(filename, header=TRUE, sep=",", na.strings="NA", dec=".", stringsAsFactors=FALSE, strip.white=TRUE)
data.import   <- read.xlsx(filename)
data.temp <- data.import
head(data.temp)
summary(data.temp)

#---------------------------------
# Get external data
file.filter <- matrix(c("Spreadsheets","*.csv;*.xls;*.xlsx","All","*.*"),byrow=TRUE,ncol=2)
filename = choose.files(filters = file.filter,caption = "Select data file",index = 1,multi=FALSE)
data<- read.table(filename, header=TRUE, sep=",", na.strings="NA", dec=".", strip.white=TRUE)
#---------------------------------

head(data)
summary(data)

#label(data[["Local.control.or.benefit"]]) <- "Local control or benefit"
names(data)[12] <- "Local control or benefit"
#data[12] <- as.factor(data[12])

#Conversions:
for(i in 1:nrow(data))
  {
  if (data$measure[i] == "BETA")
  {
    data$effect.size[i] <- 2.71828182845^data$effect.size[i]
    data$ci.l[i]        <- 2.71828182845^data$ci.l[i]
    data$ci.u[i]        <- 2.71828182845^data$ci.u[i]
  }
  if (data$measure[i] == "MD")
  {
    # https://handbook-5-1.cochrane.org/chapter_9/9_2_3_2_the_standardized_mean_difference.htm
    data$effect.size[i] <- data$effect.size[i]/data$sd[i]
    data$ci.l[i]        <- data$ci.l[i]/data$sd[i]
    data$ci.u[i]        <- data$ci.u[i]/data$sd[i]
    data$measure[i] <- "SMD"
  }
  if (data$measure[i] == "SMD")
    {
    # https://handbook-5-1.cochrane.org/chapter_12/12_6_3_re_expressing_smds_by_transformation_to_odds_ratio.htm
    data$effect.size[i] <- exp(3.14159*data$effect.size[i]/sqrt(3))
    data$ci.l[i]        <- exp(3.14159*data$ci.l[i]/sqrt(3))
    data$ci.u[i]        <- exp(3.14159*data$ci.u[i]/sqrt(3))
    }
}

meta1 <- data
meta1 <- meta1[ which(meta1$Study!='VA (Locatelli)*'), ]
meta1$Study <- paste(meta1$Study, ", ",meta1$Year, sep="")
meta1$OR.Ln <- log(meta1$effect.size)
meta1$CI.l.Ln <- log(meta1$ci.l)
meta1$CI.u.Ln <- log(meta1$ci.u)
meta1$CI.Ln_Width <- meta1$CI.u.Ln - meta1$CI.l.Ln
meta1$SE.Ln <- meta1$CI.Ln_Width/2/1.96 
# Package meta appears to already do Chinn's conversion of OR to effect size.
meta1$effect.size <- meta1$OR.Ln#/1.81
meta1$effect.size.SE <- meta1$SE.Ln#/1.81

#-------------------------
Title <- "Do QI projects and practice redesign affect burnout?"
# Meta-analysis from library meta
results <- metagen(meta1$effect.size, meta1$effect.size.SE, sm="OR", data=meta1, comb.fixed=FALSE, backtransf = TRUE, studlab=meta1$Study, byvar = meta1$`Local control or benefit`)

#par(mar=c(12.1,4.1,4.1,2.1))
forest(results,print.p=FALSE, leftcols=c("studlab"), just.addcols.left = "left", colgap.left = "5mm", print.tau2=FALSE,col.diamond="blue", col.diamond.lines="blue", print.I2.ci=TRUE,overall=TRUE,
    xlab = "Odds ratio (estimated)",smlab ="Burnout", 
    overall.hetstat = TRUE, test.subgroup.random = TRUE,
    label.test.subgroup.random = "Subgroup diff: ", test.effect.subgroup.random = FALSE,
    plotwidth = "10cm"
    #label.left = "Favors QI", label.right="Favors no QI"
    )
grid.text(Title, 0.5, 0.97, gp = gpar(fontsize = 14, fontface = "bold"))

#grid.text("Notes:",  0.02, 0.1, just = c("left", "bottom"), gp = gpar(fontsize = 12, fontface = "bold"))
#grid.text("* Comparison is personnel reporting voice in practice changes versus those not reporting voice.", 0.02, 0.05, just = c("left", "bottom"), gp = gpar(fontsize = 12, fontface = "plain"))

#---------------------------------
# Printing
Title <- gsub("\\?|\\!", "", Title)
plotname = Title
rstudioapi::savePlotAsImage(
  paste(plotname," - ",current.date,".png",sep=""),
  format = "png", width = 1000, height = 800)
