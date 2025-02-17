# Analyse results from Results_different_methods.xlsx

library(dplyr, quietly = TRUE)
library(tidyverse)
library("ggpubr")
library(scales)
library(readxl)
library(car) # Test Homogeneity of variances
library(factoextra) # Plot PCA 
library(multcompView)
library(RColorBrewer) #brewer.pal()
library(openxlsx)

# Loading ####
rm(list=ls()) # cleaning console
graphics.off() # cleaning plots
'%!in%' <- function(x,y)!('%in%'(x,y))
setwd("W:/ESG/DOW_SLM/Data_archive/IR_Data/Method_validation.tests/4. Results analysis")
wd.out="W:/ESG/DOW_SLM/Data_archive/IR_Data/Method_validation.tests/4. Results analysis"
# Load compiled excel
DATA=read_excel("Results_different_methods.xlsx","Data_R") 

# Julia export Data

write.xlsx(DATA, "W:/ESG/DOW_SLM/Data_archive/IR_Data/Method_validation.tests/4. Results analysis/DATA_Master_010724.xlsx")


DATA$Type.Method.Polymer.Soil= paste(DATA$Type, DATA$Method , DATA$Polymer , DATA$Soil )
DATA$Method.Polymer.Soil= paste(DATA$Method , DATA$Polymer , DATA$Soil )
DATA$Polymer.Soil= paste(DATA$Polymer , DATA$Soil )
DATA$Method.Soil= paste(DATA$Method , DATA$Soil )

Title.Plot.Y=list(Num.measure= "Num. measured",
                  Area.measure.px= "Area [px]", 
                  Area.measure.um=  expression( "Area [?m"^2 *"]"),
                  Num.cor.spiked= "Num.cor.spiked",
                  Area.cor.spiked.um= expression( "Area.cor.spiked [?m"^2 *"]"), 
                  Recovery.num= "Recovery.num [%]",
                  Recovery.area= "Recovery.area [%]",
                  Num.mean.soil.method="Num.mean.soil.method",
                  Num.spiked="Num.spiked",
                  Area.um.mean.soil.method="Area.mean.soil.method",
                  Area.spiked.um="Area.spiked.um"
                  )

Soils= c("ES",  "NL",  "CH") # unique(DATA$Soil)
Polymers= unique(DATA$Polymer)
Polymers.spike=c( "PE" ,    "PLA" ,   "PVC", "Total")
Methods= unique(DATA$Method)

# Spike calculation ####

  # * Summary of environmental, non spiked soil####
    # Results per method, polymer and soil 
  
    DATA_Summary.n= subset(DATA, Type=="environmental" & Spiked == "n")%>% #
      dplyr::group_by(Method, Polymer, Soil ) %>%
      dplyr::summarise( Num.analysis=n(),
                        Num.measure.mean=mean(Num.measure),
                          Area.measure.um.mean=mean(Area.measure.um) )
    # Export: write.csv( DATA_Summary.n, paste(wd.out,"DATA_Summary.n_2024.02.20.csv", sep = "/"))

  # * Recovery correction with background level
    DATA$Num.mean.soil.method=0 # Average number of uP found for the polymer (p), method (m), soil (s)    
    DATA$Area.um.mean.soil.method=0# Average area of uP found for the polymer (p), method (m), soil (s)        
    DATA$Num.cor.spiked=0       # Correct the Measured Num with the Average number of uP found for the polymer p, method m, soil s 
    DATA$Area.cor.spiked.um=0   # Correct the Measured Area with the Average number of uP found for the polymer p, method m, soil s
   
    
    # Calculation per polymer (p), method (m), soil (s) 
     for(p in seq_along(Polymers.spike)) {
       for (m in seq_along(Methods)){
         for(s in seq_along(Soils)){ 
        
           # Average number of uP found for the polymer (p), method (m), soil (s) 
           DATA$Num.mean.soil.method [ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m]] = 
             DATA_Summary.n$Num.measure.mean[DATA_Summary.n$Polymer== Polymers.spike[p] & DATA_Summary.n$Soil==Soils[s] & DATA_Summary.n$Method==Methods[m] ]
           
           # Average number of uP found for the polymer (p), method (m), soil (s) 
           DATA$Area.um.mean.soil.method [ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m]] = 
             DATA_Summary.n$Area.measure.um.mean[DATA_Summary.n$Polymer== Polymers.spike[p] & DATA_Summary.n$Soil==Soils[s] & DATA_Summary.n$Method==Methods[m] ]
           
           # Correct the Measured Num with the Average number of uP found for the polymer p, method m, soil s 
           DATA$Num.cor.spiked[ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m]] =
                 DATA$Num.measure[ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m]] -
           DATA_Summary.n[DATA_Summary.n$Polymer== Polymers.spike[p] & DATA_Summary.n$Soil==Soils[s] & DATA_Summary.n$Method==Methods[m] , "Num.measure.mean"][[1]]
         
           # Correct the Measured Area with the Average number of uP found for the polymer p, method m, soil s 
         DATA$Area.cor.spiked.um[ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m]] =
           DATA$Area.measure.um[ DATA$Polymer== Polymers.spike[p] & DATA$Soil==Soils[s] & DATA$Method==Methods[m] ] -
           DATA_Summary.n[DATA_Summary.n$Polymer== Polymers.spike[p] & DATA_Summary.n$Soil==Soils[s] & DATA_Summary.n$Method==Methods[m] ,"Area.measure.um.mean"][[1]]
         }
       }
     }
         
# Recovery ####   
DATA$Recovery.num=   DATA$Num.cor.spiked/DATA$Num.spiked *100
DATA$Recovery.area=   DATA$Area.cor.spiked.um/DATA$Area.spiked.um *100

# Aggregation and visualization for specific factor.X*Y ----------------    

# Choose factor of interest (per parcel, per farm, per layer, per Management 
str( Title.Plot.Y)

#1. DATA subset ################

# * Environmental samples, non spiked 
  DATA.plot=subset(DATA, Type== "environmental" & Type == "n" & Polymer=="Total")
  Plot_title="Environmental samples, Sum Polymers"
  
  DATA.plot=subset(DATA, Type== "environmental" & Spiked == "n" & Polymer=="PVC")
  Plot_title="Environmental samples, PVC"
  
  # ** Per method
    DATA.plot=subset(DATA, Type== "environmental" & Type == "n" & Polymer%in%c("Others","PE","PLA","PVC") & Method=="LDIR")
    Plot_title="Environmental samples, LDIR, All Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Type == "n" & Polymer=="Total" & Method=="Microscope" )
    Plot_title="Environmental samples, Microscope, Sum Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Type == "n" & Polymer%in%c("Others","PE","PLA","PVC") & Method=="uFTIR 4x"  )
    Plot_title="Environmental samples, uFTIR 4x, All Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Type == "n" & Polymer%in%c("Others","PE","PLA","PVC") & Method=="uFTIR 15x"  )
    Plot_title="Environmental samples, uFTIR 15x, All Polymers"

# * Environmental samples, spiked 
  DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer=="Total")
  Plot_title="Spiked samples, Sum Polymers"

  DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer=="PE")
  Plot_title="Spiked samples, PE"

  # ** Per method
    DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in% c("Others","PE","PLA","PVC") & Method=="LDIR")
    Plot_title="Spiked  samples, LDIR, All Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer=="Total" & Method=="Microscope" )
    Plot_title="Spiked  samples, Microscope, Sum Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in%c("Others","PE","PLA","PVC") & Method=="uFTIR 4x"  )
    Plot_title="Spiked  samples, uFTIR 4x, All Polymers"
    
    DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in%c("Others","PE","PLA","PVC") & Method=="uFTIR 15x"  )
    Plot_title="Spiked samples, uFTIR 15x, All Polymers"


# * Environmental samples, spiked, only PE, PLA, PVC
DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in%c("PE","PLA","PVC") & Method=="LDIR")
Plot_title="Spiked  samples, LDIR"

DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer=="Total" & Method=="Microscope" )
Plot_title="Spiked  samples, Microscope"

DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in%c("PE","PLA","PVC") & Method=="uFTIR 4x"  )
Plot_title="Spiked  samples, uFTIR 4x"

DATA.plot=subset(DATA, Type== "environmental" & Spiked == "spiked" & Polymer %in%c("PE","PLA","PVC") & Method=="uFTIR 15x"  )
Plot_title="Spiked samples, uFTIR 15x"


# 2. Y ###########
Ytxt="Num.measure"
DATA.plot$Y=DATA.plot$Num.measure

Ytxt="Area.measure.um"
DATA.plot$Y=DATA.plot$Area.measure.um

Ytxt="Num.mean.soil.method"
DATA.plot$Y=DATA.plot$Num.mean.soil.method

Ytxt="Area.um.mean.soil.method"
DATA.plot$Y=DATA.plot$Area.um.mean.soil.method

Ytxt="Num.cor.spiked" # = Mesured - Background
DATA.plot$Y=DATA.plot$Num.cor.spiked

Ytxt="Area.cor.spiked.um" # = Mesured - Background
DATA.plot$Y=DATA.plot$Area.cor.spiked.um

Ytxt="Num.spiked" # 
DATA.plot$Y=DATA.plot$Num.spiked

Ytxt="Area.spiked.um" # 
DATA.plot$Y=DATA.plot$Area.spiked.um

Ytxt="Recovery.area"
DATA.plot$Y=DATA.plot$Recovery.area

Ytxt="Recovery.num"
DATA.plot$Y=DATA.plot$Recovery.num

# 3. X ###########
Xtxt="Type.Method.Polymer.Soil"
DATA.plot$X=factor(DATA.plot$Method.Polymer.Soil)

Xtxt="Type.Polymer.Soil"
DATA.plot$X=factor(DATA.plot$Polymer.Soil)

# PLOT ####
  # ORder the DATA frame: 
  # Follow the order in   Label.X
  # DATA= DATA[ order(match(  DATA$X, Label.X[[Xtxt]]$ID)), ] 
  
  # * Data Summary.X*Y ---------------- 
  quantile.3 <-function(x) {
    quantile(x)[[4]]
  }
  
  DATA_Summary.XY=DATA.plot %>% 
    dplyr::group_by(X) %>%
    dplyr::summarise( across(.cols =  Ytxt, .fns = list(Mean=mean, SD=sd, Q3=quantile.3 ), #Q3= quantile.3 #.cols =  Ytxt #.cols = is.numeric 
                             .names = "{col}_{fn}") )  
  
  
  # * Multicompare.X*Y ----------------  
  
  p.adj.method= "holm"#"holm", "hochberg", "hommel", "bonferroni", "BH", "BY", "fdr", "none"
  # 1. check the normality
  p.val_res.Nd=shapiro.test( DATA.plot$Y)[[2]] # p<0.05, the data does not follow a normal distribution
  
  if (p.val_res.Nd<0.05) { # not normal => Kruskal-Wallis then "wilcox.test"
    # 2. check the overall difference with Kruskal-Wallis
    p.val_multi=kruskal.test(Y ~ X, data = DATA.plot)[[3]]
    if (p.val_multi<0.05){
      # 3. Pair-wise comparison with wilcox.test  
      cm=compare_means(Y ~ X, data = DATA.plot, method = "wilcox.test",  p.adjust.method=p.adj.method,  
                       paired = FALSE, group.by = NULL, ref.group = NULL)
    }
    
  } else { # normal => Annova then "t.test"
    # 2. check the overall difference Anova
    p.val_multi=summary(aov(Y ~ X, data = DATA.plot))[[1]][[5]][1]
    if (p.val_multi<0.05){
      # 3. Pair-wise comparison with t.test
      cm=compare_means(Y ~ X, data = DATA.plot, method = "t.test",  p.adjust.method=p.adj.method,
                       paired = FALSE, group.by = NULL, ref.group = NULL)
    }#else if () #/!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\The Assign letter does not work when no sign dif is found.  
      
      
      
  }
  p.val_cm= cm$p.adj
  names(p.val_cm)<- paste(cm$group1,cm$group2,sep = "-")
  
  mcl_cm=multcompLetters(p.val_cm,
                         compare="<=",
                         threshold=0.05,
                         Letters=letters,
                         reversed = FALSE)
  unique( mcl_cm$Letters )
  
  
  
  # * Boxplot Y~X ---------------- 
  # Letters significant differences for PLOT uPs.BoxPlot
  DATA_Summary.XY$Sign.letter="a" #mcl_cm$Letters #mcl_cm[[1]]
  
  # color patette BoxPlot
  
  my_palette.X=list(#Plastic_lab=colors()[c(616,76,373)], 
    Plastic_lab= c( c(brewer.pal(9, "Set1")[c(9,2,5,1)])),
    Treatment_lab= c( c(brewer.pal(9, "Set1")[c(9,2,5,1)]),
                      c(brewer.pal(9, "Set1")[c(9,2,5,1)]),
                      c(brewer.pal(9, "Set1")[c(9,2,5,1)]),
                      c(brewer.pal(9, "Set1")[c(9,2,5,1)]),
                      c(brewer.pal(9, "Set1")[c(9,2,5,1)]))  )
  
  my_palette = my_palette.X[[Xtxt]] 
  
  #BoxPlot_XY=
  ggplot()+
    geom_boxplot(data =DATA.plot, aes(x=X, y=Y, fill=X), outlier.shape = -1, alpha=.66)+
    geom_jitter( data =DATA.plot, aes(x=X, y=Y, fill=X), color="black", width=0.11, size=1, alpha=0.9) +
    #scale_fill_manual(values = my_palette )+
    stat_summary(data =DATA.plot, aes(x=X, y=Y, fill=X), fun=mean, colour="black", geom="point", 
                 shape="+", size=7,  show.legend = FALSE)+ #pch = 24, cex=5, lwd=50,
    geom_text(data = DATA_Summary.XY ,aes(x=X, y=(DATA_Summary.XY[,paste(Ytxt, "Q3", sep="_")][[1]]+ mean(DATA_Summary.XY[,2][[1]])/50) , label=Sign.letter),
              position=position_nudge(x=-0.30, y=0), size=5, colour="black")+
    
    #scale_y_continuous(expand = c(0, 0), limits = c(0, max(DATA$Y)))+
    ylab(label = Title.Plot.Y[[Ytxt]])+ 
    ggtitle(Plot_title)+
    #ylab(label = expression("Plastic debris area [cm"^2 *"/Kg]") ) +#
    theme_minimal()+
    theme(
      plot.title = element_text(size=18),
      axis.title.x = element_blank(),
      axis.title.y = element_text(size=16,hjust = 0.5),
      axis.text.x  = element_text(size=14), #hjust : align the labels 
      axis.text.y  = element_text(size=14),
      legend.title = element_blank(),
      legend.text = element_blank(),
      legend.position="none" ) 



  # * Bar plot ####
  ggplot()+
    geom_bar(data =DATA.plot, aes(x=X, y=Y, fill=X)
             
             