rm(list = ls())
if (!is.null(dev.list()))
  dev.off()
cat("\014")

library("readxl")

# Importing the data ####

nsim = 10
powerPerActivity = c(0.01,4,3,2,3.5,5,1,2.5,5.5,1.5)
solar_rad <- t(as.matrix(read.csv("solar_rad.csv", header=FALSE)))

for (i in 1:11){ # Man, Woman, Child, Household, ManWomanChild, ManWoman
  assign(paste("Sim",i,sep=""),(t(as.matrix(read_xlsx("Household_activity.xlsx", sheet = as.character(i), col_names=FALSE)))))
}
activityLists <- mapply(get, ls(pattern='Sim*'))
rm(list = ls(pattern='Sim*'))
activityLists[[3]] <- NULL



# Solar System Function ####

simulation_solar <- function(N_bat,act,solar,area){ # solar in Wh/m^2; area in m^2
  
  totalConsumption  = matrix(0,ncol = nsim)
  totalProduction   = matrix(0,ncol = nsim)
  importedEnergy    = matrix(0,ncol = nsim)
  exportedEnergy    = matrix(0,ncol = nsim)
  importRatio       = matrix(0,ncol = nsim)
  exportRatio       = matrix(0,ncol = nsim)
  greenSystemRatio  = matrix(0,ncol = nsim)
  meanStorageLevel  = matrix(0,ncol = nsim)
  meanBatteryOutput = matrix(0,ncol = nsim)
  
  for (sim in 1:nsim){

  # Converting solar iradiation to energy production in kWh
  eff = 0.15
  production = eff*solar*area/10^3
  
  # Converting activity list to power consumption
  consumption = matrix(NA,nrow=nrow(act[[sim]]),
                          ncol=ncol(act[[sim]]))
  
    for (i in 1:nrow(act[[sim]])){
      for (j in 1:ncol(act[[sim]])){
           consumption[i,j] = as.double(powerPerActivity[(as.double(act[[sim]][i,j]) + 1)])
      }
    }
  
  consumption = t(as.matrix(colSums(consumption)))
  
  input = production-consumption # Input of Power production-Power consumption
  
  Battery = list(Output            = matrix(0,ncol=length(input),nrow=1),
                 Storage           = matrix(0,ncol=length(input),nrow=1),
                 maxBatteryStorage = N_bat*13.5,         # Maximum storage  
                 maxBatteryOutput  = N_bat*5)            # Max output in kWh per hour 

  Battery$Storage[1] = Battery$maxBatteryStorage #Starting with fully charged batteries 
  extraEnergy = list(import = list(),
                 export = list())

  a = 1 
  b = 1
  u = 1
  t = 1 
  
  sourcing <- list()
  extra <- list()

  for (i in 1:(length(input)-1)){
    if (input[i] < 0){
        sourcing[u]          = abs(input[i])
        Battery$Output[i]    = pmin(Battery$maxBatteryOutput,Battery$Storage[i])
        if (sourcing[u] > Battery$Output[i]){
            Battery$Storage[i+1]  = Battery$Storage[i] - Battery$Output[i]
            extraEnergy$import[a] = as.double(sourcing[u]) - Battery$Output[i] # Missing Power input
            a = a + 1
        }
        else {
            Battery$Storage[i+1] = Battery$Storage[i] - as.double(sourcing[u])
        }
        u = u + 1
    } 
    else {
        extra[t] = input[i]
        Battery$Storage[i+1] = Battery$Storage[i] + as.double(extra[t])
        if (Battery$Storage[i+1] > Battery$maxBatteryStorage){
            extraEnergy$export[b] = Battery$Storage[i+1] - Battery$maxBatteryStorage
            Battery$Storage[i+1]  = Battery$maxBatteryStorage
            b = b + 1
        }
        t = t + 1
    }
  }
  
  # Totals for each simulation
  totalConsumption[sim] = sum(as.double(consumption)) # total consumed energy
  totalProduction[sim] = sum(as.double(production)) # total produced energy
  
  importedEnergy[sim] = sum(as.double(extraEnergy$import)) # Imported energy when batteries are 0
  exportedEnergy[sim] = sum(as.double(extraEnergy$export)) # Exported energy when batteries are full
  
  exportRatio[sim] = exportedEnergy[sim]/totalProduction[sim]*100 # export energy relative to production
  importRatio[sim] = importedEnergy[sim]/totalConsumption[sim]*100 # how much energy was imported to meet total demand
  greenSystemRatio[sim] = 100-importRatio[sim] # how much energy demand was met by pv panels and battery systems
  
  meanStorageLevel[sim] = mean(as.double(Battery$Storage))/(N_bat*13.5)*100 # mean battery storage level normalized per batter
  meanBatteryOutput[sim] = mean(as.double(Battery$Output))/(N_bat*5)*100 # output normalized per battery
  }
  
  meanTotalCons = mean(totalConsumption)
  meanTotalProd = mean(totalProduction)
  meanImportEnergy = mean(importedEnergy)
  meanExportEnergy = mean(exportedEnergy)
  meanImportRatio = mean(importRatio)
  meanExportRatio = mean(exportRatio)
  meanSystemRatio = mean(greenSystemRatio)
  mean_meanStorLevel = mean(meanStorageLevel)
  mean_meanBatOut = mean(meanBatteryOutput)
  
  
  # Uncertainty calculations 
  CI_tco = matrix(0,ncol=2)
  CI_tpr = matrix(0,ncol=2)
  CI_ien = matrix(0,ncol=2)
  CI_een = matrix(0,ncol=2)
  CI_ira = matrix(0,ncol=2)
  CI_era = matrix(0,ncol=2)
  CI_gsr = matrix(0,ncol=2)
  CI_msl = matrix(0,ncol=2)
  CI_mbo = matrix(0,ncol=2)
  
  CI_tco[1] = meanTotalCons + qt(0.975,(nsim-1))*sd(totalConsumption)
  CI_tco[2] = meanTotalCons - qt(0.975,(nsim-1))*sd(totalConsumption)
  
  CI_tpr[1] = meanTotalProd + qt(0.975,(nsim-1))*sd(totalProduction)
  CI_tpr[2] = meanTotalProd - qt(0.975,(nsim-1))*sd(totalProduction)
  
  CI_ien[1] = meanImportEnergy + qt(0.975,(nsim-1))*sd(importedEnergy)
  CI_ien[2] = meanImportEnergy - qt(0.975,(nsim-1))*sd(importedEnergy)
  
  CI_een[1] = meanExportEnergy + qt(0.975,(nsim-1))*sd(exportedEnergy)
  CI_een[2] = meanExportEnergy - qt(0.975,(nsim-1))*sd(exportedEnergy)
  
  CI_ira[1] = meanImportRatio + qt(0.975,(nsim-1))*sd(importRatio)
  CI_ira[2] = meanImportRatio - qt(0.975,(nsim-1))*sd(importRatio)
  
  CI_era[1] = meanExportRatio + qt(0.975,(nsim-1))*sd(exportRatio)
  CI_era[2] = meanExportRatio - qt(0.975,(nsim-1))*sd(exportRatio)
  
  CI_gsr[1] = meanSystemRatio + qt(0.975,(nsim-1))*sd(greenSystemRatio)
  CI_gsr[2] = meanSystemRatio - qt(0.975,(nsim-1))*sd(greenSystemRatio)
  
  CI_msl[1] = mean_meanStorLevel + qt(0.975,(nsim-1))*sd(meanStorageLevel)
  CI_msl[2] = mean_meanStorLevel - qt(0.975,(nsim-1))*sd(meanStorageLevel)
  
  CI_mbo[1] = mean_meanBatOut + qt(0.975,(nsim-1))*sd(meanBatteryOutput)
  CI_mbo[2] = mean_meanBatOut - qt(0.975,(nsim-1))*sd(meanBatteryOutput)
  
  output <<- list("totalCons"=meanTotalCons, "totalProd" = meanTotalProd,
                "import"=meanImportEnergy,"export"=meanExportEnergy, "meanStor"=mean_meanStorLevel, 
                "meanBatOut" = mean_meanBatOut ,"importRatio"=meanImportRatio,
                "gsystemRatio"=meanSystemRatio, "exportRatio"=meanExportRatio,
                
                "CI_for_cons" = CI_tco, "CI_for_prod" = CI_tpr, "CI_for_import" = CI_ien,
                "CI_for_export" = CI_een, "CI_for_import_ratio" = CI_ira, 
                "CI_for_system_ratio" = CI_gsr, "CI_for_mean_storage" = CI_msl,
                "CI_for_mean_bat_out" = CI_mbo, "CI_for_export_ratio" = CI_era
                )
}



#----------------------------------------------------------------
#Running Simulation for various number of batteries and Areas ####

for (NofBatt in seq(1,20,1)){ 
  Area = 1400 # Change it accordingly 
  if (NofBatt < 10)
  assign(paste("NofBatteries = 0.",NofBatt,sep=" "),simulation_solar(NofBatt,activityLists,solar_rad,Area))
  else if (NofBatt >= 10 && NofBatt < 20 )
  assign(paste("NofBatteries = 1.",NofBatt,sep=" "),simulation_solar(NofBatt,activityLists,solar_rad,Area))
  else if (NofBatt >= 20 && NofBatt < 30)
  assign(paste("NofBatteries = 2.",NofBatt,sep=" "),simulation_solar(NofBatt,activityLists,solar_rad,Area))
}

BatterySims <- mapply(get, ls(pattern='NofBatt*'))
rm(list = ls(pattern='NofBatt*')) # correct the wrong output
BatterySims[[1]] <- NULL
NofBatt = 20

#-------Import Ratio data 
import_ratio = matrix(NA,ncol = NofBatt)
CI_low_ir = matrix(NA,ncol = NofBatt)
CI_high_ir = matrix(NA,ncol = NofBatt)
for (i in 1:NofBatt){
  import_ratio[i] = BatterySims[[i]]$importRatio
  CI_low_ir[i]       = BatterySims[[i]]$CI_for_import_ratio[1,2]
  CI_high_ir[i]      = BatterySims[[i]]$CI_for_import_ratio[1,1]
}
lines(1:NofBatt,import_ratio,
     xlim = c(0,NofBatt),ylim = c(0,100),
     ylab = "Import Ratio [%]",xlab = "Number of batteries",col='red',pch = 20,'b',lty=1)
lines(1:NofBatt,CI_low_ir,col='red',lty=2)
lines(1:NofBatt,CI_high_ir,col='red',lty = 2)

#grid(nx = 20, ny = 20, col = "lightgray", lty = "dotted")
#title(main = "Import Ratio")


#-------Export Ratio data 
export_ratio = matrix(NA,ncol = NofBatt)
CI_low_er = matrix(NA,ncol = NofBatt)
CI_high_er = matrix(NA,ncol = NofBatt)
for (i in 1:NofBatt){
  export_ratio[i] = BatterySims[[i]]$exportRatio
  CI_low_er[i]       = BatterySims[[i]]$CI_for_export_ratio[1,2]
  CI_high_er[i]      = BatterySims[[i]]$CI_for_export_ratio[1,1]
}
lines(1:NofBatt,export_ratio,
     xlim = c(0,NofBatt),ylim = c(0,100),
     ylab = "Export Ratio [%]",xlab = "Number of batteries",col='red',pch = 20,'b',lty=1)
lines(1:NofBatt,CI_low_er,col='red',lty=2)
lines(1:NofBatt,CI_high_er,col='red',lty = 2)

#grid(nx = 20, ny = 20, col = "lightgray", lty = "dotted")
#title(main = "Export Ratio")


#--------Green System Ratio
gs_ratio = matrix(NA,ncol = NofBatt)
CI_low_gs = matrix(NA,ncol = NofBatt)
CI_high_gs = matrix(NA,ncol = NofBatt)
for (i in 1:NofBatt){
  gs_ratio[i] = BatterySims[[i]]$gsystemRatio
  CI_low_gs[i]       = BatterySims[[i]]$CI_for_system_ratio[1,2]
  CI_high_gs[i]      = BatterySims[[i]]$CI_for_system_ratio[1,1]
}
lines(1:NofBatt,gs_ratio,
     xlim = c(0,NofBatt),ylim = c(0,100),
     ylab = "Green System Ratio [%]",xlab = "Number of batteries",col='red',pch = 20,'b',lty=1)
lines(1:NofBatt,CI_low_gs,col='red',lty=2)
lines(1:NofBatt,CI_high_gs,col='red',lty = 2)

#grid(nx = 20, ny = 20, col = "lightgray", lty = "dotted")
#title(main = "System Ratio")

#--------Battery properties 
mean_stor = matrix(NA,ncol = NofBatt)
CI_low_mstor = matrix(NA,ncol = NofBatt)
CI_high_mstor = matrix(NA,ncol = NofBatt)
for (i in 1:NofBatt){
  mean_stor[i] = BatterySims[[i]]$meanStor
  CI_low_mstor[i]       = BatterySims[[i]]$CI_for_mean_storage[1,2]
  CI_high_mstor[i]      = BatterySims[[i]]$CI_for_mean_storage[1,1]
}
lines(1:NofBatt,mean_stor,
     xlim = c(0,NofBatt),ylim = c(0,100),
     ylab = "Mean Storage Level [% of maximum capacity]",xlab = "Number of batteries",col='red',pch = 20,'b',lty=1)
lines(1:NofBatt,CI_low_mstor,col='red',lty=2)
lines(1:NofBatt,CI_high_mstor,col='red',lty = 2)

#grid(nx = 20, ny = 20, col = "lightgray", lty = "dotted")
#title(main = "Normalized Battery Storage Level ")



#--------Battery properties 2
mean_out = matrix(NA,ncol = NofBatt)
CI_low_out = matrix(NA,ncol = NofBatt)
CI_high_out = matrix(NA,ncol = NofBatt)
for (i in 1:NofBatt){
  mean_out[i] = BatterySims[[i]]$meanBatOut
  CI_low_out[i]       = BatterySims[[i]]$CI_for_mean_bat_out[1,2]
  CI_high_out[i]      = BatterySims[[i]]$CI_for_mean_bat_out[1,1]
}
lines(1:NofBatt,mean_out,
     xlim = c(0,NofBatt),ylim = c(0,100),
     ylab = "Mean Output Level [% of maximum capacity]",xlab = "Number of batteries",col='red',pch = 20,'b',lty=1,
     )
lines(1:NofBatt,CI_low_out,col='red',lty=2)
lines(1:NofBatt,CI_high_out,col='red',lty = 2)

#grid(nx = 20, ny = 20, col = "lightgray", lty = "dotted")
#title(main = "Normalized Battery Output Level ")

legend("topleft", legend=c("Area = 100", "Area = 250","Area = 400", "Area = 550","Area = 750","Area = 1000","Area = 1400"),
       fill = c("black", "honeydew4", "firebrick4", "blue", "#CCCC00", "#009900", "red"), 
       cex=0.75,inset=0.05,horiz=TRUE,title=expression(paste("Area shown in [", m^2,"]")),box.lty=0)

