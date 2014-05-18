library('cmaes')
library('quantmod')
library('zoo')
library('doSNOW')

#setwd('/Users/gregorychevalley/Documents/R/lppl/') # mac mini
#setwd('C:/Users/Gregory Chevalley/RStudio/lppl/') # notebook
setwd('C:/Users/Gregory Chevalley/Documents/lppl/') # server
#setwd('~/lppl/') # vps

rm(list=ls())

fileName <- 'spx.csv'
filePath <- paste('./data/' , fileName, sep='')
fileSName <- substr(filePath,nchar('./data/')+1, nchar(filePath)-4)
ticker <- read.csv(filePath, header=TRUE, sep=",")
ticker <- ticker[with(ticker, order(t)), ]


date_txt_from = "2003-03-12"
date_txt_to_base = "2010-10-31"

nbre_core <- 4 # cores x threads
nbre_step_backward <- 270
nbre_generation <- 100


ticker$Date <- as.Date(ticker$Date)
df_result <- NULL
vec_control <- data.frame(maxit = c(nbre_generation))  
from <- as.Date(date_txt_from)
to_base <- as.Date(date_txt_to_base)




LPPL <- function(data, m=1, omega=1, tc=0) {
  data$X <- tc - data$t
  data$Xm <- data$X ** m #B
  data$Xm.cos <- data$X ** m * cos(omega * log(data$X)) #C1
  data$Xm.sin <- data$X ** m * sin(omega * log(data$X)) #C2
  data$logP <- log(data$Close)
  return(lm(logP ~ Xm + Xm.cos + Xm.sin, data=data))
}

FittedLPPL <- function(data, lm.result, m=1, omega=1, tc=0) {
  data$X <- tc - data$t
  A <- lm.result$coefficients[1]
  B <- lm.result$coefficients[2]
  C1 <- lm.result$coefficients[3]
  C2 <- lm.result$coefficients[4]
  result <- exp(A + B * (data$X ** m) + C1 * (data$X ** m) * cos(omega * log(data$X)) + C2 * (data$X ** m) * sin(omega * log(data$X))) 
  return(result)
}



FittedLPPLwithexpected <- function(data, lm.result, x_vector, m=1, omega=1, tc=0) {
  tmp_vector <- tc - x_vector
  A <- lm.result$coefficients[1]
  B <- lm.result$coefficients[2]
  C1 <- lm.result$coefficients[3]
  C2 <- lm.result$coefficients[4]
  result <- exp(A + B * (tmp_vector ** m) + C1 * (tmp_vector ** m) * cos(omega * log(tmp_vector)) + C2 * (tmp_vector ** m) * sin(omega * log(tmp_vector))) 
  return(result)
  
}

getlinear_param <- function(m, omega, tc) {
  lm.result <- LPPL(rTicker, m, omega, tc)
  return(c(lm.result$coefficients[1],lm.result$coefficients[2], lm.result$coefficients[3], lm.result$coefficients[4]))
}


getlinear_param_with_ts <- function(ts, m, omega, tc) {
  lm.result <- LPPL(ts, m, omega, tc)
  return(c(lm.result$coefficients[1],lm.result$coefficients[2], lm.result$coefficients[3], lm.result$coefficients[4]))
}



tryParams <- function (m, omega, tc) {  
  lm.result <- LPPL(rTicker, m, omega, tc)
  plot(rTicker$t, rTicker$Close, typ='l') #base graph based on data
  generate_vector = seq(min(rTicker$t), tc-0.002, 0.002)
  lines(generate_vector, FittedLPPLwithexpected(rTicker, lm.result, generate_vector, m, omega, tc), col="red")
}


residuals_with_ts <- function(ts, m, omega, tc) {
  lm.result <- LPPL(ts, m, omega, tc)
  return(sum((FittedLPPL(ts, lm.result, m, omega, tc) - ts$Close) ** 2))
}


residuals_with_ts_obj <- function(x, ts) {
  return(residuals_with_ts(ts, x[1], x[2], x[3]))
}


lppl_cmaes <- function(dt_txt_from, dt_from, dt_from_base, step_backward, ts, vec_ctl) {
  dt_to <- dt_from_base
  dt_to <- dt_to-step_backward
  
  if (as.POSIXlt(dt_to)$wday != 0 & as.POSIXlt(dt_to)$wday != 6) { #saute weekend
    sub_ts <- subset(ts, ts$Date >= dt_from & ts$Date <= dt_to)
    last_row <- tail(sub_ts, 1) #pour recuperer le dernier prix et t
    
    test <- cma_es(c(0.1, 5, max(ts$t)+0.002), residuals_with_ts_obj, sub_ts, lower=c(0.1, 5, max(ts$t)+0.002), upper=c(0.9, 16, max(ts$t)+2), control=vec_ctl)
    linear_param <- getlinear_param_with_ts(sub_ts, test$par[1], test$par[2], test$par[3])
    
    return(c(dt_txt_from, format(dt_to, "%Y-%m-%d"), last_row$t, last_row$Close, -step_backward, vec_ctl$maxit, test$par[3]-last_row$t, as.integer((test$par[3]-last_row$t)/(1/365)), test$par[1], test$par[2], test$par[3], linear_param[1], linear_param[2], linear_param[3], linear_param[4]))
  }
  
}



cl <- makeCluster(nbre_core) #nbre core
registerDoSNOW(cl)
s <- system.time(results <- foreach(i = 0:nbre_step_backward, .packages="cmaes", .combine="rbind")  %dopar% {
  lppl_cmaes(date_txt_from, from, to_base, i, ticker, vec_control)
} )

stopCluster(cl)

df_result <- data.frame(results)
colnames(df_result) <- c("date_from", "date_to", "t", "price", "step_backward", "nbre_generation", "t_until_critical_point", "days_before_critical_time", "m", "omega", "tc", "A", "B", "C1", "C2")
nowdatetime <- paste(format(Sys.Date(), "%Y%m%d"), format(Sys.time(), "%H%M%S"), sep="_")
write.csv(df_result, paste('./data/', fileSName, '_parallel_analysis_done_on_', nowdatetime, "_from_", date_txt_from, "_to_", date_txt_to_base, ".csv", sep=''), row.names = FALSE)

print(s)