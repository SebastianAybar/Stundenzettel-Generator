library("tidyverse")
library("tidyr")
library("dplyr")
library("ggplot2")

Stundenlohn <- 12
SV_Brutto <- 520
Stundensatz <- SV_Brutto / Stundenlohn
Arbeitstage <- round((Stundensatz / 2.5))

x_double <- pmax(rnorm(Arbeitstage, mean = 2.5, sd = 1), 0)
x_double <- x_double * (Stundensatz / sum(x_double))
x_double
sum(x_double)
x_time <- (x * 60)


