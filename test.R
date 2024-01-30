library("tidyverse")
library("tidyr")
library("dplyr")
library("ggplot2")

#Mean automatisch anheben bei höherem SV-Brutto
Stundenlohn <- 12
SV_Brutto <- 520
Stundensatz <- SV_Brutto / Stundenlohn
Arbeitstage <- round((Stundensatz / 2.5))

#Wir erstellen einen Vektor mit so vielen normalverteilten Zufallswerten wie die Anzahl an Arbeitstagen ist
#Es dürfen keine negativen Outlayers geben deshalb nutzen wir pmax() um immer mindestens 0 Stunden zu bekommen
x_double <- pmax(rnorm(Arbeitstage, mean = 2.5, sd = 1), 0)
#Nun haben wir einen Vektor dessen Summe aus den erarbeiteten Stunden, die nicht genau gleich wie die Variable Stundensatz ist.
#Also müssen wir die einzelnen Werte des Vektors so berechnen, dass die beiden Werte übereinstimmen
x_double <- x_double * (Stundensatz / sum(x_double))
#Da wir nicht wissen wie wir in Excel einspielen, so bekommen wir die Anzahl an Minuten.
x_time <- (x_double * 60)

#Vorschlag: Wir befüllen erst die Spalte "Dezimal" mit doubles und danach die Spalte "Arbeitszeiten"


###########################################################################################################################
###### Nun müssen wir die 17 Arbeitstage von oben auf die potenziellen Werktage des Monats random verteilen. ##############


#Anzahl der Elemente im neuen Vektor
neue_Vector_Groesse <- 25

#Erstellen eines leeren Vektors mit der neuen Größe
neuer_Vector <- numeric(neue_Vector_Groesse)

#zufällige Indizes für die Positionen der vorhandenen Werte
zufalls_Indizes <- sample(neue_Vector_Groesse, Arbeitstage)

#die vorhandenen Werte den zufälligen Positionen im neuen Vektor zuweisen
neuer_Vector[zufalls_Indizes] <- x_double
neuer_Vector

