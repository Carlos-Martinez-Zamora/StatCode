# APUNTES

# library(NCmisc)
# list.functions.in.file("Análisis_qPCR_CMZ.R")

# Final<- "Sí"                           # Si es el análisis final, para ajustar las columnas al texto (hay un bug de setcolwidths)
# ggplot--> facet_wrap(~ ifelse(Resultados_tests$Acumulación_relativa < 3, "Valores", "Atípicos"))+

# --> doble barra en el eje?
# --> outliers?

# El ; es para no ver el print de todo lo que ocurre
# En el while, siemore hay que declarar la variable antes. Peligrosos, porque es muy fácil hacer un bucle infinito. 
# Inicializar (poner valor inicial de las variables) antes de un bucle que va a actuar en una columna nueva
# Abrir parentesis seleccionando un trozo lo rodea


# Tibbles funcionan en tidyverse
# control 1 sube a script
# control 2 baja a la consola
# escape borra lo que estás escribiendo en la consola
# Alt+guion pone <- directamente
# la estructura se ve con str
# crtl+shift+m pone |> directamente 
# NULL hace un vector vacío (lo mismo que c())
# seq_along hace una secuencia que recorre un vector entero
# ifelse
# función which
# head() da los primeros seis valores y tail() los últimos seis
# rev() le da la vuelta a los valores (ultimo pasa a primero)
# array=matriz con nombres de columna
# apply() actúa sobre matrices