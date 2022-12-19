library(openxlsx)
library(readr)
library(magrittr)
library(tidyxl)

######################################################################
libro <- "../data/2021/INFO PROGRAMAS 2021/LAGUNAS 2021/RA LAGUNA SUESCA 2021.xlsm"
#convertir a xlsx si hay algún problema en la lectura
######################################################################

hojatidy <- xlsx_cells(libro,sheets = 2)



#libro <-"D:/DiscoD/CAR 2022/Trabajo/Compilados/RA RIO BOGOTA 2021 2 CAMP.xlsm"
#libro <-"../data/RA RIO BOGOTA 2021 2 CAMP.xlsm" ### Cambiar por el nombre de la hoja a analizar, convertir a xlsx si hay algún problema en la lectura
wb <- loadWorkbook(libro)
names(wb)
repetidos2 <- c()
hoja <- readWorkbook(wb, sheet = 2)
hoja <- hojatidy[hojatidy$is_blank == FALSE,]
###Informe
posicion <- grep("INFORME", hoja$character)[1]
posicion <- posicion + 1
while(is.na(hoja$character[posicion])){posicion <- posicion + 1}
informe <- hoja$character[posicion]
###Cliente
posicion <- grep("CLIENTE", hoja$character)[1]
posicion <- posicion + 1
while(is.na(hoja$character[posicion])){posicion <- posicion + 1}
cliente <- hoja$character[posicion]
###Programa
posicion <- grep("PROGRAMA", hoja$character)[1]
posicion <- posicion + 1
while(is.na(hoja$character[posicion])){posicion <- posicion + 1}
programa <- hoja$character[posicion]

###Municipio
posicion <- grep("Municipio", hoja$character)[1]
posicion <- posicion + 1
while(is.na(hoja$character[posicion])){posicion <- posicion + 1}
municipio <- hoja$character[posicion]

###Fecha de muestreo
posicion <- grep("^Muestreo",hoja$character) 
posicion <- posicion + 1
if(is.na(hoja$date[posicion])){
  fecha_muestreo <- hoja$character[posicion]}else{
  fecha_muestreo <- hoja$date[posicion]
}
fecha_muestreo
### Fecha de recepción
posicion <- grep("^Recepción",hoja$character) 
posicion <- posicion + 1
if(is.na(hoja$date[posicion])){
  fecha_recepcion <- hoja$character[posicion]}else{
  fecha_recepcion <- hoja$date[posicion]
}
fecha_recepcion

### Fecha de reporte
posicion <- grep("^Reporte",hoja$character) 
posicion <- posicion + 1
if(is.na(hoja$date[posicion])){
  fecha_reporte <- hoja$character[posicion]}else{
  fecha_reporte <- hoja$date[posicion]
}
fecha_reporte

### Muestras
posiciones_muestra <- grep("Muestra N",hoja$character) 
muestra1 <- hoja$character[posiciones_muestra[1]+1]
muestra2 <- hoja$character[posiciones_muestra[2]+1]
muestra3 <- hoja$character[posiciones_muestra[3]+1]

###############Usar rows y cols para definir posiciones (continuar desde acá)
punto1 <- paste(hoja$X7[10],hoja$X7[11],hoja$X7[12])
punto2 <- paste(hoja$X17[10],hoja$X17[11],hoja$X17[12])
punto3 <- paste(hoja$X24[10],hoja$X24[11],hoja$X24[12])
idpunto1 <- parse_number(hoja$X7[11])
idpunto2 <- parse_number(hoja$X17[11])
idpunto3 <- parse_number(hoja$X24[11])

cond_ambientales <- t((hoja[177:183,c(18,21,23)])) %>% data.frame() ###Está desplazada una fila con respecto a las demás hojas del libro
names(cond_ambientales) <- (hoja$X8[177:183])
cond_ambientales$`Hora de toma` <- as.character(cond_ambientales$`Hora de toma`) %>% as.numeric() %>% convertToDateTime() %>%  format( format = "%H:%M:%S")

georeferenciacion <- data.frame(t(hoja[184:187,c(18,21,23)]))###Está desplazada una fila con respecto a las demás hojas del libro
names(georeferenciacion) <- hoja$X15[184:187]

tabla <- list(informe = informe,cliente = cliente, programa = programa,municipio= municipio,fechamuestreo = fecha_muestreo, fecharecepcion = fecha_recepcion, 
              fechareporte = fecha_reporte, muestra = c(muestra1,muestra2,muestra3), punto = c(punto1,punto2,punto3),id_punto = c(idpunto1,idpunto2,idpunto3), cond_ambientales,
              georeferenciacion) %>% data.frame()

parametros_muestra1 <- hoja[18:165,c(3,8,11,18,19,21)]
parametros_muestra1 <- cbind(muestra1,parametros_muestra1)
colnames(parametros_muestra1) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra1 <- parametros_muestra1[!is.na(parametros_muestra1$valor),]
parametros_muestra1 <- merge(tabla,parametros_muestra1)

parametros_muestra2 <- hoja[18:165,c(3,8,11,18,19,23)]
parametros_muestra2 <- cbind(muestra2,parametros_muestra2)
colnames(parametros_muestra2) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra2 <- parametros_muestra2[!is.na(parametros_muestra2$valor),]
parametros_muestra2 <- merge(tabla,parametros_muestra2)

parametros_muestra3 <- hoja[18:165,c(3,8,11,18,19,25)]
parametros_muestra3 <- cbind(muestra3,parametros_muestra3)
colnames(parametros_muestra3) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra3 <- parametros_muestra3[!is.na(parametros_muestra3$valor),]
parametros_muestra3 <- merge(tabla,parametros_muestra3)

print(names(wb)[2])
print(length(parametros_muestra1$valor))
print(length(parametros_muestra2$valor))
print(length(parametros_muestra3$valor))


consolidado <- rbind(parametros_muestra1,parametros_muestra2,parametros_muestra3)


for(i in 3:length(getSheetNames(libro))){
  hoja <- readWorkbook(wb, sheet = i)
  print(names(wb)[i])
informe <- hoja$X24[1]
cliente <- hoja$X7[2]
programa <- hoja$X7[3]
municipio <- hoja$X8[4]
fecha_muestreo <- convertToDate(hoja$X8[5])
fecha_recepcion <- convertToDate(hoja$X8[6])
fecha_reporte <- convertToDate(hoja$X8[7])
muestra1 <- hoja$X6[10]
muestra2 <- hoja$X16[10]
muestra3 <- hoja$X23[10]
punto1 <- paste(hoja$X7[10],hoja$X7[11],hoja$X7[12])
punto2 <- paste(hoja$X17[10],hoja$X17[11],hoja$X17[12])
punto3 <- paste(hoja$X24[10],hoja$X24[11],hoja$X24[12])
idpunto1 <- parse_number(hoja$X7[11])
idpunto2 <- parse_number(hoja$X17[11])
idpunto3 <- parse_number(hoja$X24[11])

cond_ambientales <- t((hoja[177:183,c(18,21,23)])) %>% data.frame() ###Está desplazada una fila con respecto a las demás hojas del libro
names(cond_ambientales) <- (hoja$X8[177:183])
cond_ambientales$`Hora de toma` <- as.character(cond_ambientales$`Hora de toma`) %>% as.numeric() %>% convertToDateTime() %>%  format( format = "%H:%M:%S")

georeferenciacion <- data.frame(t(hoja[184:187,c(18,21,23)]))###Está desplazada una fila con respecto a las demás hojas del libro
names(georeferenciacion) <- hoja$X15[184:187]

tabla <- list(informe = informe,cliente = cliente, programa = programa,municipio= municipio,fechamuestreo = fecha_muestreo, fecharecepcion = fecha_recepcion,
              fechareporte = fecha_reporte, muestra = c(muestra1,muestra2,muestra3), punto = c(punto1,punto2,punto3),id_punto = c(idpunto1,idpunto2,idpunto3), cond_ambientales,
              georeferenciacion) %>% data.frame()

parametros_muestra1 <- hoja[18:165,c(3,8,11,18,19,21)]
parametros_muestra1 <- cbind(muestra1,parametros_muestra1)
colnames(parametros_muestra1) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra1 <- parametros_muestra1[!is.na(parametros_muestra1$valor),]
parametros_muestra1 <- merge(tabla,parametros_muestra1)

parametros_muestra2 <- hoja[18:165,c(3,8,11,18,19,23)]
parametros_muestra2 <- cbind(muestra2,parametros_muestra2)
colnames(parametros_muestra2) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra2 <- parametros_muestra2[!is.na(parametros_muestra2$valor),]
parametros_muestra2 <- merge(tabla,parametros_muestra2)

parametros_muestra3 <- hoja[18:165,c(3,8,11,18,19,25)]
parametros_muestra3 <- cbind(muestra3,parametros_muestra3)
colnames(parametros_muestra3) <- c("muestra","parametro","unidades","metodo","Tipo de limite","Limite","valor")
parametros_muestra3 <- parametros_muestra3[!is.na(parametros_muestra3$valor),]
parametros_muestra3 <- merge(tabla,parametros_muestra3)


  consolidado <- rbind(consolidado,parametros_muestra1,parametros_muestra2,parametros_muestra3)

  print(length(parametros_muestra1$valor))
  print(length(parametros_muestra2$valor))
  print(length(parametros_muestra3$valor))

}


# Verificaci?n


which(is.na(consolidado$muestra))
consolidado <- consolidado[-which(duplicated(paste0(consolidado$muestra,consolidado$metodo,consolidado$parametro))),]
#verificar si numero de muestro son na arrojar error, contar cuantos parametros po muestra Hay en cada
#parametro 



# 
# which(duplicated(consolidado$muestra))
# 
# consolidado <- consolidado[!is.na(consolidado$muestra),]
# 
# which(duplicated(consolidado$muestra))


write.xlsx(consolidado,"../salidas/Programas/Lagunas/laguna_suesca.xlsx") ### Ajustast el nombre de la salida

