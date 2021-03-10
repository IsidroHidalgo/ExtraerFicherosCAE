# -----------------------------------------------------------------------------
# (c) Observatorio del Mercado de Trabajo. Junta de Comunidades de Castilla-La Mancha
# (c) Isidro Hidalgo Arellano (ihidalgo@jccm.es)
# -----------------------------------------------------------------------------

# Borramos los objetos en memoria
rm(list = ls(all.names = TRUE))

# Cargamos las librerías necesarias
library(openxlsx)

# Cargamos las variables globales...
meses = c("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
          "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
idoficinas <- c("DA02003000", "OA02003910", "OA02003915", "OA02008910", "OA02009910", "OA02024910",
               "OA02025910", "OA02030910", "OA02037910", "OA02069910", "OA02081910",
               "WW02000107", "OA02003991", "Albacete",
               "DA13034000", "OA13005910", "OA13011910", "OA13013910", "OA13034910", "OA13039910",
               "OA13053910", "OA13063910", "OA13071910", "OA13079910", "OA13082910",
               "OA13087910", "OA13093910", "WW13000107", "OA13034991", "Ciudad Real",
               "DA160078000", "OA16033910", "OA16052910", "OA16078910", "OA16134910", "OA16203910",
               "WW16000107", "OA16078991","Cuenca",
               "DA19130000", "OA19046910", "OA19086910", "OA19130910", "OA19190910",
               "OA19212910", "OA19257910", "WW19000107",
               "OA19130991", "Guadalajara", 
               "DA45168000", "OA45081910", "OA45106910", "OA45121910", "OA45142910", "OA45165910",
               "OA45165915", "OA45168910", "OA45173910", "OA45185910",
               "WW45000107", "OA45168991", "Toledo",
               "Castilla-La Mancha", "Castilla-La Mancha")
oficinas = c("DIR.PROV.AUT.AB", "ALBACETE-CARRETAS", "ALBACETE-CUBAS", "ALCARAZ", "ALMANSA", "CASAS-IBAÑEZ",
               "CAUDETE", "ELCHE DE LA SIERRA", "HELLIN", "LA RODA", "VILLAROBLEDO",
               "VIRTUAL ALBACETE", "DIRECCIÓN PROVINCIAL ALBACETE", "PROVINCIA DE ALBACETE",
               "DIR.PROV.AUT.CR", "ALCAZAR DE SAN JUAN", "ALMADEN", "ALMAGRO", "CIUDAD REAL", "DAIMIEL",
               "MANZANARES", "PIEDRABUENA", "PUERTOLLANO", "LA SOLANA", "TOMELLOSO",
               "VALDEPEÑAS", "VILLANUEVA DE LOS INFANTES", "VIRTUAL CIUDAD REAL",
               "DIRECCIÓN PROVINCIAL CIUDAD REAL", "PROVINCIA DE CIUDAD REAL",
               "DIR.PROV.AUT.CU", "BELMONTE", "CAÑETE", "CUENCA", "MOTILLA DEL PALANCAR", "TARANCON",
               "VIRTUAL CUENCA", "DIRECCIÓN PROVINCIAL CUENCA", "PROVINCIA DE CUENCA",
               "DIR.PROV.AUT.GU", "AZUQUECA DE HENARES", "CIFUENTES", "GUADALAJARA", "MOLINA DE ARAGON",
               "PASTRANA", "SIGUENZA", "VIRTUAL GUADALAJARA",
               "DIRECCIÓN PROVINCIAL GUADALAJARA", "PROVINCIA DE GUADALAJARA",
               "DIR.PROV.AUT.TO", "ILLESCAS", "MORA", "OCAÑA", "QUINTANAR DE LA ORDEN", "TALAVERA DE LA REINA- CERVANTES",
               "TALAVERA DE LA REINA- A.DIVINO", "TOLEDO", "TORRIJOS", "VILLACAÑAS",
               "VIRTUAL TOLEDO", "DIRECCIÓN PROVINCIAL TOLEDO", "PROVINCIA DE TOLEDO",
               "CASTILLA-LA MANCHA", "CASTILLA-LA MANCHA")

# ...encabezados
encabezados <- list()
encabezados[["TDDOEDAD"]] <- c("IDOFICINA", "OFICINA", "TOTAL", "HOMBRES", "MUJERES",
                               "TOTALm25", "HOMBRESm25", "MUJERESm25", "TOTAL2544",
                               "HOMBRES2544", "MUJERES2544", "TOTALM44", "HOMBRESM44",
                               "MUJERESM44")
encabezados[["TDDOSECT"]] <- c("IDOFICINA", "OFICINA", "TOTAL", "AGRICULTURA", "INDUSTRIA",
                               "CONSTRUCCIÓN", "SERVICIOS", "SIN_EMPLEO_ANTERIOR")

# ...puntos de corte
puntosCorte <- list()
puntosCorte[["TDDOSECT"]] <- c(12, 38, 14, 14, 12, 15, 12, 13)
puntosCorte[["TDDOEDAD"]] <- c(10, 33, 7, 10, 10, 14, 10, 10, 14, 10, 10, 14, 10, 13)
# names(puntosCorte) <- c("TDDOSECT", "TDDOEDAD")

# directorio <- choose.dir(caption = "Selecciona el directorio con los archivos a procesar:")
directorio <- "//jclm.es/PROS/SC/OBSERVATORIOEMPLEO/_Extractores/Ficheros"
nombres <- dir(directorio)
if (length(nombres) == 0){
     stop("No hay ficheros que procesar")
}
archivos <- dir(directorio, full.names = TRUE, include.dirs = FALSE)

nombresPartidos <- strsplit(nombres, "_")
nombrePrimero <- sapply(nombresPartidos,"[[", 1)

ficheros <- nombrePrimero

procesar <- ficheros %in% c("TDDOSECT", "TDDOEDAD")
puntosCorte <- puntosCorte[ficheros[procesar]]

if (sum(procesar) == 0){
     print("No hay ningún fichero del tipo correcto en el directorio de procesado.")
} else {
     archivos <- archivos[procesar]
     ficheros <- ficheros[procesar]
     
     if (max(table(ficheros)) > 1) {
     print("No es posible procesar más de un fichero por tipo: elimina duplicados en el directorio de procesado.")
     
     } else {
          # Creamos el fichero Excel de salida
          ficheroSalida <- paste0(directorio,"/Ficheros_Estimación_diaria.xlsx")
          if(file.exists(ficheroSalida)){
               file.remove(ficheroSalida)
               print("Fichero de salida anterior borrado")
          }
          ficheroSalidaExcel <- createWorkbook(ficheroSalida)
          
          # Procesado de los ficheros
     
          for (i in 1:length(ficheros)){
               print(paste0("PROCESANDO EL FICHERO ", ficheros[i], "..."))
               datos <- read.fwf(archivos[i], puntosCorte[i], stringsAsFactors = FALSE)
               datos[, 1] <- gsub(" ", "", datos[, 1])
               
               # Extraemos la fecha y la hora del TDDOSECT:
               ficheroTDDOSECT <- ficheros == "TDDOSECT"
               if (ficheros[i] == "TDDOSECT"){
                    fecha <- sub("-", "", datos[5, 7])
                    fecha <- sub(" ", "", fecha)
                    hora <- unlist(strsplit(sub(" ", "", datos[5, 8]), ":"))
                    hora <- paste0(hora[1], ":", hora[2])
                    fechaHora <- data.frame(Fecha = fecha, Hora = hora)
               }
               
               datos <- datos[datos[ ,1] %in% idoficinas, ]
               names(datos) <- encabezados[[ficheros[i]]]
               cols <- 3:dim(datos)[2]
               datos[,cols] <- apply(datos[,cols], 2, function(x) as.numeric(x))
               addWorksheet(wb = ficheroSalidaExcel, sheetName = ficheros[i])
               writeData(wb = ficheroSalidaExcel, sheet = ficheros[i], x = datos)
               if (ficheros[i] == "TDDOSECT"){
                    writeData(wb = ficheroSalidaExcel, sheet = "TDDOSECT", x = fechaHora, startCol= 9)
               }
               print("...FINALIZADO.")
          }
          # Grabamos el fichero
          saveWorkbook(wb = ficheroSalidaExcel, file = ficheroSalida)
          
          # Mensaje de finalización:
          print(paste0("FICHEROS GUARDADOS COMO ", ficheroSalida))
     }
}

