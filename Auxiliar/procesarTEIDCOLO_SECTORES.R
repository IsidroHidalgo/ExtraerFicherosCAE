# (c) Observatorio Regional de Empleo de Castilla-La Mancha
# Isidro Hidalgo Arellano (ihidalgo@jccm.es)

procesarTEIDCOLO_SECTORES <- function(fichero){     
     teidcolosectores = readLines(fichero)
     
     # Extracción del mes
     MES = gsub(" ", "", substr(teidcolosectores[6], 73, 82))
     ANNO = as.numeric(substr(teidcolosectores[6], 93, 96))
     print(paste0("EXTRAYENDO EL FICHERO TEIDCOLO_SECTORES DE ", MES, " DE ", ANNO, "..."))
     mesTeidcolosectores = as.POSIXct(paste0("1", "/", which(meses == MES), "/", ANNO, " 00:00:00"), format = "%d/%m/%Y", tz = "Europe/Paris")
     
     # Extracción de las oficinas
     nLineas <- length(teidcolosectores)
     nOficinas <- nLineas/312
     lineasOficinas <- seq(8, nLineas, by = 312)
     oficinasExtraidas <- substr(teidcolosectores[lineasOficinas], 8, 17)
     oficinas <- OFICINAS[OFICINAS$oficinasExtraidas %in% oficinasExtraidas, ]
     
     # Montaje del data frame
     salida = data.frame(CATEGORÍA = rep(rep(c(NA,'Tramos de edad',NA,'Tramos de edad',NA,
                                               'Nivel de estudios terminados',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Grupo de ocupación solicitado',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA,
                                               'Actividad económica de origen',NA),
                                        times = c(18,10,2,1,9,13,17,5,1,1,1,9,1,1,1,8,1,1,
                                                  1,5,1,1,1,7,9,3,1,1,1,4,1,1,1,8,1,1,1,4,
                                                  1,1,1,8,1,1,1,1,1,1,16,3,1,1,1,34,1,1,10,
                                                  3,1,1,1,41,10,7,1,1,1,1,1,1,4)
                                        ),
                                        nOficinas
                                        ),
                         teidcolosectores = teidcolosectores,
                         stringsAsFactors = FALSE
                         )
     
     salida <- salida[!is.na(salida$CATEGORÍA),]
     
     salida$MES = mesTeidcolosectores
     salida$IDOFICINA <- rep(oficinas$idoficinas, each = 190)
     salida$OFICINA <- rep(oficinas$oficinas, each = 190)
     
     # Extracción de las columnas
     salida$APARTADO = substr(salida$teidcolosectores, 1, 23)
     salida$HOMBRESCONOFERTA = suppressWarnings(as.numeric(substr(salida$teidcolosectores, 24, 36)))
     salida$HOMBRESSINOFERTA = suppressWarnings(as.numeric(substr(salida$teidcolosectores, 37, 49)))
     salida$MUJERESCONOFERTA = suppressWarnings(as.numeric(substr(salida$teidcolosectores, 61, 73)))
     salida$MUJERESSINOFERTA = suppressWarnings(as.numeric(substr(salida$teidcolosectores, 74, 86)))
     salida$HOMBRESCONOFERTA[is.na(salida$HOMBRESCONOFERTA)] <- 0
     salida$HOMBRESSINOFERTA[is.na(salida$HOMBRESSINOFERTA)] <- 0
     salida$MUJERESCONOFERTA[is.na(salida$MUJERESCONOFERTA)] <- 0
     salida$MUJERESSINOFERTA[is.na(salida$MUJERESSINOFERTA)] <- 0
     salida$HOMBRESTOTAL = salida$HOMBRESCONOFERTA + salida$HOMBRESSINOFERTA
     salida$MUJERESTOTAL = salida$MUJERESCONOFERTA + salida$MUJERESSINOFERTA
     salida$CONOFERTATOTAL = salida$HOMBRESCONOFERTA + salida$MUJERESCONOFERTA
     salida$SINOFERTATOTAL = salida$HOMBRESSINOFERTA + salida$MUJERESSINOFERTA
     salida$TOTAL = salida$CONOFERTATOTAL + salida$SINOFERTATOTAL
     
     salida <- salida[, c('MES', 'IDOFICINA', 'OFICINA', 'APARTADO',
                          'HOMBRESCONOFERTA', 'HOMBRESSINOFERTA', 'HOMBRESTOTAL',
                          'MUJERESCONOFERTA', 'MUJERESSINOFERTA', 'MUJERESTOTAL',
                          'CONOFERTATOTAL', 'SINOFERTATOTAL', 'TOTAL')]
     
     print("...FINALIZADO.")
     return(salida)
}
