#                                #
#  DOI: 10.5281/zenodo.15585615  #
#                                #


library(shiny)
library(readxl)  

# Interfaz de usuario: define la estructura visual de la app (título, pestañas y paneles de opciones)
ui <- fluidPage(
    tags$style(HTML("
      .well .shiny-input-container input[type='text'],
      .well .shiny-input-container input[type='number'],
      .well .shiny-input-container select { width: 150px !important; min-width: 150px !important; max-width: 180px !important; box-sizing: border-box; }
      .well .shiny-input-container select { height: 34px !important; min-height: 34px !important; padding: 2px 6px !important; line-height: 1.4 !important; overflow: hidden; text-overflow: ellipsis; }
      .well .input-row { display: flex; align-items: center; gap: 8px; }
      .well .input-row > label { margin-bottom: 0; margin-right: 4px; min-width: 150px; flex-shrink: 0; }
      .well .shiny-input-container .checkbox { display: flex; align-items: center; }
      .well .shiny-input-container .checkbox label { display: flex !important; align-items: center; margin-left: 6px; }
      .well .shiny-input-container .checkbox input[type='checkbox'] { flex-shrink: 0; }
      /* Botones deshabilitados: estilo gris y sin interacción */
      .btn.disabled, .btn[disabled] {
        opacity: 0.65;
        cursor: not-allowed;
        pointer-events: none;
      }
    ")),
    tags$script(HTML("
      Shiny.addCustomMessageHandler('toggleButton', function(message) {
        var el = document.getElementById(message.id);
        if (!el) return;
        if (message.disabled) {
          el.classList.add('disabled');
          el.setAttribute('disabled', 'disabled');
        } else {
          el.classList.remove('disabled');
          el.removeAttribute('disabled');
        }
      });
    ")),
############ Título y botón en la parte superior
    fluidRow(
        column(8,
            titlePanel("StatCode"),
            tags$p(
              class = "text-muted",
              style = "font-size: 10px; margin-top: -15px;",
              "DOI: 10.5281/zenodo.15585615"
            )
        ),
        column(4, align="right",
            div(style="margin-top: 20px;",
                downloadButton("downloadResult", "Download file", class = "btn-default", style = "margin-right: 8px;"),
                actionButton("startButton", "Start Analysis", 
                            style="color: #fff; background-color: #337ab7; border-color: #2e6da4; padding: 10px 24px; font-size: 18px;"),
                br(),
                actionLink("helpButton", "Help", style = "font-size: 11px; color: #777;")
            )
        )
    ),

########### Panel de mensajes justo debajo del título
    fluidRow(
        column(12,
            verbatimTextOutput("status")
        )
    ),
    
########### Contenido principal: tres pestañas (sin títulos en paneles)
    tabsetPanel(
        tabPanel("Data",
            fluidRow(
                column(4,
                    wellPanel(
                        fileInput("archivo_fuente",
                                 "Select Excel file:",
                                 multiple = FALSE,
                                 accept = c(".xlsx")),
                        div(class = "input-row",
                            tags$label("Data:", style = "margin-bottom: 0; margin-right: 0px;"),
                            selectInput("crear_datos", label = NULL, choices = c("qPCR Cq", "Final Data"), selected = "qPCR Cq")
                        ),
                        div(class = "input-row",
                            tags$label("Data sheet:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("hoja_fuente", label = NULL, value = 2, min = 1)
                        ),
                        div(class = "input-row",
                            tags$label("Selection:", style = "margin-bottom: 0; margin-right: 5px;")
                        ),
                        
                        div(style = "display: flex; align-items: center;", checkboxInput("pegar_grafica", "Paste graph", value = TRUE)),
                        div(style = "display: flex; align-items: center;", checkboxInput("exportar_analisis_estadistico", "Paste statistical analysis", value = TRUE)),
                        hr(),
                        div(class = "input-row",
                            tags$label("Subgroup:", style = "margin-bottom: 0; margin-right: 5px;"),
                            selectInput("subgrupo", label = NULL, choices = c("Yes", "No"), selected = "No")
                        ),
                        conditionalPanel(
                            condition = "input.subgrupo == 'Yes'",
                            div(
                                style = "border: 1px solid #ddd; padding: 10px; margin-top: 10px;",
                                uiOutput("columnas_checkboxes")
                            )
                        ),
                        div(class = "input-row",
                            tags$label("Cq replicas:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("replicas", label = NULL, value = 3, min = 1)
                        ),
                        div(class = "input-row",
                            tags$label("Cq difference limit:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("diferencia_ct", label = NULL, value = 1, min = 0, step = 0.1)
                        ),
                        div(class = "input-row",
                            tags$label("Decimals:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("decimales", label = NULL, value = 4, min = 0)
                        ),
                        div(class = "input-row",
                            tags$label("Decimal marker:", style = "margin-bottom: 0; margin-right: 5px;"),
                            selectInput(
                                "comas",
                                label = NULL,
                                choices = c("." = "No", "," = "Yes"),
                                selected = "No"
                            )
                        )
                    )
                ),
                column(8,
                    wellPanel(
                        fluidRow(
                            column(6,
                                div(class = "input-row",
                                    tags$label("Graph name:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    textInput("nombre_grafica", label = NULL, value = "Graph")
                                ),
                                div(class = "input-row",
                                    tags$label("Y axis name:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    textInput("nombre_eje_y", label = NULL, value = "Relative accumulation")
                                ),
                                div(class = "input-row",
                                    tags$label("X axis name:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    textInput("nombre_eje_x", label = NULL, value = "Line")
                                )
                            ),
                            column(6,
                                div(class = "input-row",
                                    tags$label("Graph type:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    selectInput("grafico", label = NULL, choices = c("Boxplot", "Bars", "Violins"), selected = "Boxplot")
                                ),
                                div(class = "input-row",
                                    tags$label("Graph format:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    selectInput("formato_grafica", label = NULL, choices = c("jpeg", "png", "svg", "tiff", "pdf", "eps", "ps", "tex", "bmp", "wmf"), selected = "jpg")
                                ),
                                div(class = "input-row",
                                    tags$label("Resolution:", style = "margin-bottom: 0; margin-right: 5px;"),
                                    numericInput("resolucion", label = NULL, value = 400)
                                )
                            )
                        ),
                        hr(),
                        div(style = "min-height: 280px;",
                            imageOutput("grafica", height = "380px")
                        ),
                        downloadButton("downloadGraph", "Descargar imagen", class = "btn-default")
                    )
                )
            )
        ),
        tabPanel("Statistics",
            wellPanel(
                fluidRow(
                    column(6,
                        div(class = "input-row",
                            tags$label("Normality p-value:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("p_valor_normalidad", label = NULL, value = 0.05, min = 0, max = 1, step = 0.01)
                        ),
                        div(class = "input-row",
                            tags$label("Cedasticity p-value:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("p_valor_cedasticidad", label = NULL, value = 0.05, min = 0, max = 1, step = 0.01)
                        ),
                        div(class = "input-row",
                            tags$label("Difference p-value:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("p_valor_diferencias", label = NULL, value = 0.05, min = 0, max = 1, step = 0.01)
                        )
                    ),
                    column(6,
                        div(class = "input-row",
                            tags$label("Posthoc p-value:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("p_valor_posthoc", label = NULL, value = 0.05, min = 0, max = 1, step = 0.01)
                        ),
                        div(class = "input-row",
                            tags$label("Sample number limit:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("num_muestras", label = NULL, value = 50, min = 1)
                        )
                    )
                )
            )
        ),
        tabPanel("Aesthetic",
            wellPanel(
                fluidRow(
                    column(6,
                        div(class = "input-row",
                            tags$label("Fill color:", style = "margin-bottom: 0; margin-right: 5px;"),
                            textInput("color_relleno", label = NULL, value = "white")
                        ),
                        div(class = "input-row",
                            tags$label("Border color:", style = "margin-bottom: 0; margin-right: 5px;"),
                            textInput("color_bordes", label = NULL, value = "black")
                        ),
                        div(class = "input-row",
                            tags$label("Point color:", style = "margin-bottom: 0; margin-right: 5px;"),
                            textInput("color_puntos", label = NULL, value = "black")
                        ),
                        div(class = "input-row",
                            tags$label("Error bar color:", style = "margin-bottom: 0; margin-right: 5px;"),
                            textInput("color_barra_error", label = NULL, value = "black")
                        ),
                        div(class = "input-row",
                            tags$label("Background color:", style = "margin-bottom: 0; margin-right: 5px;"),
                            textInput("color_fondo", label = NULL, value = "white")
                        ),
                        div(class = "input-row",
                            tags$label("Point size:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("tamano_puntos", label = NULL, value = 1.5, step = 0.1)
                        ),
                        div(class = "input-row",
                            tags$label("Point shape:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("forma_puntos", label = NULL, value = 16, min = 0)
                        ),
                        div(class = "input-row",
                            tags$label("Sideways scatter:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("movimiento_lateral", label = NULL, value = 0.1, step = 0.01)
                        ),
                        div(class = "input-row",
                            tags$label("Box width:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("ancho_cajas", label = NULL, value = 0.5, step = 0.1)
                        ),
                        div(class = "input-row",
                            tags$label("Error bar width:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("ancho_barra_error", label = NULL, value = 0.15, step = 0.01)
                        )
                        
                    ),
                    column(6,
                           div(class = "input-row",
                               tags$label("Border width:", style = "margin-bottom: 0; margin-right: 5px;"),
                               numericInput("grosor_borde", label = NULL, value = 0.5, step = 0.1)
                           ),
                           div(class = "input-row",
                               tags$label("Median line width:", style = "margin-bottom: 0; margin-right: 5px;"),
                               numericInput("grosor_mediana", label = NULL, value = 1, step = 0.1)
                           ),
                        div(class = "input-row",
                            tags$label("Y axis start:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("inicio", label = NULL, value = 0)
                        ),
                        div(class = "input-row",
                            tags$label("Y axis marks:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("marcas", label = NULL, value = 1)
                        ),
                        div(class = "input-row",
                            tags$label("Graph title size:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("tamano_titulo_ejes", label = NULL, value = 17, min = 1, step = 0.5)
                        ),
                        div(class = "input-row",
                            tags$label("Graph text size:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("tamano_texto_ejes", label = NULL, value = 14, min = 1, step = 0.5)
                        ),
                        div(class = "input-row",
                            tags$label("Y axis cut distance:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("distancia_corte_puntos", label = NULL, value = 0.5, step = 0.1)
                        ),
                        div(class = "input-row",
                            tags$label("Units:", style = "margin-bottom: 0; margin-right: 5px;"),
                            selectInput("unidades", label = NULL, choices = c("px", "in", "cm", "mm"), selected = "px")
                        ),
                        div(class = "input-row",
                            tags$label("Width:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("ancho", label = NULL, value = 4000)
                        ),
                        div(class = "input-row",
                            tags$label("Height:", style = "margin-bottom: 0; margin-right: 5px;"),
                            numericInput("alto", label = NULL, value = 3500)
                        )
                    )
                )
            )
        )
    )
) 

server <- function(input, output, session) {
############ Valores reactivos para guardar variables
    column_data <- reactiveVal()
    analysis_running <- reactiveVal(FALSE)
    stored_values <- reactiveVal(list())
    last_analysis_values <- reactiveVal(NULL)
    last_graph_path <- reactiveVal(NULL)
    last_result_path <- reactiveVal(NULL) 
    
    # Sección de ayuda: muestra una ventana flotante al pulsar el botón "Help"
    observeEvent(input$helpButton, {
        showModal(modalDialog(
            "Make sure the data is in the selected sheet of the document and that it is arranged following the explanation in the README file (Respository).",
            title = "Help",
            easyClose = TRUE,
            size = "s"
        ))
    })
    
    # Inicializa los valores por defecto de todos los controles de la interfaz
    observe({
        if (is.null(stored_values())) {
            stored_values(list(
                # File input (initially NULL)
                archivo_fuente = NULL,
                
                # Text inputs
                nombre_grafica = "Graph",
                nombre_eje_y = "Relative accumulation",
                nombre_eje_x = "Line",
                color_relleno = "white",
                color_bordes = "black",
                color_puntos = "black",
                color_barra_error = "black",
                color_fondo = "white",
                
                # Numeric inputs
                hoja_fuente = 2,
                replicas = 3,
                diferencia_ct = 1,
                mock = 0,
                decimales = 4,
                p_valor_normalidad = 0.05,
                p_valor_cedasticidad = 0.05,
                p_valor_diferencias = 0.05,
                p_valor_posthoc = 0.05,
                num_muestras = 50,
                tamano_puntos = 1.5,
                forma_puntos = 16,
                movimiento_lateral = 0.1,
                ancho_cajas = 0.5,
                ancho_barra_error = 0.15,
                grosor_borde = 0.5,
                grosor_mediana = 1,
                inicio = 0,
                marcas = 1,
                resolucion = 400,
                ancho = 4000,
                alto = 3500,
                # Nuevas variables
                tamano_texto_ejes = 14,
                tamano_titulo_ejes = 17,
                distancia_corte_puntos = 0.5,
                
                # Select inputs
                crear_datos = "qPCR Cq",
                exportar_resultados = "Yes",
                pegar_grafica = "Yes",
                exportar_analisis_estadistico = "Yes",
                subgrupo = "No",
                comas = "No",
                grafico = "Boxplot",
                formato_grafica = "jpg",
                unidades = "px"
            ))
            last_analysis_values(stored_values())
        }
    })
    
############ Actualizar valores guardados al cambiar cualquier control de la interfaz
    observe({
        # Forzar reactividad en todos los inputs
        input$nombre_grafica
        input$nombre_eje_y
        input$nombre_eje_x
        input$color_relleno
        input$color_bordes
        input$color_puntos
        input$color_barra_error
        input$color_fondo
        input$hoja_fuente
        input$replicas
        input$diferencia_ct
        input$mock
        input$decimales
        input$p_valor_normalidad
        input$p_valor_cedasticidad
        input$p_valor_diferencias
        input$p_valor_posthoc
        input$num_muestras
        input$tamano_puntos
        input$forma_puntos
        input$movimiento_lateral
        input$ancho_cajas
        input$ancho_barra_error
        input$grosor_borde
        input$grosor_mediana
        input$inicio
        input$marcas
        input$resolucion
        input$ancho
        input$alto
        input$tamano_texto_ejes
        input$tamano_titulo_ejes
        input$distancia_corte_puntos
        input$crear_datos
        input$exportar_resultados
        input$pegar_grafica
        input$exportar_analisis_estadistico
        input$subgrupo
        input$comas
        input$grafico
        input$formato_grafica
        input$unidades

        values <- list(
            archivo_fuente = if (!is.null(input$archivo_fuente)) list(
                name = input$archivo_fuente$name,
                datapath = input$archivo_fuente$datapath
            ) else NULL,

            nombre_grafica = input$nombre_grafica,
            nombre_eje_y = input$nombre_eje_y,
            nombre_eje_x = input$nombre_eje_x,
            color_relleno = input$color_relleno,
            color_bordes = input$color_bordes,
            color_puntos = input$color_puntos,
            color_barra_error = input$color_barra_error,
            color_fondo = input$color_fondo,
            hoja_fuente = input$hoja_fuente,
            replicas = input$replicas,
            diferencia_ct = input$diferencia_ct,
            mock = input$mock,
            decimales = input$decimales,
            p_valor_normalidad = input$p_valor_normalidad,
            p_valor_cedasticidad = input$p_valor_cedasticidad,
            p_valor_diferencias = input$p_valor_diferencias,
            p_valor_posthoc = input$p_valor_posthoc,
            num_muestras = input$num_muestras,
            tamano_puntos = input$tamano_puntos,
            forma_puntos = input$forma_puntos,
            movimiento_lateral = input$movimiento_lateral,
            ancho_cajas = input$ancho_cajas,
            ancho_barra_error = input$ancho_barra_error,
            grosor_borde = input$grosor_borde,
            grosor_mediana = input$grosor_mediana,
            inicio = input$inicio,
            marcas = input$marcas,
            resolucion = input$resolucion,
            ancho = input$ancho,
            alto = input$alto,
            # Nuevas variables
            tamano_texto_ejes = input$tamano_texto_ejes,
            tamano_titulo_ejes = input$tamano_titulo_ejes,
            distancia_corte_puntos = input$distancia_corte_puntos,
            crear_datos = input$crear_datos,
            exportar_resultados = if (isTRUE(input$exportar_resultados)) "Yes" else "No",
            pegar_grafica = if (isTRUE(input$pegar_grafica)) "Yes" else "No",
            exportar_analisis_estadistico = if (isTRUE(input$exportar_analisis_estadistico)) "Yes" else "No",
            subgrupo = input$subgrupo,
            comas = input$comas,
            grafico = input$grafico,
            formato_grafica = input$formato_grafica,
            unidades = input$unidades
        )
        stored_values(values)
    })
    
    # Lee el archivo de Excel seleccionado y lo guarda como datos reactivos
    observeEvent(input$archivo_fuente, {
        req(input$archivo_fuente)
        tryCatch({
            df <- read_excel(input$archivo_fuente$datapath, 
                           sheet = input$hoja_fuente)
            column_data(df)
        }, error = function(e) {
            showNotification(paste("Error reading file:", e$message), 
                           type = "error")
        })
    })
    
    # Actualiza los datos cuando el usuario cambia el número de hoja del Excel
    observeEvent(input$hoja_fuente, {
        req(input$archivo_fuente)
        tryCatch({
            df <- read_excel(input$archivo_fuente$datapath, 
                           sheet = input$hoja_fuente)
            column_data(df)
        }, error = function(e) {
            showNotification(paste("Error reading file:", e$message), 
                           type = "error")
        })
    })
    
########### Generar cajas UI para seleccionar columnas de subgrupo
    output$columnas_checkboxes <- renderUI({
        req(column_data())
        df <- column_data()
        
        div(
            style = "display: flex; flex-wrap: wrap; gap: 10px;",
            lapply(seq_len(ncol(df)), function(i) {
                col_name <- names(df)[i]
                if (is.null(col_name) || is.na(col_name) || col_name == "") {
                    col_name <- paste("Column", i)
                }
                
                div(
                    style = "margin: 5px; text-align: center; min-width: 120px;",
                    checkboxInput(
                        inputId = paste0("col_", i),
                        label = col_name,
                        value = FALSE
                    )
                )
            })
        )
    })
    
########### Mostrar gráfica generada (dimensiones fijas para evitar errores de viewport no finito)
   # 1. Definir el contador al inicio del server
actualizar_grafica <- reactiveVal(0)

# 2. En la parte donde GENERAS la imagen, suma 1
# plot_grafica(...) 
# actualizar_grafica(actualizar_grafica() + 1)

# 3. Tu salida modificada
output$grafica <- renderImage({
  path <- last_graph_path()
  if (is.null(path) || !file.exists(path)) {
    return(list(src = "", contentType = "image/png", width = 1L, height = 1L, alt = "No graph"))
  }
  ext <- tolower(tools::file_ext(path))
  mime <- switch(ext, jpg = "image/jpeg", jpeg = "image/jpeg", png = "image/png",
                 gif = "image/gif", webp = "image/webp", "image/png")
  list(src = path, contentType = mime, width = 800L, height = 380L, alt = "Graph")
}, deleteFile = FALSE)


########### Descargar imagen de la gráfica generada a un archivo
    output$downloadGraph <- downloadHandler(
        filename = function() {
            path <- last_graph_path()
            if (is.null(path) || !file.exists(path)) return("graph.png")
            basename(path)
        },
        content = function(file) {
            path <- last_graph_path()
            if (!is.null(path) && file.exists(path)) file.copy(path, file)
        }
    )
    
    # Habilitar / deshabilitar visualmente el botón de descarga de gráfica
    observe({
        path <- last_graph_path()
        disabled <- is.null(path) || !file.exists(path)
        session$sendCustomMessage("toggleButton", list(
            id = "downloadGraph",
            disabled = disabled
        ))
    })
    
########### Descargar archivo de resultados (Excel con las salidas del análisis)
    output$downloadResult <- downloadHandler(
        filename = function() {
            path <- last_result_path()
            if (is.null(path) || !file.exists(path)) return("results.xlsx")
            basename(path)
        },
        content = function(file) {
            path <- last_result_path()
            if (!is.null(path) && file.exists(path)) file.copy(path, file)
        }
    )
    
    # Habilitar / deshabilitar visualmente el botón de descarga de resultados
    observe({
        path <- last_result_path()
        disabled <- is.null(path) || !file.exists(path)
        session$sendCustomMessage("toggleButton", list(
            id = "downloadResult",
            disabled = disabled
        ))
    })
    
########### Función análisis: prepara el script externo con los parámetros de la interfaz y lo ejecuta
    run_analysis <- function(values) {
        req(values$archivo_fuente)  # Ensure we have file information
        tryCatch({
            # Escapar barras invertidas en reemplazos para evitar "argumento 'replacement' inválido" en gsub
            esc <- function(x) {
                if (is.null(x) || (length(x) == 1L && is.na(x))) return("")
                x <- as.character(x)
                gsub("\\\\", "\\\\\\\\", x, fixed = TRUE)
            }
            # Valores numéricos seguros (evitar NA en replacement)
            num <- function(x, default = 0) if (length(x) && !is.na(x) && is.finite(x)) x else default
            
            # Generar un nombre de archivo de salida nuevo para no modificar el original
            ext_in <- tools::file_ext(values$archivo_fuente$name)
            base_in <- tools::file_path_sans_ext(values$archivo_fuente$name)
            if (is.null(ext_in) || ext_in == "") {
                ext_in <- "xlsx"
            }
            result_name <- paste0(base_in, "_StatCode.", ext_in)
            
            # Copiar el archivo cargado a un nuevo archivo de trabajo en el directorio actual
            file.copy(values$archivo_fuente$datapath,
                      result_name,
                      overwrite = TRUE)

            script_content <- strsplit(r"(
            #                                #
#  DOI: 10.5281/zenodo.15585615  #
#                                #



setwd(dirname(rstudioapi::getSourceEditorContext()$path))

######################### ESSENTIAL: #######################################
Archivo_fuente<-"Ahora.xlsx"    # Original file name
Nombre_Gráfica<-"Ahora"     # Graph name (with format, e.g: Graph.jpg)
Formato_grafica <- "jpg"                # Exported graph format 


######################### CREATE: #######################################
Crear_Datos<-"Create"                   # "Create" / "Import" ("Create" means data is raw Ct values, to calculate foldchange. "Import" means data is already calculated and ready to create graph and statistical analysis)
Crear_Gráfica<-"Yes"                    # "Yes" / "No" (Create a graph or not)
Crear_Análisis_Estadístico<-"Yes"       # "Yes" / "No" (Create statistical analysis or not)
Hoja_fuente<-2                          # Sheet with the data in the original file


######################### EXPORT: #######################################
Exportar_Resultados<-"No"              # Results: "Yes" / "No"
Exportar_Gráfica<-"Yes"                  # Graph: "Yes" / "No"
Pegar_Gráfica<-"Yes"                    # Paste graph in excel: "Yes" / "No"
Exportar_Análisis_Estadístico<-"No"    # Statistical analysis: "Yes" / "No"


######################### MODIFIABLE VARIABLES: #######################################
################ DATA:
Subgrupo<-"No"                          # Exclude Ct columns: "Yes" / "No"
Columnas<-c(1,2,11,12)                  # Columns to be used (write columns to be included)
Réplicas<-3                             # Number of technical replicas
Diferencia_Ct <- 1                      # Ct difference limit (if the difference between any of the replicas is larger than this number, the sample will be excluded)
Mock<- 0                                # Number of mock (will be excluded. To be written at the end)
Comas<-"No"                             # Export decimals as commas
Decimales<-4                            # Number of decimals in exported tables


################ GRAPH:
Gráfico<-"Boxplot"                      # Type of graph: "Boxplot" / "Bars" / "Violins"
Nombre_eje_y <- "Relative accumulation" # y axis name
Nombre_eje_x <- "Line"                  # x axis name
Color_relleno <- "white"                # Graph element filling colour (All the colours you can use: https://r-charts.com/es/colores/)
Color_bordes <- "black"                 # Graph border colour
Color_puntos <- "black"                 # Graph point colour
Color_barra_error <- "black"            #  Error bar colour (only bar plot)
Color_fondo <- "white"                  # Background colour (for exported graph)
Tamaño_puntos <- 1.5                    # Point size in graph
Forma_puntos <- 16                      # Point shape in graph (guide: https://blog.albertkuo.me/post/point-shape-options-in-ggplot/)
Movimiento_lateral <- 0.1               # Sideways jitter of points in graph
Ancho_cajas <- 0.5                      # Width of the graph element (box/bar/violin)
Ancho_barra_error <- 0.15               # Error bar width
Grosor_borde <- 0.5                     # Border line width
Grosor_mediana <- 1                     # Median line width
Inicio<-0                               # Y axis starting point
Marcas<-1                               # Y axis marking
Distancia_corte_puntos<-0.5             # Y axis cut difference
Resolución<-400                         # Graph image resolution
Unidades<-"px"                          # Units of graph image size
Ancho<-4000                             # Graph image width
Alto<-3500                              # Graph image height
Tamaño_titulo_ejes<-17                  # Axis title font size
Tamaño_texto_ejes<-14                   # Axis text font size


################ STATISTICAL ANALYSIS:
p_valor_normalidad<- 0.05               # Normality test p-value
p_valor_cedasticidad<- 0.05             # -cedasticity test p-value
p_valor_diferencias<- 0.05              # Difference test p-value (ANOVA or equivalent)
p_valor_posthoc<- 0.05                  # Post hoc test p-value (t-student is unpaired)
Num_muestras<- 50                       # Number of samples to change from SW to KS/AD (Normality)
 
 
###########################################################################################
library(openxlsx)
library(stats)
library(readxl)
if(Comas=="No"){
  options(OutDec=".")}else{
    options(OutDec=",")
  }

if(Crear_Datos=="Create"){
  ############## IMPORTAR DATOS CT #############################
  datos_crudos <- read_excel(Archivo_fuente, sheet = Hoja_fuente)
  if(Subgrupo == "Yes"){
    datos_crudos<-datos_crudos[,Columnas]
  }
  
  ############## CALCULAR MEDIAS DE LAS RÉPLICAS TÉCNICAS CT #############################
  
  medias_Ct=data.frame(); #Declarar data frame vacío para rellenarlo con el bucle
  for(j in 1:ncol(datos_crudos)){
    k=1 #Filas en la tabla de las medias_Cts de las réplicas técnicas
    i=2 #Filas de la tabla de datos crudos
    while(i <= nrow(datos_crudos)){
      x=as.double(datos_crudos[i,j])+as.double(datos_crudos[i+1,j])+as.double(datos_crudos[i+2,j]); 
      medias_Ct[k,j]<-x/Réplicas
      trio<-c(as.double(datos_crudos[i,j]), as.double(datos_crudos[i+1,j]), as.double(datos_crudos[i+2,j]))
      if(is.na(abs(medias_Ct[k,j])) | 
         abs(max(trio)-min(trio))>=Diferencia_Ct){
        medias_Ct[k,j] = NaN;
      }
      k=k+1;
      i=i+3;
    }
    
  }
  colnames(medias_Ct)<-colnames(datos_crudos);
  
  
  ##################### CALCULAR dCT  #############################
  k=1;
  j=1;
  x=as.double(nrow(medias_Ct));
  col=as.double(ncol(medias_Ct)/2);
  dCt<-as.data.frame(matrix(0,ncol=col,nrow=x));
  while(j<=ncol(medias_Ct)){
    dCt[,k]<-(as.double(medias_Ct[,j])-as.double(medias_Ct[,j+1]))
    k=k+1
    j=j+2
  }
  colnames(dCt)<-colnames(datos_crudos[,seq(1,ncol(datos_crudos),2)])
  
  
  ##################### CALCULAR ddCT #############################
  k=1;
  i=2;
  x=as.double(nrow(dCt))-1;
  col=as.double(ncol(dCt));
  ddCt<-as.data.frame(matrix(0,ncol=col,nrow=x));
  while(i<=nrow(dCt)){
    ddCt[k,]<-(as.double(dCt[i,])-as.double(dCt[1,]))
    k=k+1
    i=i+1
  }
  colnames(ddCt)<-colnames(dCt)
  
  
  ################## CALCULAR 2^-ddCT (foldchange) #############################
  k=1;
  i=1;
  x=as.double(nrow(ddCt));
  col=as.double(ncol(ddCt));
  foldchange<-as.data.frame(matrix(0,ncol=col,nrow=x));
  while(i<=nrow(ddCt)){
    foldchange[k,]<-2^-(as.double(ddCt[i,]))
    k=k+1
    i=i+1
  }
  colnames(foldchange)<-colnames(ddCt)
  
  
  ################### TABLA RESULTADOS #############################
  Final_Resultados<-foldchange 
}

if(Crear_Datos=="Import"){
  ############## IMPORTAR DATOS CT #############################
  datos_crudos_lista<- read_excel(Archivo_fuente, sheet = Hoja_fuente)
  datos_crudos <- as.data.frame(
    lapply(datos_crudos_lista, 
           function(col){
             as.double(gsub(",", ".", col))
           }
    )
  )
  
  if(Subgrupo == "Yes"){
    datos_crudos<-datos_crudos[,Columnas]
  }
  
  ################### TABLA RESULTADOS #############################
  Final_Resultados<-datos_crudos
}

################### EXPORTAR RESULTADOS #############################
if(Crear_Datos=="Create"){
  if(Exportar_Resultados=="Yes") {

    wb<-loadWorkbook(Archivo_fuente)
    
    if(any(sheets(wb)== "Results")){
      nombres_hojas <- excel_sheets(Archivo_fuente)
      repes <- sum(grepl("Results", nombres_hojas, ignore.case = TRUE))
      pagina_resultados <- paste("Results", repes)
      addWorksheet(wb, pagina_resultados)
    }else{
      pagina_resultados<- "Results"
      addWorksheet(wb, pagina_resultados)
    }
    
    setColWidths(wb, pagina_resultados, cols = 1:ncol(Final_Resultados), widths = 12)
    
    Tabla1<-Final_Resultados
    Tabla1<-format(round(Final_Resultados, Decimales))
    
    j=1
    while(j <= ncol(Tabla1)) {
      i=1
      while(i <= nrow(Tabla1)){
        if (Final_Resultados[i,j]=="NaN") {
          Tabla1[i,j]<-" "
        }
        i=i+1
      }
      j=j+1
    }
    
    numeros<-createStyle(fontName = "Microsoft Sans Serif", fontSize = 9, halign = "right")
    addStyle(wb, pagina_resultados, style=numeros, rows=2:(nrow(Tabla1)+1), cols=1:ncol(Tabla1), gridExpand = TRUE)
  
    if(Comas=="No"){
      options(OutDec=".")
      }else{
      Tabla1 <- apply(Tabla1, c(1,2), function(x) {
     if (x != "") gsub("\\.", ",", x) else x})
    }
    
    writeData(wb,pagina_resultados, Tabla1)
    saveWorkbook(wb,Archivo_fuente,overwrite=TRUE)  
      
    
  }  
}



#################### DISPONER DATOS PARA TESTS ESTADÍSTICOS #############################
x=ncol(Final_Resultados)*nrow(Final_Resultados)
Datos_tests<-as.data.frame(matrix(,x));
Datos_tests[,1] <-data.frame(unlist(Final_Resultados));
Resultados_tests <- data.frame(Acumulación_relativa = Datos_tests, 
                               Line=names(Final_Resultados)[col(Final_Resultados)], stringsAsFactors=FALSE)
Resultados_tests$Line <- factor(Resultados_tests$Line, levels = unique(Resultados_tests$Line))
colnames(Resultados_tests)<- c("Acumulación_relativa", "Line")


#####################                #############################
##################### REPRESENTACIÓN #############################
#####################                #############################

if(Crear_Gráfica=="Yes"){
  
  library(ggbreak)
  library(ggplot2)
  library(dplyr)
  
  Resultados_tests_g <- Resultados_tests
  
  
  
  if(Gráfico == "Boxplot"){
    # Límite superior:
    lim.sup=max(na.omit(Resultados_tests_g[,1]))+0.25
    
    ## Cortes de ejes:
    # Cual es el punto más alto dentro de los bigotes 
    j=1
    rango.bigotes<-matrix(0,nrow=ncol(Final_Resultados),ncol = 1);
    while(j <= ncol(Final_Resultados)){
      valores_ordenados<-sort(na.omit(Final_Resultados[,j]))
      Q1 <- quantile(valores_ordenados, 0.25)
      Q3 <- quantile(valores_ordenados, 0.75)
      rango.bigotes[j,]<-(Q3+1.5*(Q3-Q1));
      j=j+1
    }
    
    j=1
    tops<-matrix(0,nrow=ncol(Final_Resultados),ncol = 1)
    while(j <= ncol(Final_Resultados)){
      valores.bigotes<-subset(Final_Resultados[,j], Final_Resultados[,j]<rango.bigotes[j,])
      tops[j,]<-max(valores.bigotes);
      j=j+1
    }
    aparte<-subset(Resultados_tests_g[,1], Resultados_tests_g[,1]>max(tops[,1]))
    
    ristra<-vector()
    ristra<-append(max(tops[,1]), aparte)
    ristra<-sort(ristra)
    
    parejas <- matrix(0,ncol=2, nrow =0)
    
    if(length(ristra)>=2){
      for (i in 1:(length(ristra) - 1)) {
        if (abs(ristra[i] - ristra[i + 1]) > Distancia_corte_puntos) {
          fila_nueva<-c(ristra[i],ristra[i + 1])
          parejas<-rbind(parejas, fila_nueva)
        }
       }
      
      if(nrow(parejas)>=1){
        i=1
        cortes<-c()
        while(i<=nrow(parejas)){
          cortes<-c(cortes,parejas[i,1]+0.25, parejas[i,2]-0.25)
          i=i+1
        }
      }
    }
    
    Final_Gráfica <- ggplot(Resultados_tests_g, 
                            aes(x=Line, y=Acumulación_relativa))  + 
      stat_boxplot(geom = "errorbar",
                   width = Ancho_barra_error,
                   colour = Color_bordes,
                   size = Grosor_borde)+ 
      geom_boxplot(notch=FALSE,
                   outlier.shape=NA,
                   colour = Color_bordes,
                   fill = Color_relleno,
                   width = Ancho_cajas,
                   size = Grosor_borde,
                   fatten = Grosor_mediana)+
      geom_jitter(width = Movimiento_lateral,
                  height = 0,
                  shape = Forma_puntos, 
                  size = Tamaño_puntos,
                  colour = Color_puntos)+
      xlab(Nombre_eje_x) + 
      ylab(Nombre_eje_y)+
      scale_y_continuous(breaks = seq(Inicio, lim.sup, by = Marcas),limits = c(Inicio, lim.sup))+
      theme_light()+
      theme(
        axis.title.x = element_text(size = Tamaño_titulo_ejes),  
        axis.title.y = element_text(size = Tamaño_titulo_ejes),
        axis.text.x = element_text(size = Tamaño_texto_ejes),  
        axis.text.y = element_text(size = Tamaño_texto_ejes)
      )

    if(nrow(parejas)>=1){ 
      Final_Gráfica <- Final_Gráfica+
                       scale_y_break(cortes)
    }
   }
  
  if(Gráfico=="Bars"){
    
    Resumen <- Resultados_tests_g %>%
      group_by(Line) %>%
      summarise(
        Media = mean(Acumulación_relativa, na.rm = TRUE),
        SD = sd(Acumulación_relativa, na.rm = TRUE)
      )
    
    tops<-matrix(0, nrow=nrow(Resumen), ncol=1)
    
    i=1
    while(i<=nrow(Resumen)){
      tops[i,1]<-(Resumen$Media[i]+Resumen$SD[i])
      i=i+1
    }
    
    # Límite superior:
    lim.sup=max(na.omit(Resultados_tests_g[,1]))+0.25
    
    ## Cortes de ejes:
    aparte<-subset(Resultados_tests_g[,1], Resultados_tests_g[,1]>max(tops[,1]))
    ristra<-vector()
    ristra<-append(max(tops[,1]), aparte)
    ristra<-sort(ristra)
    
    parejas <- matrix(0,ncol=2, nrow =0)
    if(length(ristra)>=2){
      for (i in 1:(length(ristra) - 1)) {
        if (abs(ristra[i] - ristra[i + 1]) > 0.5) {
          fila_nueva<-c(ristra[i],ristra[i + 1])
          parejas<-rbind(parejas, fila_nueva)
        }
      }
      
      if(nrow(parejas)>=1){
        i=1
        cortes<-c()
        while(i<=nrow(parejas)){
          cortes<-c(cortes,parejas[i,1]+0.25, parejas[i,2]-0.25)
          i=i+1
        }
      }
    }
    
    Final_Gráfica <-ggplot(Resumen, aes(x=Line, y=Media)) +
      geom_bar(stat = "identity", 
               fill=Color_relleno, 
               colour= Color_bordes,
               width = Ancho_cajas,
               size = Grosor_borde) +
      geom_errorbar(aes(ymax = Media + SD, ymin = Media - SD), 
                    width = Ancho_barra_error,
                    colour= Color_barra_error,
                    size = Grosor_borde) +
      geom_point(data=Resultados_tests_g, aes(x=Line, y=Acumulación_relativa), 
                 position = position_jitter(width = Movimiento_lateral),
                 colour= Color_puntos,
                 size = Tamaño_puntos, 
                 shape = Forma_puntos,
                 alpha = 1)+
      scale_y_continuous(breaks = seq(Inicio, lim.sup, by = Marcas),
                         limits = c(Inicio, lim.sup))+
      labs(x = Nombre_eje_x, y = Nombre_eje_y) +
      theme_light()+
      theme(
        axis.title.x = element_text(size = Tamaño_titulo_ejes),  
        axis.title.y = element_text(size = Tamaño_titulo_ejes),
        axis.text.x = element_text(size = Tamaño_texto_ejes),  
        axis.text.y = element_text(size = Tamaño_texto_ejes)
      )
    
    if(nrow(parejas)>=1){ 
      Final_Gráfica <-Final_Gráfica+
        scale_y_break(cortes)
    }
  }
  
  if(Gráfico == "Violins"){
    # Límite superior:
    lim.sup=max(na.omit(Resultados_tests_g[,1]))+0.25
    
    Final_Gráfica <- ggplot(Resultados_tests_g, 
                            aes(x = Line, y = Acumulación_relativa)) + 
      stat_boxplot(geom = "errorbar",
                   width = Ancho_barra_error,
                   colour = Color_bordes) + 
      geom_violin(fill = Color_relleno,
                  colour = Color_bordes,
                  width = Ancho_cajas,
                  size = Grosor_borde) +
      geom_jitter(width = Movimiento_lateral,
                  height = 0,
                  shape = Forma_puntos, 
                  size = Tamaño_puntos,
                  colour = Color_puntos) +
      xlab(Nombre_eje_x) + 
      ylab(Nombre_eje_y) +
      scale_y_continuous(breaks = seq(Inicio, lim.sup, by = Marcas), limits = c(Inicio, lim.sup)) +
      theme_light()+
      theme(
        axis.title.x = element_text(size = Tamaño_titulo_ejes),  
        axis.title.y = element_text(size = Tamaño_titulo_ejes),
        axis.text.x = element_text(size = Tamaño_texto_ejes),  
        axis.text.y = element_text(size = Tamaño_texto_ejes)
      )
    
   }
  
  if(Exportar_Gráfica=="Yes" | Pegar_Gráfica=="Yes" ){
    Nombre_Gráfica<-paste0(Nombre_Gráfica,".", Formato_grafica)
    ggsave(Nombre_Gráfica, 
           plot = Final_Gráfica, 
           device = Formato_grafica, 
           dpi = Resolución,
           width = Ancho, 
           height = Alto, 
           units = Unidades,
           pointsize = Tamaño_puntos,
           bg=Color_fondo)
    }
    
    if(Pegar_Gráfica=="Yes"){
    wb<-loadWorkbook(Archivo_fuente)
 
      if(Crear_Datos=="Create"){
        insertImage(wb,pagina_resultados,Nombre_Gráfica, startRow=2, startCol = (ncol(Final_Resultados)+2))
      }
      if(Crear_Datos=="Import"){
        insertImage(wb, Hoja_fuente,Nombre_Gráfica, startRow=2, startCol = (ncol(Final_Resultados)+2))
      }
    
    saveWorkbook(wb,Archivo_fuente,overwrite=TRUE)  
  }  
  
  if(Exportar_Gráfica=="No"){
  file.remove(Nombre_Gráfica)
  }
  
  Final_Gráfica
}


#####################                      #############################
##################### ANÁLISIS ESTADÍSTICO #############################
#####################                      #############################

if(Crear_Análisis_Estadístico=="Yes"){
  
  library(DescTools)
  library(car)
  library(rstatix)
  library(broom)
  library(PMCMRplus)
  
  ##################### TEST NORMALIDAD #############################
  k=1
  j=1
  Normalidad<-as.data.frame(matrix(0,ncol=6));
  if(nrow(Final_Resultados) <= Num_muestras){
    while(j <= ncol(Final_Resultados)){
      columna<-c(Final_Resultados[,j])
      Normalidad[k,] <-shapiro.test(columna)
      k=k+1;
      j=j+1}
    Normalidad<-Normalidad[,c(3,6,2,4,5,1)]
    colnames(Normalidad)<-c("Test", "Line", "p-value", "Significance level", "Result", "W")
    
  }else{
    library(nortest)
    if(any(duplicated(Final_Resultados[, 1:j]))){
      while(j <= ncol(Final_Resultados)){
        Normalidad[k,] <- ad.test(Final_Resultados[,j])
        k=k+1;
        j=j+1        
      }
      Normalidad<-Normalidad[,c(3,6,2,4,5,1)]
      colnames(Normalidad)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
    } else{
      while(j <= ncol(Final_Resultados)){
        Normalidad[k,] <- ks.test(Final_Resultados[,j], "pnorm",mean(Final_Resultados[,j]), sd(Final_Resultados[,j]))
        k=k+1;
        j=j+1
      }
      Normalidad<-Normalidad[,c(4,6,2,3,5,1)]
      colnames(Normalidad)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
      Normalidad[,1]="Kolmogorov-Smirnov test"
    }
  }
  Normalidad[,2]=c(colnames(Final_Resultados))
  Normalidad[,4]=p_valor_normalidad
  
  i=1
  for(i in 1:nrow(Normalidad)){
    if(Normalidad[i,3] > Normalidad[i,4]){
      Normalidad[i,5]= "Normal"
    }else(
      Normalidad[i,5]= "Not normal"
    )
    i=i+1
  }
  
  Final_Normalidad<-Normalidad[,1:5]
  Resultados_usados<-Final_Resultados
  
  ############################# TRANSFORMAR
  Resultados_sin_NA<-as.vector(na.omit(Resultados_tests[,1]))
  if(any(Resultados_sin_NA==0)){
    ################sqrt
    if(any(Normalidad[,5] == "Not normal")){
      x=as.double(nrow(Final_Resultados));
      col=as.double(ncol(Final_Resultados));
      Resultados_transformados_2<-as.data.frame(matrix(0,ncol=col,nrow=x));
      k=1
      j=1
      while(j <=col){
        Resultados_transformados_2[,k] <- as.double(sqrt(Final_Resultados[,j]))
        j=j+1
        k=k+1
      }
      
      k=1
      j=1
      Normalidad_transformados_2<-as.data.frame(matrix(0,ncol=6));
      if(nrow(Final_Resultados) <= Num_muestras){
        while(j <= ncol(Final_Resultados)){
          columna<-c(Resultados_transformados_2[,j])
          Normalidad_transformados_2[k,] <- shapiro.test(columna)
          k=k+1;
          j=j+1
        }
        Normalidad_transformados_2<-Normalidad_transformados_2[,c(3,6,2,4,5,1)]
        colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
      }else{
        library(nortest)
        if(any(duplicated(Final_Resultados[, 1:j]))){
          while(j <= ncol(Final_Resultados)){
            Normalidad_transformados_2[k,] <- ad.test(Resultados_transformados_2[,j])
            k=k+1;
            j=j+1        
          }
          Normalidad_transformados_2<-Normalidad_transformados_2[,c(3,6,2,4,5,1)]
          colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
        } else{
          while(j <= ncol(Final_Resultados)){
            Normalidad_transformados_2[k,] <- ks.test(Resultados_transformados_2[,j], "pnorm",mean(Resultados_transformados_2[,j]), sd(Resultados_transformados_2[,j]))
            k=k+1;
            j=j+1
          }
          Normalidad_transformados_2<-Normalidad_transformados_2[,c(4,6,2,3,5,1)]
          colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
          Normalidad_transformados_2[,1]="Kolmogorov-Smirnov test"
        }
      }
      Normalidad_transformados_2[,2]=c(colnames(Final_Resultados))
      Normalidad_transformados_2[,4]=p_valor_normalidad
      
      i=1
      for(i in 1:nrow(Normalidad_transformados_2)){
        if(Normalidad_transformados_2[i,3] > Normalidad_transformados_2[i,4]){
          Normalidad_transformados_2[i,5]= "Normal [sqrt]"
        }else(
          Normalidad_transformados_2[i,5]= "Not normal"
        )
i=i+1
    }

Final_Normalidad<-Normalidad_transformados_2
Resultados_usados<-Resultados_transformados_2

if(any(Final_Normalidad[,5] == "Not normal")){
  Final_Normalidad<-Normalidad[,1:5]
  Resultados_usados<-Final_Resultados
}  
}
}


if(all(Resultados_sin_NA>0)){
  ################log
  if(any(Normalidad[,5] == "Not normal")){
    x=as.double(nrow(Final_Resultados));
    col=as.double(ncol(Final_Resultados));
    Resultados_transformados<-as.data.frame(matrix(0,ncol=col,nrow=x));
    k=1
    j=1
    while(j <=col){
      Resultados_transformados[,k] <- as.double(log(Final_Resultados[,j]))
      j=j+1
      k=k+1
    }
    
    k=1
    j=1
    Normalidad_transformados<-as.data.frame(matrix(0,ncol=6));
    if(nrow(Final_Resultados) <= Num_muestras){
      while(j <= ncol(Final_Resultados)){
        columna<-c(Resultados_transformados[,j])
        Normalidad_transformados[k,] <- shapiro.test(columna)
        k=k+1;
        j=j+1
      }
      Normalidad_transformados<-Normalidad_transformados[,c(3,6,2,4,5,1)]
      colnames(Normalidad)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
    }else{
      library(nortest)
      if(any(duplicated(Final_Resultados[, 1:j]))){
        while(j <= ncol(Final_Resultados)){
          Normalidad_transformados[k,] <- ad.test(Resultados_transformados[,j])
          k=k+1;
          j=j+1        
        }
        Normalidad_transformados<-Normalidad_transformados[,c(3,6,2,4,5,1)]
        colnames(Normalidad_transformados)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
      } else{
        while(j <= ncol(Final_Resultados)){
          Normalidad_transformados[k,] <- ks.test(Resultados_transformados[,j], "pnorm",mean(Resultados_transformados[,j]), sd(Resultados_transformados[,j]))
          k=k+1;
          j=j+1
        }
        Normalidad_transformados<-Normalidad_transformados[,c(4,6,2,3,5,1)]
        colnames(Normalidad_transformados)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
        Normalidad_transformados[,1]="Kolmogorov-Smirnov test"
      }
    }
    Normalidad_transformados[,2]=c(colnames(Final_Resultados))
    Normalidad_transformados[,4]=p_valor_normalidad
    
    i=1
    for(i in 1:nrow(Normalidad_transformados)){
      if(Normalidad_transformados[i,3] > Normalidad_transformados[i,4]){
        Normalidad_transformados[i,5]= "Normal [ln]"
      }else(
        Normalidad_transformados[i,5]= "Not normal"
      )
      i=i+1
    }
    
    Final_Normalidad<-Normalidad_transformados
    Resultados_usados<-Resultados_transformados
    
    
    ################sqrt
    if(any(Normalidad_transformados[,5] == "Not normal")){
      x=as.double(nrow(Final_Resultados));
      col=as.double(ncol(Final_Resultados));
      Resultados_transformados_2<-as.data.frame(matrix(0,ncol=col,nrow=x));
      k=1
      j=1
      while(j <=col){
        Resultados_transformados_2[,k] <- as.double(sqrt(Final_Resultados[,j]))
        j=j+1
        k=k+1
      }
      
      k=1
      j=1
      Normalidad_transformados_2<-as.data.frame(matrix(0,ncol=6));
      if(nrow(Final_Resultados) <= Num_muestras){
        while(j <= ncol(Final_Resultados)){
          columna<-c(Resultados_transformados_2[,j])
          Normalidad_transformados_2[k,] <- shapiro.test(columna)
          k=k+1;
          j=j+1
        }
        Normalidad_transformados_2<-Normalidad_transformados_2[,c(3,6,2,4,5,1)]
        colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
      }else{
        library(nortest)
        if(any(duplicated(Final_Resultados[, 1:j]))){
          while(j <= ncol(Final_Resultados)){
            Normalidad_transformados_2[k,] <- ad.test(Resultados_transformados_2[,j])
            k=k+1;
            j=j+1        
          }
          Normalidad_transformados_2<-Normalidad_transformados_2[,c(3,6,2,4,5,1)]
          colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
        } else{
          while(j <= ncol(Final_Resultados)){
            Normalidad_transformados_2[k,] <- ks.test(Resultados_transformados_2[,j], "pnorm",mean(Resultados_transformados_2[,j]), sd(Resultados_transformados_2[,j]))
            k=k+1;
            j=j+1
          }
          Normalidad_transformados_2<-Normalidad_transformados_2[,c(4,6,2,3,5,1)]
          colnames(Normalidad_transformados_2)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
          Normalidad_transformados_2[,1]="Kolmogorov-Smirnov test"
        }
      }
      Normalidad_transformados_2[,2]=c(colnames(Final_Resultados))
      Normalidad_transformados_2[,4]=p_valor_normalidad
      
      i=1
      for(i in 1:nrow(Normalidad_transformados_2)){
        if(Normalidad_transformados_2[i,3] > Normalidad_transformados_2[i,4]){
          Normalidad_transformados_2[i,5]= "Normal [sqrt]"
        }else(
          Normalidad_transformados_2[i,5]= "Not normal"
        )
        i=i+1
      }
      
      Final_Normalidad<-Normalidad_transformados_2
      Resultados_usados<-Resultados_transformados_2
      
      
      ################inverse
      if(any(Normalidad_transformados_2[,5] == "Not normal")){
        x=as.double(nrow(Final_Resultados));
        col=as.double(ncol(Final_Resultados));
        Resultados_transformados_3<-as.data.frame(matrix(0,ncol=col,nrow=x));
        k=1
        j=1
        while(j <=col){
          Resultados_transformados_3[,k] <- as.double(1/(Final_Resultados[,j]))
          j=j+1
          k=k+1
        }
        
        k=1
        j=1
        Normalidad_transformados_3<-as.data.frame(matrix(0,ncol=6));
        if(nrow(Final_Resultados) <= Num_muestras){
          while(j <= ncol(Final_Resultados)){
            columna<-c(Resultados_transformados_3[,j])
            Normalidad_transformados_3[k,] <- shapiro.test(columna)
            k=k+1;
            j=j+1
          }
          Normalidad_transformados_3<-Normalidad_transformados_3[,c(3,6,2,4,5,1)]
          colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
        }else{
          library(nortest)
          if(any(duplicated(Final_Resultados[, 1:j]))){
            while(j <= ncol(Final_Resultados)){
              Normalidad_transformados_3[k,] <- ad.test(Resultados_transformados_3[,j])
              k=k+1;
              j=j+1        
            }
            Normalidad_transformados_3<-Normalidad_transformados_3[,c(3,6,2,4,5,1)]
            colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
          }else{
            while(j <= ncol(Final_Resultados)){
              Normalidad_transformados_3[k,] <- ks.test(Resultados_transformados_3[,j], "pnorm",mean(Resultados_transformados_3[,j]), sd(Resultados_transformados_3[,j]))
              k=k+1;
              j=j+1
            }
            Normalidad_transformados_3<-Normalidad_transformados_3[,c(4,6,2,3,5,1)]
            colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
            Normalidad_transformados_3[,1]="Kolmogorov-Smirnov test"
          }
        }
        Normalidad_transformados_3[,2]=c(colnames(Final_Resultados))
        Normalidad_transformados_3[,4]=p_valor_normalidad
        
        i=1
        for(i in 1:nrow(Normalidad_transformados_3)){
          if(Normalidad_transformados_3[i,3] > Normalidad_transformados_3[i,4]){
            Normalidad_transformados_3[i,5]= "Normal [1/x]"
          }else(
            Normalidad_transformados_3[i,5]= "Not normal"
          )
          i=i+1
        }
        
        Final_Normalidad<-Normalidad_transformados_3
        Resultados_usados<-Resultados_transformados_3
      }
    }
  }
  
  if(any(Final_Normalidad[,5] == "Not normal")){
    Final_Normalidad<-Normalidad[,1:5]
    Resultados_usados<-Final_Resultados
  }
}

if(all(Resultados_sin_NA<0)){
  
  ################inverse
  if(any(Normalidad[,5] == "Not normal")){
    x=as.double(nrow(Final_Resultados));
    col=as.double(ncol(Final_Resultados));
    Resultados_transformados_3<-as.data.frame(matrix(0,ncol=col,nrow=x));
    k=1
    j=1
    while(j <=col){
      Resultados_transformados_3[,k] <- as.double(1/(Final_Resultados[,j]))
      j=j+1
      k=k+1
    }
    
    k=1
    j=1
    Normalidad_transformados_3<-as.data.frame(matrix(0,ncol=6));
    if(nrow(Final_Resultados) <= Num_muestras){
      while(j <= ncol(Final_Resultados)){
        columna<-c(Resultados_transformados_3[,j])
        Normalidad_transformados_3[k,] <- shapiro.test(columna)
        k=k+1;
        j=j+1
      }
      Normalidad_transformados_3<-Normalidad_transformados_3[,c(3,6,2,4,5,1)]
      colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
    }else{
      library(nortest)
      if(any(duplicated(Final_Resultados[, 1:j]))){
        while(j <= ncol(Final_Resultados)){
          Normalidad_transformados_3[k,] <- ad.test(Resultados_transformados_3[,j])
          k=k+1;
          j=j+1        
        }
        Normalidad_transformados_3<-Normalidad_transformados_3[,c(3,6,2,4,5,1)]
        colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
      }else{
        while(j <= ncol(Final_Resultados)){
          Normalidad_transformados_3[k,] <- ks.test(Resultados_transformados_3[,j], "pnorm",mean(Resultados_transformados_3[,j]), sd(Resultados_transformados_3[,j]))
          k=k+1;
          j=j+1
        }
        Normalidad_transformados_3<-Normalidad_transformados_3[,c(4,6,2,3,5,1)]
        colnames(Normalidad_transformados_3)<-c("Test", "Line", "p-value", "Significance level", "Result", "D")
        Normalidad_transformados_3[,1]="Kolmogorov-Smirnov test"
      }
    }
    Normalidad_transformados_3[,2]=c(colnames(Final_Resultados))
    Normalidad_transformados_3[,4]=p_valor_normalidad
    
    i=1
    for(i in 1:nrow(Normalidad_transformados_3)){
      if(Normalidad_transformados_3[i,3] > Normalidad_transformados_3[i,4]){
        Normalidad_transformados_3[i,5]= "Normal [1/x]"
      }else(
        Normalidad_transformados_3[i,5]= "Not normal"
      )
      i=i+1
    }
    
    Final_Normalidad<-Normalidad_transformados_3
    Resultados_usados<-Resultados_transformados_3
  }
  
  if(any(Final_Normalidad[,5] == "Not normal")){
    Final_Normalidad<-Normalidad[,1:5]
    Resultados_usados<-Final_Resultados
  }
}





############################## CEDASTICIDAD ################################

###Normales
if(all(Final_Normalidad[,5]=="Normal" |
       Final_Normalidad[,5]=="Normal [ln]"|
       Final_Normalidad[,5]=="Normal [sqrt]"|
       Final_Normalidad[,5]=="Normal [1/x]")){
  Cedasticidad<-as.data.frame(matrix(0,ncol = 6));
  Cedasticidad[,]<-bartlett.test(Resultados_tests$Acumulación_relativa~Resultados_tests$Line)
  Cedasticidad<-Cedasticidad[,c(5,3,1,4,6,2)]
  Cedasticidad[,1]="Bartlett"
  Cedasticidad[,3]=p_valor_cedasticidad
  colnames(Cedasticidad)<-c("Test", "p-value", "Significance level", "Result", "K-cuadrado", "Degrees of freedom")
  
  i=1
  for(i in 1:nrow(Cedasticidad)){
    if(as.double(Cedasticidad[i,2]) > as.double(Cedasticidad[i,3])){
      Cedasticidad[i,4]= "Homocedastic"
    }else{
      Cedasticidad[i,4]= "Heterocedastic"
    }
    i=i+1
  }
}else{
  
  ##NO NORMALES
  Cedasticidad<-as.data.frame(matrix(0,ncol = 6));
  Cedasticidad[,]<-leveneTest(Resultados_tests$Acumulación_relativa~Resultados_tests$Line)
  Cedasticidad<-Cedasticidad[,c(4,6,5,3,2,1)]
  Cedasticidad[,3]=p_valor_cedasticidad
  Cedasticidad[,1]="Levene"
  colnames(Cedasticidad)<-c("Test", "p-value", "Significance level", "Result", "F value", "Degrees of freedom")
  
  i=1
  for(i in 1:nrow(Cedasticidad)){
    if(as.double(Cedasticidad[i,2]) > as.double(Cedasticidad[i,3])){
      Cedasticidad[i,4]= "Homocedastic"
    }else{
      Cedasticidad[i,4]= "Heterocedastic"
    }
    i=i+1
  }
}

Final_Cedasticidad<-Cedasticidad[,c(1,6,2,3,4)]


#################### DIFERENCIAS NORMALES HOMOCEDÁSTICAS#############################
if((all(Final_Normalidad[,5]=="Normal" |
        Final_Normalidad[,5]=="Normal [ln]"|
        Final_Normalidad[,5]=="Normal [sqrt]"|
        Final_Normalidad[,5]=="Normal [1/x]")&
    Final_Cedasticidad[,5]=="Homocedastic")){
  
  ##ANOVA (hay diferentes)
  Diferencias<-as.data.frame(matrix(0,ncol=6,nrow=1))
  anova<-aov(Resultados_tests$Acumulación_relativa~Resultados_tests$Line)
  Diferencias<-cbind(Test="ANOVA",summary(anova)[[1]])
  Diferencias<-Diferencias[1,]
  Diferencias<-Diferencias[,c(1,6,3,4,5,2)]
  Diferencias[,3]=p_valor_diferencias
  colnames(Diferencias)<-c("Test", "p-value", "Significance level", "Result", "F value", "Degrees of freedom")
  rownames(Diferencias)<-NULL
  i=1
  for(i in 1:nrow(Diferencias)){
    if(as.double(Diferencias[i,2]) < as.double(Diferencias[i,3])){
      Diferencias[i,4]= "Different"
    }else{
      Diferencias[i,4]= "NO differences"
    }
    i=i+1
  }
  Final_Diferencias<-Diferencias[,c(1,6,2,3,4)]
  
  if(ncol(Resultados_usados)<=2){
    Resultados_tests<-Resultados_tests[1:34,]
    tstudent<-matrix(0,ncol = 9);
    tstudent<-tidy(t.test(Resultados_tests$Acumulación_relativa~Resultados_tests$Line, conf_level=(1-p_valor_posthoc), var.equal = TRUE))
    tstudent<-tstudent[,c(9,3,5,1,2,4,6,7,8)]
    colnames(tstudent)<-c("Test", "Compared lines","p-value", "Significance level", "Result", "Estadístico", "Parámetro", "Conf.low", "Conf.high")
    tstudent[1,4]=p_valor_posthoc
    tstudent[,2] = paste(colnames(Resultados_usados)[1],colnames(Resultados_usados)[2], sep = " - ")
    if(as.double(tstudent[1,3]) < as.double(tstudent[1,4])){
      tstudent[,5]=as.character("Significant")
    }else{
      tstudent[,5]<- as.character("ns")
    }
    
    PostHoc_1<-tstudent
    Final_PostHoc<-tstudent[,1:5]
  }
  
  
  if(ncol(Resultados_usados)>2){
    tukey<-TukeyHSD(anova)
    tukey2 <- matrix(unlist(tukey), ncol = 4, byrow = FALSE)
    tukey3<-as.data.frame(matrix(0,ncol=3,nrow=1));
    tukey4<-cbind(tukey2,tukey3)
    tukey4<-tukey4[,c(5,4,6,7,1,2,3)]
    colnames(tukey4)<-c("Test", "p-value", "Significance level", "Result", "diff", "lower", "upper")
    tukey4[,1]="Tukey Kramer"
    tukey4[,3]=p_valor_posthoc
    tukey5<-matrix( ncol = 1)
    tukey5<-rownames(tukey[[1]])
    tukey4<-cbind(tukey5,tukey4)
    tukey4<-tukey4[,c(2,1,3,4,5,6,7,8)]
    colnames(tukey4)<-c("Test", "Compared lines", "p-value", "Significance level", "Result", "diff", "lower", "upper")
    
    i=1
    for(i in 1:nrow(tukey4)){
      if(as.double(tukey4[i,3]) < as.double(tukey4[i,4])){
        tukey4[i,5]= "Significant"
      }else{
        tukey4[i,5]= "ns"
      }
      i=i+1
    }
    PostHoc_1<-tukey4
    
    
    Dunnett<-matrix(0,ncol=5, nrow=(ncol(Final_Resultados)-1))
    Dunnett<-as.matrix(DunnettTest(Resultados_tests$Acumulación_relativa~Resultados_tests$Line))[[1]]
    Dunnett<-cbind(Dunnett, "Significance level"=0, "Test"="Dunnett test", Resultado=0, "Compared lines"=rownames(Dunnett))
    rownames(Dunnett) <- NULL
    Dunnett<-Dunnett[,c(6,8,4,5,7,1,2,3)]
    Dunnett[,4]=p_valor_posthoc
    colnames(Dunnett)<-c("Test", "Compared lines", "p-value", "Significance level", "Result", "diff", "lower", "upper")
    
    i=1
    for(i in 1:nrow(Dunnett)){
      if(as.double(Dunnett[i,3]) < as.double(Dunnett[i,4])){
        Dunnett[i,5]= "Significant"
      }else{
        Dunnett[i,5]= "ns"
      }
      i=i+1
    }
    PostHoc_2<-Dunnett
    
    Final_PostHoc<-rbind(tukey4[,1:5], " ", Dunnett[,1:5])
  }
  
}


#################### DIFERENCIAS NORMALES HETEROCEDÁSTICAS#############################
if((all(Final_Normalidad[,5]=="Normal" |
        Final_Normalidad[,5]=="Normal [ln]"|
        Final_Normalidad[,5]=="Normal [sqrt]"|
        Final_Normalidad[,5]=="Normal [1/x]")&
    Final_Cedasticidad[,5]=="Heterocedastic")){
  
  ##ANOVA Welch (hay diferentes)
  Diferencias<-as.data.frame(matrix(0,ncol=6,nrow=1))
  Diferencias[,]<-oneway.test(Resultados_tests$Acumulación_relativa~Resultados_tests$Line)
  Diferencias<-Diferencias[,c(4,3,5,6,1,2)]
  Diferencias[,3]=p_valor_diferencias
  Diferencias[,1]="ANOVA Welch"
  colnames(Diferencias)<-c("Test", "p-value", "Significance level", "Result", "F value", "Degrees of freedom")
  i=1
  for(i in 1:nrow(Diferencias)){
    if(as.double(Diferencias[i,2]) < as.double(Diferencias[i,3])){
      Diferencias[i,4]= "Different"
    }else{
      Diferencias[i,4]= "NO differences"
    }
    i=i+1
  }
  Final_Diferencias<-Diferencias[,c(1,6,2,3,4)]
  
  if(ncol(Resultados_usados)==2){
    Resultados_tests<-Resultados_tests[1:34,]
    tstudent<-matrix(0,ncol = 6);
    tstudent<-tidy(t.test(Resultados_tests$Acumulación_relativa~Resultados_tests$Line, conf_level=(1-p_valor_posthoc), var.equal = FALSE))
    tstudent<-tstudent[,c(9,3,5,1,2,4,6,7,8)]
    colnames(tstudent)<-c("Test", "Compared lines","p-value", "Significance level", "Result", "Estadístico", "Parámetro", "Conf.low", "Conf.high")
    tstudent[1,4]=p_valor_posthoc
    tstudent[,2] = paste(colnames(Resultados_usados)[1],colnames(Resultados_usados)[2], sep = " - ")
    if(as.double(tstudent[1,3]) < as.double(tstudent[1,4])){
      tstudent[,5]=as.character("Significant")
    }else{
      tstudent[,5]<- as.character("ns")
    }
    
    PostHoc_1<-tstudent
    Final_PostHoc<-tstudent[,1:5]
  }
  
  if(ncol(Resultados_usados)>2){
    ##Games-Howell (cuáles son diferentes)
    GH<-games_howell_test(Resultados_tests, Acumulación_relativa~Line)
    GH2<-matrix(ncol=2)
    Games_Howell<-cbind(GH, GH2)
    Games_Howell$group1 <- paste(Games_Howell$group1, Games_Howell$group2, sep = " - ")
    Games_Howell<-Games_Howell[,c(1,2,7,9,10,8,4,5,6)]
    Games_Howell[,4]=p_valor_posthoc
    Games_Howell[,1]="Games Howell"
    colnames(Games_Howell)<-c("Test", "Compared lines", "p-value", "Significance level", "Result", "Significancia","Estimate", "Lower", "Upper")
    
    i=1
    for(i in 1:nrow(Games_Howell)){
      if(as.double(Games_Howell[i,3]) < as.double(Games_Howell[i,4])){
        Games_Howell[i,5]= "Significant"
      }else{
        Games_Howell[i,5]= "ns"
      }
      i=i+1
    }
    
    PostHoc_1<-Games_Howell
    
    Final_PostHoc<-Games_Howell[,1:5]
  }
}

#################### DIFERENCIAS NO NORMALES #############################
if(any(Final_Normalidad[,5]=="Not normal")){
  
  ## Kruskal Wallis (hay diferentes o no)
  Diferencias<-as.data.frame(matrix(0,ncol = 6));
  Diferencias[,]<-kruskal.test(Resultados_tests$Acumulación_relativa~Resultados_tests$Line)
  Diferencias<-Diferencias[,c(4,3,6,5,1,2)]
  Diferencias[,3]=p_valor_diferencias
  Diferencias[,1]="Kruskal Wallis"
  colnames(Diferencias)<-c("Test", "p-value", "Significance level", "Result", "F value", "Degrees of freedom")
  i=1
  for(i in 1:nrow(Diferencias)){
    if(as.double(Diferencias[i,2]) < as.double(Diferencias[i,3])){
      Diferencias[i,4]= "Different"
    }else{
      Diferencias[i,4]= "NO differences"
    }
    i=i+1
  }
  Final_Diferencias<-Diferencias[,c(1,6,2,3,4)]
  
  
  ##Whitney (cuáles son diferentes)
  combinaciones<-choose(ncol(Resultados_usados),2)
  Whitney<-as.data.frame(matrix(nrow=combinaciones, ncol=6))
  i=1
  x=1
  while(i <= ncol(Resultados_usados)){
    k=i+1
    while(k<=ncol(Resultados_usados)){
      Whitney[x,1]=wilcox.test(Resultados_usados[,i],Resultados_usados[,k])$p.value
      Whitney[x,2]=wilcox.test(Resultados_usados[,i],Resultados_usados[,k])$statistic
      Whitney[x,3]=colnames(Resultados_usados[i])
      Whitney[x,4]=colnames(Resultados_usados[k])
      k=k+1
      x=x+1
    }
    i=i+1
  }
  
  Whitney[,3] <- paste(Whitney[,3], Whitney[,4], sep = " - ")
  Whitney<-Whitney[,c(4,3,1,6,5,2)]
  Whitney[,4]=round(((p_valor_posthoc)/(nrow(Whitney))), digits=Decimales)
  Whitney[,1]="Mann-Whitney-Wilcoxon"
  
  i=1
  for(i in 1:nrow(Whitney)){
    if(as.double(Whitney[i,3]) < as.double(Whitney[i,4])){
      Whitney[i,5]= "Significant"
    }else{
      Whitney[i,5]= "ns"
    }
    i=i+1
  }
  ncol(Whitney)
  
  colnames(Whitney)<-c("Test", "Compared lines", "p-value", "Significance level", "Result", "W")
  
  PostHoc_1<-Whitney
  Final_PostHoc<-Whitney[,1:5]
  
}

if(Exportar_Análisis_Estadístico=="Yes"){
  i=1
  while(i<(nrow(Final_PostHoc)+1)){
    if(Final_PostHoc[i,5]=="Significant"){
      if(0.01<as.double(Final_PostHoc[i,3]) & as.double(Final_PostHoc[i,3])<=0.05){
        Final_PostHoc[i,5] <- "*"
      }
      if(0.001<as.double(Final_PostHoc[i,3]) & as.double(Final_PostHoc[i,3])<=0.01){
        Final_PostHoc[i,5] <- "**"
      }
      if(0.0001<as.double(Final_PostHoc[i,3]) & as.double(Final_PostHoc[i,3])<=0.001){
        Final_PostHoc[i,5] <- "***"
      }
      if(as.double(Final_PostHoc[i,3])<=0.0001){
        Final_PostHoc[i,5] <- "****"
      }
    }
    i=i+1
  }
  
  
  
  Tabla2<-Final_Normalidad[,1:5]
  Tabla2[,3] <- round(Tabla2[,3], Decimales)
  Tabla2[,3]<-signif(as.double(Tabla2[,3]), digits=Decimales)
  Tabla2[nrow(Tabla2) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla2[nrow(Tabla2) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla2[nrow(Tabla2) + 1,] <- c("Test","Degrees of freedom", "p-value", "Significance level", "Result")
  Final_Cedasticidad[1,3] <- round(Final_Cedasticidad[1,3], Decimales)
  Final_Cedasticidad[1,3]<-signif(as.double(Final_Cedasticidad[1,3]), digits=Decimales)
  Tabla2[nrow(Tabla2) + 1,] <- Final_Cedasticidad[1,]
  
  Tabla3<-Final_Diferencias
  Tabla3[,3]<-signif(as.double(Tabla3[,3]), digits=Decimales)
  Tabla3[nrow(Tabla3) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla3[nrow(Tabla3) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla3[nrow(Tabla3) + 1,] <- c("Test","Compared lines", "p-value", "Significance level", "Result")
  Final_PostHoc[1:nrow(PostHoc_1),3] <- round(as.double(Final_PostHoc[1:nrow(PostHoc_1),3]), Decimales)
  
  if(all(Final_Normalidad[,5]=="Normal" |
         Final_Normalidad[,5]=="Normal [ln]"|
         Final_Normalidad[,5]=="Normal [sqrt]"|
         Final_Normalidad[,5]=="Normal [1/x]")&
     Final_Cedasticidad[,5]=="Homocedastic"&
     ncol(Final_Resultados)>2){
    Final_PostHoc[nrow(PostHoc_1)+2:nrow(Final_PostHoc),3] <- round(as.double(Final_PostHoc[nrow(PostHoc_1)+2:nrow(Final_PostHoc),3]), Decimales)
  }
  
  Tabla3[nrow(Tabla3) + 1:nrow(Final_PostHoc),] <- Final_PostHoc[1:nrow(Final_PostHoc),]
  
  Tabla2[nrow(Tabla2) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla2[nrow(Tabla2) + 1,] <- c(" ", " ", " ", " ", " ")
  Tabla2[nrow(Tabla2) + 1,] <- c("Test","Degrees of freedom", "p-value", "Significance level", "Result")
  
  colnames(Tabla3)<-colnames(Tabla2)
  
  Tabla4<-rbind(Tabla2,Tabla3)
  
  wb<-loadWorkbook(Archivo_fuente)
  
  if(any(sheets(wb)== "Statistics")){
    nombres_hojas <- excel_sheets(Archivo_fuente)
    repes <- sum(grepl("Statistics", nombres_hojas, ignore.case = TRUE))
    pagina_estadistica <- paste("Statistics", repes)
    addWorksheet(wb, pagina_estadistica)
  }else{
    pagina_estadistica<- "Statistics"
    addWorksheet(wb, pagina_estadistica)
  }
  
  setColWidths(wb, pagina_estadistica, cols = 1:5, widths = "auto")
  titulos<-createStyle(fontName = "Microsoft Sans Serif", fontSize = 11, borderStyle= "double", borderColour="black", fgFill= "#DCDCDC", halign = "left")
  addStyle(wb, pagina_estadistica, style=titulos, rows=c(1, (nrow(Final_Normalidad)+4), (nrow(Tabla2)+1), (nrow(Tabla2)+nrow(Final_Diferencias)+4)), cols=1:ncol(Tabla2), gridExpand = TRUE)
  
  writeData(wb,pagina_estadistica, Tabla4)
  
  saveWorkbook(wb,Archivo_fuente,overwrite=TRUE)
}
}
)", split = "\n")[[1]]
            
            script_content <- gsub('Archivo_fuente\\s*<-\\s*"[^"]*"', 
                                 sprintf('Archivo_fuente<-"%s"', esc(result_name)), 
                                 script_content)
            crear_datos_script <- if (values$crear_datos == "qPCR Cq") "Create" else "Import"
            script_content <- gsub('Crear_Datos\\s*<-\\s*"[^"]*"', 
                                 sprintf('Crear_Datos<-"%s"', esc(crear_datos_script)), 
                                 script_content)
            script_content <- gsub('Crear_Gráfica\\s*<-\\s*"[^"]*"', 
                                 'Crear_Gráfica<-"Yes"', 
                                 script_content)
            script_content <- gsub('Crear_Análisis_Estadístico\\s*<-\\s*"[^"]*"', 
                                 'Crear_Análisis_Estadístico<-"Yes"', 
                                 script_content)
            script_content <- gsub('Hoja_fuente\\s*<-\\s*\\d+', 
                                 sprintf('Hoja_fuente<-%d', num(values$hoja_fuente, 2L)), 
                                 script_content)
            script_content <- gsub('Exportar_Resultados\\s*<-\\s*"[^"]*"', 
                                 sprintf('Exportar_Resultados<-"%s"', esc(values$exportar_resultados)), 
                                 script_content)
            script_content <- gsub('Exportar_Gráfica\\s*<-\\s*"[^"]*"', 
                                 'Exportar_Gráfica<-"Yes"', 
                                 script_content)
            script_content <- gsub('Pegar_Gráfica\\s*<-\\s*"[^"]*"', 
                                 sprintf('Pegar_Gráfica<-"%s"', esc(values$pegar_grafica)), 
                                 script_content)
            script_content <- gsub('Exportar_Análisis_Estadístico\\s*<-\\s*"[^"]*"', 
                                 sprintf('Exportar_Análisis_Estadístico<-"%s"', esc(values$exportar_analisis_estadistico)), 
                                 script_content)
            script_content <- gsub('Gráfico\\s*<-\\s*"[^"]*"', 
                                 sprintf('Gráfico<-"%s"', esc(values$grafico)), 
                                 script_content)
            script_content <- gsub('p_valor_normalidad\\s*<-\\s*[0-9.]+', 
                                 sprintf('p_valor_normalidad<-%.2f', num(values$p_valor_normalidad, 0.05)), 
                                 script_content)
            script_content <- gsub('p_valor_cedasticidad\\s*<-\\s*[0-9.]+', 
                                 sprintf('p_valor_cedasticidad<-%.2f', num(values$p_valor_cedasticidad, 0.05)), 
                                 script_content)
            script_content <- gsub('p_valor_diferencias\\s*<-\\s*[0-9.]+', 
                                 sprintf('p_valor_diferencias<-%.2f', num(values$p_valor_diferencias, 0.05)), 
                                 script_content)
            script_content <- gsub('p_valor_posthoc\\s*<-\\s*[0-9.]+', 
                                 sprintf('p_valor_posthoc<-%.2f', num(values$p_valor_posthoc, 0.05)), 
                                 script_content)
            script_content <- gsub('Num_muestras\\s*<-\\s*\\d+', 
                                 sprintf('Num_muestras<-%d', num(values$num_muestras, 50L)), 
                                 script_content)
            script_content <- gsub('Subgrupo\\s*<-\\s*"[^"]*"', 
                                 sprintf('Subgrupo<-"%s"', esc(values$subgrupo)), 
                                 script_content)
            script_content <- gsub('Nombre_Gráfica\\s*<-\\s*"[^"]*"', 
                                 sprintf('Nombre_Gráfica<-"%s"', esc(values$nombre_grafica)), 
                                 script_content)
            script_content <- gsub('Formato_grafica\\s*<-\\s*"[^"]*"', 
                                 sprintf('Formato_grafica<-"%s"', esc(values$formato_grafica)), 
                                 script_content)
            script_content <- gsub('Nombre_eje_y\\s*<-\\s*"[^"]*"', 
                                 sprintf('Nombre_eje_y<-"%s"', esc(values$nombre_eje_y)), 
                                 script_content)
            script_content <- gsub('Nombre_eje_x\\s*<-\\s*"[^"]*"', 
                                 sprintf('Nombre_eje_x<-"%s"', esc(values$nombre_eje_x)), 
                                 script_content)
            script_content <- gsub('Color_relleno\\s*<-\\s*"[^"]*"', 
                                 sprintf('Color_relleno<-"%s"', esc(values$color_relleno)), 
                                 script_content)
            script_content <- gsub('Color_bordes\\s*<-\\s*"[^"]*"', 
                                 sprintf('Color_bordes<-"%s"', esc(values$color_bordes)), 
                                 script_content)
            script_content <- gsub('Color_puntos\\s*<-\\s*"[^"]*"', 
                                 sprintf('Color_puntos<-"%s"', esc(values$color_puntos)), 
                                 script_content)
            script_content <- gsub('Color_barra_error\\s*<-\\s*"[^"]*"', 
                                 sprintf('Color_barra_error<-"%s"', esc(values$color_barra_error)), 
                                 script_content)
            script_content <- gsub('Color_fondo\\s*<-\\s*"[^"]*"', 
                                 sprintf('Color_fondo<-"%s"', esc(values$color_fondo)), 
                                 script_content)
            script_content <- gsub('Tamaño_puntos\\s*<-\\s*[0-9.]+', 
                                 sprintf('Tamaño_puntos<-%.2f', num(values$tamano_puntos, 1.5)), 
                                 script_content)
            script_content <- gsub('Forma_puntos\\s*<-\\s*\\d+', 
                                 sprintf('Forma_puntos<-%d', num(values$forma_puntos, 16L)), 
                                 script_content)
            script_content <- gsub('Movimiento_lateral\\s*<-\\s*[0-9.]+', 
                                 sprintf('Movimiento_lateral<-%.2f', num(values$movimiento_lateral, 0.1)), 
                                 script_content)
            script_content <- gsub('Ancho_cajas\\s*<-\\s*[0-9.]+', 
                                 sprintf('Ancho_cajas<-%.2f', num(values$ancho_cajas, 0.5)), 
                                 script_content)
            script_content <- gsub('Ancho_barra_error\\s*<-\\s*[0-9.]+', 
                                 sprintf('Ancho_barra_error<-%.2f', num(values$ancho_barra_error, 0.15)), 
                                 script_content)
            script_content <- gsub('Grosor_borde\\s*<-\\s*[0-9.]+', 
                                 sprintf('Grosor_borde<-%.2f', num(values$grosor_borde, 0.5)), 
                                 script_content)
            script_content <- gsub('Grosor_mediana\\s*<-\\s*[0-9.]+', 
                                 sprintf('Grosor_mediana<-%.2f', num(values$grosor_mediana, 1)), 
                                 script_content)
            script_content <- gsub('Inicio\\s*<-\\s*[0-9.]+', 
                                 sprintf('Inicio<-%.2f', num(values$inicio, 0)), 
                                 script_content)
            script_content <- gsub('Marcas\\s*<-\\s*[0-9.]+', 
                                 sprintf('Marcas<-%.2f', num(values$marcas, 1)), 
                                 script_content)
            # Evitar viewport no finito: asegurar resolución y dimensiones válidas
            resolucion_safe <- if (is.numeric(values$resolucion) && is.finite(values$resolucion) && values$resolucion > 0) as.integer(values$resolucion) else 400L
            ancho_safe <- if (is.numeric(values$ancho) && is.finite(values$ancho) && values$ancho > 0) as.integer(values$ancho) else 4000L
            alto_safe <- if (is.numeric(values$alto) && is.finite(values$alto) && values$alto > 0) as.integer(values$alto) else 3500L
            script_content <- gsub('Resolución\\s*<-\\s*\\d+', 
                                 sprintf('Resolución<-%d', resolucion_safe), 
                                 script_content)
            script_content <- gsub('Unidades\\s*<-\\s*"[^"]*"', 
                                 sprintf('Unidades<-"%s"', esc(values$unidades)), 
                                 script_content)
            script_content <- gsub('Ancho\\s*<-\\s*\\d+', 
                                 sprintf('Ancho<-%d', ancho_safe), 
                                 script_content)
            script_content <- gsub('Alto\\s*<-\\s*\\d+', 
                                 sprintf('Alto<-%d', alto_safe), 
                                 script_content)
            script_content <- gsub('Réplicas\\s*<-\\s*\\d+', 
                                 sprintf('Réplicas<-%d', num(values$replicas, 3L)), 
                                 script_content)
            script_content <- gsub('Diferencia_Ct\\s*<-\\s*[0-9.]+', 
                                 sprintf('Diferencia_Ct<-%.1f', num(values$diferencia_ct, 1)), 
                                 script_content)
            script_content <- gsub('Mock\\s*<-\\s*\\d+', 
                                 sprintf('Mock<-%d', num(values$mock, 0L)), 
                                 script_content)
            script_content <- gsub('Decimales\\s*<-\\s*\\d+', 
                                 sprintf('Decimales<-%d', num(values$decimales, 4L)), 
                                 script_content)
            script_content <- gsub('Comas\\s*<-\\s*"[^"]*"', 
                                 sprintf('Comas<-"%s"', esc(values$comas)), 
                                 script_content)
            # Nuevas variables - Modificando el patrón de búsqueda para que sea más preciso
            script_content <- gsub('Tamaño_titulo_ejes\\s*<-\\s*[0-9.]+', 
                                 sprintf('Tamaño_titulo_ejes<-%.1f', num(values$tamano_titulo_ejes, 17)), 
                                 script_content)
            script_content <- gsub('Tamaño_texto_ejes\\s*<-\\s*[0-9.]+', 
                                 sprintf('Tamaño_texto_ejes<-%.1f', num(values$tamano_texto_ejes, 14)), 
                                 script_content)
            script_content <- gsub('Distancia_corte_puntos\\s*<-\\s*[0-9.]+', 
                                 sprintf('Distancia_corte_puntos<-%.1f', num(values$distancia_corte_puntos, 0.5)), 
                                 script_content)
            
            # Agregar debug para ver los valores
            cat("Valores actuales:\n")
            cat("Tamaño título ejes:", values$tamano_titulo_ejes, "\n")
            cat("Tamaño texto ejes:", values$tamano_texto_ejes, "\n")
            cat("Distancia corte puntos:", values$distancia_corte_puntos, "\n")
            
            if (isTRUE(values$subgrupo == "Yes") && !is.null(column_data())) {
                selected_cols <- c()
                for (i in seq_len(ncol(column_data()))) {
                    if (isTRUE(input[[paste0("col_", i)]])) {
                        selected_cols <- c(selected_cols, i)
                    }
                }
                
                if (length(selected_cols) > 0) {
                    columnas_str <- paste(selected_cols, collapse=",")
                    script_content <- gsub('Columnas\\s*<-\\s*c\\([^)]+\\)', 
                                         sprintf('Columnas<-c(%s)', columnas_str), 
                                         script_content)
                }
            }
            
############ Crear archivo temporal con modificaciones
            temp_script <- tempfile(fileext = ".R")
            writeLines(script_content, temp_script)
            
############ Ejecutar script temporal
            source(temp_script)
            
############ Eliminar archivo temporal
            unlink(temp_script)
            
############ Guardar valores usados
            last_analysis_values(values)
            graph_file <- paste0(values$nombre_grafica, ".", values$formato_grafica)
            graph_path <- file.path(getwd(), graph_file)
            if (file.exists(graph_path)) {
                last_graph_path(graph_path)
            } else {
                last_graph_path(NULL)
            }
            result_file <- file.path(getwd(), result_name)
            if (file.exists(result_file)) last_result_path(result_file) else last_result_path(NULL)
            
            "Analysis completed succesfully!"
        }, error = function(e) {
            paste("Warning:", e$message)
        })
    }
    
    # Botón "Start Analysis": lanza el análisis completo con los valores actuales
    observeEvent(input$startButton, {
        # Si ya hay un análisis en curso, solo informamos al usuario
        if (analysis_running()) {
            output$status <- renderText("Analysis is already in progress. Please wait for it to finish.")
            return()
        }
        
        # Validar que se haya seleccionado un archivo
        if (is.null(input$archivo_fuente)) {
            output$status <- renderText("Please select an Excel file (.xlsx).")
            return()
        }
        
        # Marcar que el análisis está corriendo y mostrar mensaje claro
        analysis_running(TRUE)
        output$status <- renderText("Running analysis, please wait...")
        
        # Obtener los valores actuales almacenados
        values <- stored_values()
        
        # Ejecutar el análisis mostrando un indicador de progreso
        result <- NULL
        withProgress(message = "Running analysis...", value = 0, {
            result <- run_analysis(values)
        })
        
        # Marcar fin del análisis y mostrar el resultado
        analysis_running(FALSE)
        output$status <- renderText(result)
    })
}

########### Ejecutar la aplicación Shiny
shinyApp(ui = ui, server = server) 