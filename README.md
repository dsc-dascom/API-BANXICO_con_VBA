

# **API-BANXICO**

Este proyecto está diseñado para apoyar en la consulta recurrente de información económica proveniente del Banco de México (Banxico). Para lograr esto, se usa la API que Banxico proporciona a los usuarios. Este programa está hecho en VBA porque pretendo que todo el proceso de recolección, tratamiento, análisis y presentación de datos se realice dentro de un ambiente conocido por la mayoría de las personas: Excel.

---

<p align="center"> <img src="https://github.com/user-attachments/assets/93387aa8-816f-49dc-b03c-9f83703210c9" alt="Logo Banxico" width="600"> </p>


---

## **Documentación**   :open_file_folder:
- Recomiendo revisar la página de Banxico que explica aspectos relevantes de su API, como los parámetros de la consulta: https://www.banxico.org.mx/SieAPIRest/service/v1/

- De igual forma, en el siguiente link se puede solicitar el **Token**:
  https://www.banxico.org.mx/SieAPIRest/service/v1/token

- Si quieren visitar el portal de información de Banxico pueden acceder dando clic aquí --> [SIE Banxico]( https://www.banxico.org.mx/SieInternet/). 

- En el código se utilizan algunos objetos dentro de MSXML2 para trabajar con los datos en formato XML. Para más información dejo a su disposición las siguientes páginas: [ServerXMLHTTP](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)), [DOMDocument](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757828(v=vs.85)), y en especial [.setRequestHeader]( https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms764715(v=vs.85)).
  
---


## **Uso de API_BANXICO.bas**    :package:

En el archivo **API_BANXICO.bas** se puede encontrar el Módulo de VBA que contiene el código para utilizar la API de Banxico. 

Este código se puede dividir en tres secciones. 
- En la primera se declara el **Token** como una constante.

      Private Const banxico_token As String = "[Token]"
<br>

- En la segunda parte se construye una función que realiza el procedimiento de consulta y almacenamiento de información. Esta función depende de una única variable (_"serie"_), la cual se define en la subrutina.

      Function API_BANXICO(serie)
<br>

Una particularidad de la API de BANXICO es que se puede seleccionar un rango de tiempo para hacer la consulta de datos. En el código declaré que las fechas las obtenga de las celdas "D6" (fecha inicial) y "D7" (fecha final).
      
    fecha_inicio = Format(Range("D6").Value, "YYYY-MM-DD")
    fecha_fin = Format(Range("D7").Value, "YYYY-MM-DD")

Un ejemplo de la vista en Excel sería la siguiente:

<p align="center"> <img src="https://github.com/user-attachments/assets/a00dba7b-5322-4b8e-b110-f4cd95b1579d" alt="ejemplo1" width="650"> </p>


Otra opción es dejar fija la fecha inicial de consulta escribiendo directamente sobre el código:
    
    fecha_inicio = "2020-01-01"    'por ejemplo

Para la fecha final de consulta se puede usar la función "Date" sobre el código de VBA o la función "=Hoy( )" sobre la casilla de excel para obtener el dato disponible más reciente:
    
    fecha_fin = Date    'por ejemplo    
    
Pero esta decisión depende exclusivamente de las preferencias del usuario.  

<br>

- En la última sección se crea una subrutina donde se declara a la variable _"serie"_ y se utiliza como insumo en la función. Posteriormente se imprimen los datos en Excel.  

      Sub peso_dolar()
<br>

La lógica del código permite aumentar el número de consultas al declarar una lista con claves que junto con un ciclo _for_ permitirá realizar varias consultas de información.

También es posible crear varias subrutinas que se pueden ejecutar en distintas hojas de Excel. Se puede insertar un botón con la macro asignada y realizar las consultas repetidas veces.

<p align="center"> <img src="https://github.com/user-attachments/assets/ac566bad-3aab-438f-9b2a-001064d44b63" alt="ejemplo1" width="400"> </p>
<br>


<p align="center"> <img src="https://github.com/user-attachments/assets/6aceabad-4a4d-47d2-a3fb-e992643068f5" alt="ejemplo2" width="400"> </p>
<br>


Todo depende de las necesidades del proyecto o de las especificaciones de los usuarios para hacerlo más fácil de manejar. 

El resultado esperado es la serie de datos en el espacio seleccionado:
<p align="center"> <img src="https://github.com/user-attachments/assets/2a37ae6a-cf64-42f8-8cc4-5dca7138d763" alt="ejemplo2" width="1200"> </p>
*Nota: algunas filas se ocultaron para hacer visible el dato inicial y el dato final de la consulta.
<br>

---

## Objetivos de este proyecto   :seedling:

- Optimizar y automatizar la consulta de información que está disponible en las bases de datos del Banco de México.
  
- Facilitar la comprensión de métodos y objetos en VBA para el uso de la API de banxico.
  
- Construir un código que sirva como base para un proyecto que involucre analizar las condiciones económicas de México con información disponible del Banco de México.
