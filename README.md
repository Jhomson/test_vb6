# test_vb6

Aplicativo de consulta de datos a partir de una consulta a un servicio, servidor o aplicativo de terceros usando Visual Basic para aplicaciones
Para hacer uso del programa, debe descragar los archivos y ejecutar "CRUD.vbp" con el editor de Visual Basic 6.0 profesional.

Para ejecutar la aplicación correctamente, se deben tener marcadas las siguientes referencias

Visula Basic For Applications
Visual Basic runtime objects and procedures
Visual Basic objects and procedures
OLE Automation
Microsoft Data Binding Collection
Microsoft Internet Controls
Microsoft Activex Data Objects 2.5 Library
Microsoft Scripting Runtime
Microsoft WinHTTP Services, version 5.1
Microsoft XML, V6.0

Tambien debe agregar al proyecto, en caso de que no se hayan importado automaticamente, los siguientes archivos

cJSONScript.cls
cStringBuilder.cls
JSON.bas

Una vez configurado de esta forma, se puede ejecutar sin inconveniente, seleccione el menu "PokeAPI" para que aparesca el formulario correspondiente.
Para realizar la busqueda, debe indicar en el cuadro de texto, cuantos items desea ver por pagina, en caso de no indicar ningun valor, se mostraran
10 items por defecto.

*Nota: Faltante función para pasar de página, se deben mostrar el total de items.
