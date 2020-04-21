# BecarioDigital
Becario Digital - Querida! Acabe con el trabajo del becario.

El BecarioDigital es un script sencillo para modificar y generar documentos de word,
ademas de contar con una funcion para enviar el documento generado a impresion con solo dar un click.

En el hospital, es tu aliado para llenar la papeleria del paciente de una forma elegante y eficiente.
Lo amaras cuando la carga de trabajo sea excesiva.


Funciona de esta forma:
-Solicita al usuario informacion para generar los consentimientos medicos.
-Abre los documento de word que se encuentran en la carpeta formatos.
-Rellena la informacion necesaria en cada uno y los guarda en un documento de word nuevo.
-Envia los documentos de word a impresion, se pueden seleccionar los consentimientos por separado o en grupo.


Utilizo los siguientes modulos:

OS: Para acceder a las herramientas del sistema

PyForms: modulo de python para generar interfaces GUI, Webapp y Terminal

Pywin32: lo utilizo para acceder a la funcion de enviar documento a impresion (es necesario que el equipo tenga Word instalado para que funcione)

Python-Docx: Modulo para abrir y editar documentos de Word.

