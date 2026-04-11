---

# 📊 Cotizador Automatizado de Prendas en Excel

## 🧾 Descripción

Herramienta desarrollada en Excel para generar cotizaciones automatizadas de prendas personalizadas con impresión DTF. Permite calcular precios de forma rápida y precisa según múltiples variables, reduciendo errores y mejorando la eficiencia en la atención al cliente.

---

## 🎯 Problema

En negocios de prendas personalizadas, la generación de proformas suele ser lenta y propensa a errores, especialmente cuando los precios dependen de múltiples factores como tipo de prenda, tela, talla o volumen.
Esto puede generar retrasos en la atención y, en muchos casos, pérdida de ventas por falta de respuesta rápida al cliente.

---

## 💡 Solución

Se desarrolló un cotizador automatizado en Excel que permite seleccionar características específicas de cada prenda y calcular el precio de manera inmediata.
El sistema aplica reglas de negocio complejas de forma automática, asegurando precisión y reduciendo el tiempo de respuesta en la generación de cotizaciones.

---

## ⚙️ Funcionalidades

* Selección de prendas: polos, poleras, boxers y bodys
* Validación dinámica de datos según relaciones entre características
* Filtrado automático de opciones (tela, talla, estilo, color, etc.)
* Cálculo automático de precios según:
  * tipo de prenda
  * volumen (unidad, media docena, docena o más)
  * campaña activa
  * características específicas

* Aplicación de reglas de negocio diferenciadas por prenda
* Generación de resumen de cotización:
  * descripción del pedido
  * precio unitario
  * adicionales o descuentos
  * importe total

* Uso de tablas estructuradas para organización de datos
* Manejo de relaciones entre variables (ej. tela-prenda, estilo-talla)

---

## 🛠 Tecnologías

* Microsoft Excel
* Funciones avanzadas:

  * `FILTRAR` → para mostrar opciones dinámicas según contexto
  * `COINCIDIR` → para validar relaciones entre datos
  * `ESNUMERO` → para validaciones lógicas
  * `BUSCARX` → para búsqueda de precios en tablas
  * `SI` → para lógica condicional en el cálculo

---

## 📷 Vista previa

El flujo de uso del cotizador es el siguiente:

1. Se selecciona el tipo de prenda
2. Las opciones disponibles (tela, talla, estilo, color, etc.) se actualizan automáticamente
3. Se ingresa la cantidad requerida
4. El sistema calcula el precio unitario y total en tiempo real
5. Se genera una descripción completa de la cotización

---

## 🔗 Acceso al proyecto

* 📥 Descargar archivo: disponible en este repositorio
* 🌐 Ver online: (https://1drv.ms/x/c/5e29aba03f0ba95f/IQCLVYHMSFybQ7IaEl8d6CS2AWhgz9Rka0Ekjh2eD7ZoQLQ?e=insAWr)

---

## ▶️ Cómo usarlo

1. Seleccionar el tipo de prenda
2. Elegir las características disponibles (tela, talla, estilo, etc.)
3. Ingresar la cantidad de unidades
4. Revisar el precio calculado automáticamente
5. Visualizar el resumen de la cotización generado por el sistema

---

## 📚 Aprendizaje

Este proyecto demuestra la aplicación de conceptos clave en análisis y estructuración de datos en Excel:

* Validación de datos con referencias dinámicas
* Uso de fórmulas matriciales para filtrado inteligente
* Implementación de lógica de negocio mediante funciones condicionales
* Modelado de relaciones entre entidades dentro de Excel (simulando estructura relacional)
* Automatización de procesos manuales mediante herramientas de hoja de cálculo

