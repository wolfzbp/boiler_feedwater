Sí. Revisé el Excel completo y el problema es suficientemente estructural como para tratarlo como Master Data Management, no simplemente como limpieza de strings.

Te preparé además una primera versión funcional del pipeline y un Excel de auditoría, sin destruir los valores originales:

Descargar script ⁠￼standardizacion_textil.py⁠￼

Descargar Excel con diagnóstico y columnas estandarizadas⁠￼

Descargar perfil técnico JSON⁠￼

1. Diagnóstico general

El archivo tiene 1.042 registros.

Color

Problema / categoría	Resultado
Valores vacíos	90 (8,6%)
Valores únicos raw	225
Valores únicos tras normalización básica	221
Filas con concatenaciones repetidas	246 (23,6%)
Únicos tras colapsar repeticiones	191
Registros con código FHI tipo 19-1559	418 (40,1%)
Códigos FHI distintos	54
Código interno de 6 dígitos 77xxxx	270
Código interno de 8 dígitos	82
Pantone PMS explícito	13
Texto/otro sin estándar identificable	169
TCX explícitos	225
TPG explícitos	99
TPX explícitos	6

El primer problema grave es este tipo de dato:

191724 TCX CABERNET, 191724 TCX CABERNET, 191724 TCX CABERNET, 191724 TCX CABERNET

Eso no representa cuatro colores. Es el mismo maestro concatenado cuatro veces.

Detecté 246 filas con ese fenómeno.

Por eso pasar fuzzy matching directamente sobre Color produciría resultados malos.

⸻

2. Color debe convertirse en una entidad, no seguir siendo una columna de texto

Por ejemplo:

191559 TPG SCARLET SAGE

no debería almacenarse conceptualmente como una sola cadena.

Debería quedar así:

Campo	Valor
color_system	PANTONE_FHI
pantone_base	19-1559
pantone_source_suffix	TPG
pantone_canonical	19-1559 TCX
official_name	Scarlet Sage
internal_code	null
raw_value	191559 TPG SCARLET SAGE
canonical_color_id	PANTONE_FHI|19-1559|TCX

La columna original jamás se elimina. Se conserva como lineage.

TCX vs TPG

Aquí hay una decisión importante.

Pantone dice que para Fashion, Home + Interiors:

* TCX representa el color sobre algodón y es el sistema apropiado para textiles/apparel.
* TPG se utiliza principalmente para superficies duras, recubrimientos, cuero, etc.
* Los colores TPG tienen correspondencia con el sistema TCX.  

Por eso, para una planta textil, recomiendo que el estándar interno sea:

PANTONE FHI — TCX

pero conservando:

Pantone_Sufijo_Origen = TCX / TPG / TPX / NULL

Nunca borraría esa información.

⸻

3. Encontré errores reales de nombres Pantone

Esto confirma que el código debe dominar al nombre.

En tu Excel aparece:

19-4024 → DRESS BLUE y DRESS BLUES

Pantone oficial lo identifica como:

19-4024 TCX Dress Blues.  

También:

19-4151

aparece como:

SKYDRIVER

y

SKYDIVER

Pantone oficial es:

19-4151 Skydiver.  

Otros ejemplos:

19-1757 → Barbados Cherry.  

19-3934 → Oceana.  

Por eso sería peligroso hacer algo como:

nombre_oficial = df.groupby("codigo")["nombre"].agg(moda)

La moda del ERP puede estar equivocada.

⸻

4. Otro conflicto muy interesante: mismo nombre, distinto código

Encontré, entre otros:

Nombre extraído	Situación
NEGRO	sin FHI + 19-4007
CYBER YELLOW	14-0760 + registros sin código
ASPHALT	19-0201 + registros sin código
HIGH RISK RED	18-1763 + registros sin código
SURF THE WEB	19-3952 + registros sin código
FLAME SCARLET	18-1662 + registros sin código
AVENTURINE	19-5421 + registros sin código
DRESS BLUE	19-4024 + registros sin código
PAISLEY PURPLE	17-3130 y 17-3730

Los primeros casos probablemente permiten imputar el código mediante una tabla de alias.

Pero PAISLEY PURPLE es diferente: un mismo nombre está asociado con dos códigos diferentes.

Ahí el algoritmo no debería elegir automáticamente.

Debe marcar:

MANUAL_REVIEW_REQUIRED

⸻

5. Hay incluso valores que probablemente no son colores

Uno bastante evidente es:

LAVADO_MAQUINA

Aparece 53 veces en Color.

Lo estoy clasificando como:

NON_COLOR_VALUE

Eso es importante porque evidencia un problema de dominio de datos, no de ortografía.

La columna permite registrar cosas que conceptualmente no pertenecen a la dimensión Color.

⸻

6. Tela tiene un problema parecido, pero más interesante

Hay:

Métrica	Resultado
Tela vacía	123 (11,8%)
Valores raw únicos	217
Código de tela detectado	890 registros (~85%)
Familias de código detectadas	16
Códigos de tela distintos detectados	~106
Códigos encontrados embebidos dentro del nombre	45

Las familias de código más frecuentes son:

JAQ, SP, JL, DJAQ, BO, SPL, ES, SJ, RI, PO, IG, CW, IN, SL, DP, EJAQ.

Pero por ahora las llamaría Familia_Codigo, no todavía Familia_Tela.

No sabemos aún si JAQ, por ejemplo, representa formalmente una familia tecnológica, comercial, de construcción, proveedor, origen u otra taxonomía interna.

Eso hay que validarlo con negocio.

⸻

7. Una Tela actualmente contiene demasiadas dimensiones

Por ejemplo:

BO36446 OTTOMAN CATIONICO RAMADO C2

realmente contiene:

Dimensión	Valor
Código	BO-36446
Familia código	BO
Nombre comercial	OTTOMAN CATIONICO
Acabado	RAMADO
Ruta/proceso	C2

Y hay casos todavía más complejos:

TL.EX.TEXFINA BO31810 BLISTER SCUBA 3D COTIZACION

Aquí potencialmente tenemos:

Proveedor/origen + código + nombre comercial + variante + condición comercial.

Otro:

JL-1182 RAMADO 150D144F 91%PES 9%ELAS COTIZACION

contiene incluso:

Código + acabado + título/hilado + composición + cotización.

Por tanto Tela tampoco debería seguir siendo una única dimensión textual.

⸻

8. Encontré conflictos de composición que son mucho más importantes que la ortografía

Por ejemplo SPL-1343 aparece como:

97% PES / 3% LYC

pero también:

95% PES / 5% LYC

Eso es crítico.

Puede significar una de dos cosas:

A. Error de master data.

o

B. El mismo código está siendo utilizado para dos productos técnicamente diferentes.

Otro caso es JAQ-1550, donde aparecen variantes de composición asociadas al mismo código.

Esto hay que resolver antes de cualquier dashboard sofisticado.

De lo contrario un gráfico podría decir:

SPL-1343 tuvo X kg y Y% de rechazo

cuando en realidad estamos agregando dos construcciones diferentes bajo el mismo identificador.

⸻

9. Estructura de Master Data que recomiendo

Para Color:

Campo

color_id

color_system

pantone_base

pantone_suffix_canonical

official_name

official_family

internal_color_code

active_flag

source

validated_date

y aparte:

map_color_alias

alias_raw	color_id	confidence	mapping_method	approved_by

Así:

SKYDRIVER 194151

SKYDIVER 19-4151 TPG

194151 TCX SKYDRIVER

podrían terminar todos relacionados con:

PANTONE_FHI|19-4151|TCX

sin modificar el dato histórico.

Para Tela haría exactamente lo mismo:

dim_tela

Campo

tela_id

codigo

familia_codigo

familia_producto

subfamilia

nombre_comercial

construccion

composicion

acabado

proveedor

active_flag

y:

map_tela_alias

para relacionar todas las variantes históricas.

⸻

10. Orden de resolución que usaría el algoritmo

Aquí sí evitaría empezar con IA o embeddings.

El matching debería seguir esta jerarquía:

1. Código exacto aprobado
2. Alias exacto previamente aprobado
3. Código normalizado (191559 → 19-1559)
4. Nombre exacto con una sola correspondencia posible
5. Código interno ↔ maestro
6. Fuzzy matching de nombres
7. Fuzzy + atributos técnicos
8. Revisión humana

Y cada asignación debería producir un:

match_confidence

por ejemplo:

1.00 DIRECT_MASTER_CODE

0.98 APPROVED_ALIAS

0.90 UNIQUE_NAME_MATCH

0.76 FUZZY_CANDIDATE

0.00 MANUAL_REVIEW

Yo no autoaprobaría fuzzy matches ambiguos.

⸻

11. El script que te preparé ya trabaja así

En el Excel resultante no sobrescribí Tela ni Color.

Añadí columnas derivadas como:

Color_Normalizado
Color_Sistema_Detectado
Pantone_FHI_Base
Pantone_Sufijo_Origen
Pantone_TCX_Canonico
Color_Codigo_Interno
Pantone_PMS
Color_Nombre_Extraido
Color_Nombre_Oficial
Color_Familia_Oficial
Color_Clave_Canonica
Color_Flags
Tela_Normalizada
Tela_Codigo
Tela_Familia_Codigo
Tela_Nombre_Candidato
Tela_Acabado
Tela_Ruta
Tela_Codigo_Embebido
Tela_Flags

Y creé estas hojas:

DQ_Resumen

DQ_Master_Color

DQ_Master_Tela

DQ_Issues

DQ_Conflictos_Color

DQ_Conflictos_Tela

La hoja DQ_Conflictos_Color me parece especialmente útil: ahí puedes ver directamente cosas como:

19-4151 → SKYDRIVER / SKYDIVER

sin revisar las 1.042 filas.

El Excel generó actualmente 21 conflictos agregados de color y 64 conflictos agregados de tela para revisión.

⸻

12. Algo importante sobre Pantone

Intencionalmente dejé:

Color_Nombre_Oficial

vacío en la mayoría del workbook.

No quiero poblar automáticamente una supuesta base Pantone haciendo scraping o tomando listas no verificadas.

El script acepta una tabla maestra autorizada:

pantone_base,tcx_name,color_family
19-4151,Skydiver,Purple-Blue
...

y entonces puede completar automáticamente el maestro.

Pantone además advierte que las representaciones digitales no sustituyen al estándar físico para evaluación visual.  

Por tanto para gobernanza industrial lo correcto sería adquirir/utilizar la referencia Pantone que tenga la empresa o Pantone Connect como source of truth, y no una tabla random de Internet. Pantone describe Connect como su plataforma digital con acceso a sus bibliotecas de color.  

⸻

13. Arquitectura que propondría

Para este proyecto usaría:

RAW / Bronze

Excel/ERP original e inmutable.

↓

STANDARDIZED / Silver

Parsing, normalización, extracción de códigos, composiciones, acabados y aliases.

↓

MASTER / Gold

Solo colores y telas aprobados.

↓

FACT Producción/Calidad

Los registros productivos ya apuntan a color_id y tela_id.

Eso nos permitiría finalmente analizar calidad de verdad sin que:

DRESS BLUE

DRESS BLUES

194024 TPG DRESS BLUE

sean tres categorías diferentes.

⸻

14. Dashboard que construiría después

Mi combinación favorita aquí sería:

Streamlit + Altair + Polars/DuckDB.

No empezaría todavía con un dashboard tradicional de kg producidos. Primero construiría un Data Stewardship Dashboard.

Vista	Qué mostraría
Data Quality Overview	completeness, validity, uniqueness, consistency, master coverage
Color Master	raw → canonical → Pantone → confidence
Color Conflicts	mismo código/diferentes nombres, mismo nombre/diferentes códigos
Tela Master	familias, códigos, nombres, composición, acabado
Tela Conflicts	mismo SKU con múltiples composiciones/descripciones
Review Queue	aceptar/rechazar mappings sugeridos
Lineage	valor original → regla → valor estandarizado
Quality Analytics	después: defectos, kg, FPY, rechazo por Color/Tela

Y aquí Altair sería especialmente fuerte para hacer un heatmap Familia Tela × Color Family, histogramas de calidad de master data, gráficos de Pareto de conflictos y navegación interactiva.

Lo que haría como siguiente paso

Antes de entrar al dashboard de producción, yo haría una Fase 2 de este mismo análisis: construir automáticamente el verdadero Master_Color y Master_Tela, incluyendo composición, hilado/construcción, proveedor, familia y subfamilia, y clasificar cada uno de los 225 colores y 217 valores de tela en:

AUTO_APPROVED

HIGH_CONFIDENCE

REVIEW_REQUIRED

INVALID_DOMAIN

Eso nos dejaría una base limpia y gobernada sobre la que sí vale la pena construir analítica avanzada.