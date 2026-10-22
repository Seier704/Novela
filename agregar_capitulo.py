# -*- coding: utf-8 -*-
"""
===============================================================================
GESTOR UNIFICADO DE CAPÍTULOS - "EL CAMINO DEL REENCARNADO"
===============================================================================
Este archivo es la plantilla GENERAL y ÚNICA para agregar o sobrescribir capítulos.
Para agregar un nuevo capítulo, simplemente edite el diccionario `CAPITULO` con
el número, título y párrafos correspondientes y ejecute este script:
    python agregar_capitulo.py

El script automáticamente:
1. Genera el documento Word individual: capitulos/XX_Capitulo_XX.docx
2. Recompila el libro maestro completo: compilado/El_Camino_del_Reencarnado.docx
===============================================================================
"""

import os
import sys
import glob
import re
import docx

# Definición de rutas base
NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"c:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

for d in [NOVELA_DIR, PROYECTO_DIR]:
    if os.path.exists(d) and d not in sys.path:
        sys.path.insert(0, d)

from generar_libro import create_novel_document, create_single_chapter_document

# ===============================================================================
# DEFINICIÓN DEL CAPÍTULO A AGREGAR / ACTUALIZAR
# (Modifique esta sección para cada nuevo capítulo)
# ===============================================================================
CAPITULO = {
    "num": 71,
    "header": "Capítulo 71",
    "title": "Raíces de Insurrección: Sombras de Corrupción y el Juramento de los Pueblos",
    "paragraphs": [
        "La niebla matinal cubría como un sudario gris los campos desolados entre Sauceda y Fresnedas, el segundo pueblo del feudo de Morvath.",
        "Avanzábamos en formación de vanguardia rápida. Kaelan se deslizaba entre los matorrales con sus pasos felinos inaudibles, mientras Akame se movía por las copas de los árboles como una sombra espectral. Detrás marchábamos Vlad, Julian, Borgar, Faelar, Vance y yo, acompañados por diez milicianos seleccionados por el Sargento Roderick.",
        "Apenas habíamos recorrido tres leguas cuando Kaelan emergió del follaje con la respiración contenida y los ojos felinos dilatados:",
        "—Carretas en el camino viejo del peaje —susurró el lince señalando hacia una curva rocosa—. Huelo grano recién molido, pero también sudor rancio, orina y miedo. Mucho miedo.",
        "—¿Mercenarios? —preguntó Vlad desenvainando un palmo de su hoja.",
        "—Seis hombres armados con ballestas y cotas de malla con el blasón de Morvath —confirmó Akame, descendiendo de una rama de pino sin quebrar una sola brizna—. Escoltando dos carromatos pesados y una jaula cubierta con lonas de brea.",
        "Julian apretó los puños hasta que sus nudillos se tornaron blancos. El blasón que aquellos hombres llevaban en el pecho era el de su propia sangre.",
        "—Emboscada rápida —ordené en un susurro—. Borgar y Vlad bloquean el camino por el frente. Akame inutiliza a los ballesteros. Vance, cuida la retaguardia. No dejen que ninguno escape hacia la ciudad.",
        "* * *",
        "El asalto duró menos de un minuto.",
        "Borgar saltó desde una pendiente rugiendo con la furia de una avalancha de piedra. Su martillo de guerra impactó el eje de la carreta delantera, partiéndolo en astillas y deteniendo el convoy en seco. Antes de que los escoltas pudieran encajar sus virotes, dos sombras letales barrieron el fango: la hoja carmesí de *Murasame* destelló dos veces degollando a los tiradores de flanco, mientras Vlad aplastaba al capitán con su escudo pesado contra la rueda rota.",
        "—¡Ríndanse o sus cabezas rodarán antes de que toquen el suelo! —bramó Vlad con una autoridad que paralizó a los supervivientes.",
        "Los cuatro mercenarios restantes arrojaron sus armas al lodo, temblando al ver las capas rojas y doradas de la Orden del León y el uniforme militar de la Academia Imperial.",
        "Me adelanté cortando de un tajo las amarras de la lona negra del carromato trasero.",
        "Lo que encontramos dentro nos heló la sangre.",
        "* * *",
        "Hacinadas en una jaula de barrotes de hierro oxidado, seis personas yacían atadas con cadenas de argollas reforzadas: dos mujeres jóvenes con los vestidos desgarrados, un anciano con moretones en el rostro y tres muchachos adolescentes que apenas alcanzaban los catorce años. Todos mostraban quemaduras superficiales en los tobillos: marcas en forma de ocho puntas, el estigma preliminar de Khorne.",
        "—Por todos los dioses... —susurró Vance von Holst, retrocediendo con el rostro desencajado y una mano temblorosa sobre su estoque—. Están traficando con su propia gente...",
        "Julian se abalanzó hacia el carromato, desenvainando su daga para cortar las sogas de los prisioneros:",
        "—¡Tranquilos! ¡Están a salvo, se los juro! —gritaba con la voz quebrada por el horror y la culpa—. ¡Soy Julian de Morvath! ¡Nadie volverá a ponerles una mano encima!",
        "Al escuchar el apellido, una de las mujeres soltó un alarido de terror y el anciano intentó interponerse con desesperación:",
        "—¡No nos lleven a las fundiciones! ¡Por piedad, joven amo, pagamos el diezmo! ¡Entregamos hasta la última fanega de cebada!",
        "—¡No voy a lastimarlos! —insistió Julian cayendo de rodillas en el barro frente a la jaula, arrojando su daga al suelo para demostrar que estaba desarmado—. Mírenme... por favor, mírenme. No soy como mi padre. No soy como Valerius. He venido a detener esta locura o a morir intentándolo.",
        "El anciano observó los ojos llorosos del joven noble, su respiración agitada y la presencia protectora de los cadetes imperiales. Lentamente, la histeria se transformó en un silencio atónito.",
        "* * *",
        "Mientras Faelar y los milicianos alimentaban a los cautivos con raciones secas y agua limpia, Seier y Vlad revisaron los cofres de hierro transportados en la primera carreta.",
        "No solo contenían sacos de trigo marcados con los sellos comunales de Fresnedas; también hallaron libros contables confidenciales firmados por el Conde Raymond y cartas con el sello personal de Valerius.",
        "Vlad leyó los pergaminos a la luz de una antorcha, y su expresión se tornó cada vez más sombría:",
        "—Es peor de lo que imaginamos —dijo Vlad pasándome las páginas—. Los silos de grano de Morvath no están vacíos por malas cosechas. Raymond y su primogénito provocaron una hambruna artificial confiscando el setenta por ciento de la producción agrícola bajo supuestos 'impuestos de guerra'.",
        "—¿Y a dónde iba ese trigo? —preguntó Akame con frialdad.",
        "—Lo contrabandeaban fuera del condado a través de los puertos fluviales para comprar armas negras y pagar a los mercenarios de Malrik —respondí con rabia tras examinar los recibos—. Y lo más atroz: a los campesinos que no podían pagar el tributo en moneda, les embargaban a sus hijos bajo el eufemismo de 'leva laboral obligatoria'. En realidad, los entregaban como ganado al culto de Khorne para alimentar el portal de las minas a cambio de lingotes de plata demoníaca.",
        "Borgar escupió sobre las cenizas del camino:",
        "—Malditas ratas traidoras. Gobernar exprimiendo a tu propia sangre y vender niños a los demonios... Ni los peores orcos de las cordilleras se atreven a tanta bajeza.",
        "* * *",
        "Decidimos marchar de inmediato hacia Fresnedas, escoltando a los rescatados y a los mercenarios capturados.",
        "Fresnedas era un pueblo sombrío, de techos de paja carcomidos por la humedad y calles desiertas. Cuando la comitiva entró en la plaza, los aldeanos se asomaron desde las ventanas trancadas armados con horquillas y guadañas, temiendo otra redada de los recaudadores.",
        "Pero al ver a sus propios familiares bajar de las carretas con vida, los cerrojos comenzaron a abrirse.",
        "Una multitud de casi doscientos aldeanos rodeó la plaza central entre llantos de reencuentro y miradas llenas de rabia contenida. Julian subió a los escalones del pozo comunal de piedra. El muchacho respiró hondo, sacó el sello de plata de la Casa Morvath y lo alzó frente a todos:",
        "—¡Pueblo de Fresnedas! —proclamó Julian con una voz que resonó firme sobre el viento frío—. Sé lo que han sufrido. Sé que durante tres años mi padre, el Conde Raymond, y mi hermano mayor Valerius han saqueado sus graneros, han dejado morir de hambre a sus ancianos y han entregado a sus hijos a una secta abominable en las minas.",
        "Un murmullo de ira sorda recorrió a la multitud. Varios hombres corpulentos aferraron sus herramientas de labranza con nudillos tensos.",
        "—¡No vengo a pedirles obediencia ciega! —continuó Julian, golpeándose el pecho con el puño—. Vengo a decirles que el blasón que llevo no representará más la tiranía ni la cobardía. He traído a la Sección Uno de la Academia Imperial. La orden de auxilio ya vuela hacia la Corona y cien caballeros reales vienen en camino. Pero hasta que lleguen, ¡no permitiré que saquen a una sola persona más de sus hogares!",
        "Julian desenvainó su espada corta y la clavó en la tierra entre los adoquines:",
        "—Si alguien aquí quiere pelear por su pan, por sus hijos y por su libertad, juro por mi honor y por los dioses que estaré en la primera fila de la trinchera. ¡Morvath no le pertenece a los demonios ni a los tiranos! ¡Morvath es de su gente!",
        "* * *",
        "Las palabras de Julian no eran la retórica vacía de un cortesano mimado; eran el grito desgarrado de un joven que elegía la justicia sobre su propio linaje podrido.",
        "Un herrero corpulento de barba gris, apodado Garret, dio un paso al frente arrojando su delantal de cuero:",
        "—El conde se llevó a mi sobrina el mes pasado —dijo Garret con voz ronca—. Nos dijeron que murió de peste en las barracas. Si hay una oportunidad de aplastar a esos bastardos y recuperar lo que queda de nuestro pueblo, cuente con mi mazo, joven señor Julian.",
        "—¡Y con mis hachas! —gritó un leñador tuerto.",
        "—¡Y con nuestros arcos de caza! —corearon tres cazadores jóvenes de las colinas.",
        "Uno a uno, más de sesenta hombres y mujeres dieron un paso al frente, alzando guadañas, hachas de tala, lanzas de caza y martillos. La chispa de la rebelión había prendido en el corazón del campesinado.",
        "* * *",
        "No perdimos un solo segundo.",
        "Bajo las instrucciones tácticas de Seier y Vlad, convertimos la plaza de Fresnedas en un cuartel de adiestramiento militar y logística de guerra.",
        "Borgar se adueñó de la fragua comunal con Garret: con fuelle y carbón al máximo, el enano comenzó a transformar azadones y rejas de arado en puntas de lanza reforzadas, filos de alabarda y planchas de acero para blindar los carros de transporte.",
        "—¡El hierro templado no discute con la carne! —reía Borgar entre martillazos ensordecedores—. ¡En cuatro horas les tendré cincuenta lanzas capaces de perforar cotas de malla mercenarias!",
        "Faelar y Akame reunieron a los cazadores y exploradores locales para enseñarles señales de silbato, emboscadas en maleza y vigilancia perimetral con arcos largos. Kaelan los instruyó en el rastreo de centinelas nocturnos.",
        "Mientras tanto, Vance von Holst instruía a los campesinos más robustos en la formación de falange y disciplina de escudos:",
        "—Mantengan la línea recta —ordenaba Vance con porte impecable pero sin altanería—. Cuando un mercenario a caballo cargue, no corran; claven el talón de la lanza en la tierra y usen el peso del bruto en su contra.",
        "Julian, vestido con una cota de cuero reforzada por Garret, entrenaba espada mano a mano junto a sus nuevos soldados, ganándose el respeto sincero de cada aldeano con sudor y golpes limpios.",
        "* * *",
        "Al caer la tarde, la rebelión popular de Morvath ya no era una multitud desordenada; era la recién nacida **Guardia Popular de Julian**, integrada por casi ochenta combatientes decididos a recuperar su hogar.",
        "Además, los carros interceptados nos proveyeron de seis sacos de harina que fueron distribuidos de inmediato entre las familias hambrientas, devolviendo la esperanza a Fresnedas.",
        "En ese momento, un hombre sucio de tizne y con las ropas desgarradas corrió hacia la fragua escoltado por Kaelan:",
        "—¡Señor Julian! ¡Caballeros! —jadeó el fugitivo cayendo de rodillas—. Soy Tarek, minero del sector sur... Logré escapar por los antiguos canales de drenaje de las fundiciones.",
        "Seier se arrodilló junto al minero:",
        "—Toma agua, Tarek. Háblanos de las minas.",
        "El hombre desplegó sobre una caja de municiones un trozo de tela manchada de carbón donde había dibujado los túneles subterráneos:",
        "—Los cultistas han traído altares de bronce negro a la galería principal. Valerius estuvo allí esta mañana furioso porque Malrik no reportó... Está reuniendo a todos los mercenarios en el castillo de la colina y ha ordenado cerrar los portones de la ciudad amurallada al anochecer.",
        "Tarek señaló un conducto que cruzaba por debajo de los cimientos de la muralla de Morvath:",
        "—Este canal de desagüe desemboca directamente detrás del cuartel de la guardia baja. Si entramos por aquí al caer la noche, podemos abrir las puertas principales desde adentro sin que los centinelas de las almenas nos detecten.",
        "Vlad y Seier cruzaron miradas cargadas de resolución táctica.",
        "—Es la oportunidad que necesitábamos para limpiar Morvath desde adentro hacia afuera —sentenció Seier, poniendo su mano sobre el hombro de Julian—. Esta noche quebramos el cerco del conde."
    ]
}

# ===============================================================================
# FUNCIONES DE GESTIÓN Y COMPILACIÓN
# ===============================================================================

def guardar_capitulo_individual(cap_dict, base_dir=NOVELA_DIR):
    """Guarda el archivo .docx individual del capítulo."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    num = cap_dict.get("num", 0)
    if num == 0 or "PROLOGO" in cap_dict.get("header", "").upper() or "PRÓLOGO" in cap_dict.get("header", "").upper():
        base_name = "00_Prologo"
    else:
        base_name = f"{num:02d}_Capitulo_{num}"
        
    docx_path = os.path.join(capitulos_dir, f"{base_name}.docx")
    create_single_chapter_document(
        cap_dict,
        docx_path,
        novel_title="El camino del reencarnado",
        author="Seier Silva"
    )
    print(f"[OK] Documento Word individual guardado:\n  {docx_path}")
    return docx_path


def cargar_todos_los_capitulos(base_dir=NOVELA_DIR):
    """Lee todos los capítulos directamente desde la carpeta capitulos/*.docx."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    docx_files = sorted(glob.glob(os.path.join(capitulos_dir, "*.docx")))
    
    # Intentar cargar mapa de ilustraciones
    ilustraciones = {}
    try:
        from ilustraciones_capitulos import ILUSTRACIONES_POR_CAPITULO
        ilustraciones = ILUSTRACIONES_POR_CAPITULO
    except Exception:
        pass
        
    lista_capitulos = []
    
    for fpath in docx_files:
        filename = os.path.basename(fpath)
        num = None
        m_num = re.match(r"^(\d+)_", filename)
        if m_num:
            num = int(m_num.group(1))
            
        doc = docx.Document(fpath)
        if not doc.paragraphs:
            continue
            
        title_p = doc.paragraphs[0].text.strip()
        parts = title_p.split("\n")
        header = parts[0].strip() if len(parts) > 0 else "Capítulo"
        title = parts[1].strip() if len(parts) > 1 else ""
        
        paragraphs = []
        for p in doc.paragraphs[1:]:
            t = p.text.strip()
            if not t:
                continue
            if t in ["— ✦ —", "— ◆ —", "✦ ✦ ✦", "— ✧ —"]:
                continue
            if t.startswith("Ilustración:") or t.startswith("Ilustracion:"):
                continue
            paragraphs.append(t)
            
        chap_dict = {
            "header": header,
            "title": title,
            "paragraphs": paragraphs
        }
        
        # Asociar ilustraciones
        if num in ilustraciones:
            illus_info = ilustraciones[num]
            if "illustrations" in illus_info:
                chap_dict["illustrations"] = illus_info["illustrations"]
            elif "illustration" in illus_info:
                chap_dict["illustration"] = illus_info["illustration"]
                chap_dict["illustration_caption"] = illus_info.get("illustration_caption", "")
                
        lista_capitulos.append(chap_dict)
        
    return lista_capitulos


def compilar_libro_completo(base_dir=NOVELA_DIR):
    """Recompila el libro maestro completo a partir de todos los .docx de capitulos/."""
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    
    capitulos = cargar_todos_los_capitulos(base_dir)
    print(f"[*] Compilando Libro Maestro con {len(capitulos)} capítulos (desde .docx)...")
    
    compilado_dir = os.path.join(base_dir, "compilado")
    os.makedirs(compilado_dir, exist_ok=True)
    master_docx = os.path.join(compilado_dir, "El_Camino_del_Reencarnado.docx")
    
    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, master_docx)
        print(f"[OK] Libro Maestro actualizado exitosamente en:\n  {master_docx}")
    except PermissionError:
        print(f"[AVISO] '{master_docx}' está abierto en Microsoft Word.")
        backup_master = os.path.join(compilado_dir, "El_Camino_del_Reencarnado_temp.docx")
        create_novel_document(titulo, subtitulo, autor, capitulos, backup_master)
        print(f"[OK] Guardado en copia auxiliar:\n  {backup_master}")


def ejecutar():
    print("===============================================================")
    print(f" AGREGAR / ACTUALIZAR: {CAPITULO['header']} - {CAPITULO['title']}")
    print("===============================================================")
    guardar_capitulo_individual(CAPITULO, NOVELA_DIR)
    compilar_libro_completo(NOVELA_DIR)
    print("===============================================================")
    print(" ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("===============================================================")

if __name__ == "__main__":
    ejecutar()
