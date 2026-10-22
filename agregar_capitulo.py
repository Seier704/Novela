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
# ===============================================================================
CAPITULO = {
    "num": 74,
    "header": "Capítulo 74",
    "title": "Sangre en las Profundidades: El Altar de Bronce y la Llegada de los Cien",
    "paragraphs": [
        "El aire en la boca de las minas del sur no olía a tierra ni a humedad de montaña; olía a azufre quemado, herrumbre caliente y sangre coagulada.",
        "Apenas habíamos dejado el castillo en manos de Garret y la Guardia Popular para proteger a Leo y custodiar a los prisioneros, marchamos a paso forzado hacia las cordilleras escarpadas. Guiados por Tarek, nos internamos en los túneles principales de extracción.",
        "A medida que descendíamos por los socavones de piedra caliza, la temperatura subía como si nos adentráramos en las fauces de una fragua volcánica. Las paredes de roca rezumaban un líquido viscoso de color rojizo y las vigas de soporte crujían bajo una vibración rítmica y profunda que sacudía el polvo de los techos.",
        "—Estamos cerca de la gran cámara de fundición —susurró Tarek con la voz ahogada por un trapo húmedo—. Más adelante está el pozo del magma. Allí levantaron el altar.",
        "Kaelan se detuvo en seco, bufando con las orejas aplastadas:",
        "—Cánticos. Cientos de voces entrelazadas en una lengua podrida. Y algo enorme que respira como una caldera a presión.",
        "—Formación de cuña cerrada —ordené en voz baja, ajustando mis vendas y concentrando el flujo de maná denso hacia mis articulaciones—. Akame y Faelar toman las plataformas elevadas de los montacargas. Borgar y Vance limpian los puentes de acceso. Vlad y yo vamos directos al altar. No podemos permitir que el portal termine de abrirse.",
        "* * *",
        "Irrupimos en la galería central y el espectáculo que nos recibió fue una pesadilla dantesca.",
        "La caverna tenía más de cien metros de ancho, dominada en su fondo por un lago borboteante de cobre fundido y escoria incandescente. En el centro, conectado por dos puentes colgantes de cadenas y vigas de hierro, se alzaba una isla de piedra negra tallada en forma de un gigantesco cráneo cornudo: el Altar de Bronce de Khorne.",
        "Sobre las cuencas vacías del cráneo giraba un vórtice desgarrado en el tejido mismo del espacio: una espiral de relámpagos negros y llamas purpúreas que succionaba el aire de la cueva con un aullido sordo.",
        "Alrededor del altar, una treintena de cultistas con túnicas empapadas en sangre degollaban cabras negras y vertían cálices de aceite ígneo sobre los surcos de la piedra, liderados por un gigante de más de dos metros cubierto con placas de bronce brujo: el Sumo Sacerdote de la Sangre.",
        "—¡Sangre para el Dios de la Sangre! —aulló el sacerdote al divisarnos, alzando una alabarda envuelta en fuego demoníaco—. ¡El portal tiene sed! ¡Ofrezcan la carne de los intrusos a las fauces de Khorne!",
        "Del vórtice descendió un rugido que hizo temblar la montaña entera: una bestia cuadrúpeda forjada de bronce negro, músculo hipertrofiado y cuernos retorcidos —un Juggernaut menor del abismo— cruzó la brecha aterrizando con el estruendo de un meteorito sobre el puente de hierro.",
        "* * *",
        "El combate estalló sin tregua.",
        "Faelar y Akame se dispersaron por las pasarelas superiores con la gracia letal de cazadores nocturnos. El elfo silvano soltó una ráfaga de flechas imbuidas en vendavales cortantes que decapitaron a tres acólitos en pleno cántico, interrumpiendo el flujo del ritual. Al mismo tiempo, Akame cayó en picado desde una viga de madera: *Murasame* silbó como una guadaña invisible, cortando tendones y gargantas con su veneno fulminante, limpiando el perímetro del altar en un abrir y cerrar de ojos.",
        "En el puente oeste, Borgar y Vance hacían frente a una marea de fanáticos armados con mayales y hachas:",
        "—¡Por el orgullo de la forja! —rugió Borgar, descargando su martillo pesado contra el borde de la pasarela. La onda de choque de impacto quebró los tablones de hierro, haciendo que media docena de herejes cayeran al lago de cobre líquido entre alaridos desgarradores.",
        "Vance, con una frialdad y una elegancia insuperables, danzaba entre las arremetidas enemigas: su estoque de plata viva destellaba como una aguja implacable, encontrando las hendiduras de las viseras y las axilas con estocadas letales que detenían a los atacantes en seco.",
        "—¡Mantengan el puente! —gritó Vance, con su mirada aristocrática endurecida por la batalla—. ¡Nadie cruza hacia la retaguardia!",
        "* * *",
        "En la isla central, la mole de bronce y fuego cargó contra nosotros con la fuerza de una avalancha.",
        "—¡Yo detengo a la bestia! —bramó Vlad, dando un paso al frente.",
        "El heredero de la Orden del León clavó los talones en la piedra, activó la totalidad de su aura dorada y encajó su escudo de asedio en un ángulo inclinado perfecto. El impacto del Juggernaut resonó como una campana de catedral golpeada por un mazo de guerra. La onda expansiva agrietó el suelo bajo las botas de Vlad, pero el noble no retrocedió ni un solo milímetro: con un grito de poder puro, flexionó los brazos y desvió la embestida del monstruo hacia un costado.",
        "Esa fue mi entrada.",
        "Aproveché la fracción de segundo en que la bestia quedó desequilibrada: salté sobre el lomo blindado del engendro, envolví mis puños en una capa ultra-densa de maná telúrico y martillé la base de su cráneo de metal con tres golpes demoledores a corta distancia.",
        "El metal mágico se deformó bajo la compresión de mis nudillos. Antes de que pudiera sacudirse, me deslicé hacia su garganta, enganché mis piernas alrededor de su cuello en una llave de sumisión marcial y le hundí un codazo descendente en el núcleo rúnico de su frente. La placa de bronce estalló en esquirlas ardientes y la bestia se desplomó inmóvil sobre el suelo.",
        "* * *",
        "Solo quedaba el Sumo Sacerdote.",
        "El fanático se abalanzó sobre mí con su alabarda ígnea, trazando un corte horizontal diseñado para partirme el torso. Utilicé el desvío circular de Jeet Kune Do: me agaché por debajo de la trayectoria del fuego, atrapé su muñeca derecha con un giro de palanca inverso y apliqué presión hacia afuera, quebrando el cúbito y el radio con un chasquido seco.",
        "La alabarda cayó al suelo.",
        "El sacerdote aulló de rabia e intentó hundirme sus garras envenenadas con la izquierda, pero cerré la distancia de inmediato. Le encajé un rodillazo devastador de Muay Thai en las costillas flotantes que le vació los pulmones, seguido de un gancho ascendente al mentón que pulverizó su máscara de bronce.",
        "El hereje trastabilló hacia atrás con la mandíbula destrozada.",
        "Julian de Morvath emergió a mi lado con la espada de plata de su linaje envuelta en un fulgor brillante. De un tajo limpio y resolutivo, cortó el amuleto del sacerdote en dos mitades y atravesó su corazón con la punta del acero, poniendo fin a la herejía que había infectado su hogar.",
        "* * *",
        "—¡El portal! —gritó Tarek desde el puente—. ¡Sin el sacerdote, el vórtice está colapsando pero va a explotar hacia afuera!",
        "El torbellino de relámpagos negros se expandía con furia, agrietando los pilares de la caverna.",
        "Borgar corrió hacia el centro del altar y señaló con su martillo una grieta en la piedra basal:",
        "—¡Esa es la cuña de carga de la roca! —bramó el enano—. ¡Si rompemos ese pilar central, el altar entero se hundirá en el pozo de cobre y el vacío se cerrará sobre sí mismo!",
        "Vlad, Julian y yo intercambiamos una sola mirada cargada de determinación.",
        "Nos posicionamos alrededor del pilar agrietado. Vlad descargó su espada bastarda imbuida en luz dorada; Julian clavó su estoque de plata con toda su fuerza de voluntad; y yo concentré la totalidad de mi masa y maná telúrico en una patada circular baja con la tibia de hierro forjada en mis años de acondicionamiento.",
        "El triple impacto fue ensordecedor.",
        "La piedra basal se partió en dos como un cristal quebrado. Con un rugido colosal, la enorme estructura de bronce y hueso cedió, precipitándose hacia el abismo de cobre hirviente.",
        "El vórtice dimensional soltó un último destello purpúreo, perdió su ancla material y colapsó sobre sí mismo con una detonación de viento sónico que apagó las brasas y dejó la cueva en un silencio absoluto.",
        "El portal a Khorne estaba sellado para siempre.",
        "* * *",
        "Ascendimos por los túneles de regreso a la superficie, cubiertos de hollín, sudor y ceniza, pero caminando con el paso firme de quienes han mirado al abismo a los ojos y han prevalecido.",
        "Al cruzar la arcada exterior de las minas, la luz del mediodía nos bañó con una calidez reconfortante. El cielo sobre Morvath estaba completamente despejado, limpio de todo miasma oscuro.",
        "Y entonces, resonó en el horizonte un sonido majestuoso.",
        "No eran tambores demoníacos ni campanas de alarma fúnebres. Era el toque noble y solemne de trompetas de guerra de plata y oro que bajaba desde las colinas del este.",
        "Cruzando la loma en perfecta formación de cuatro columnas, un deslumbrante contingente militar marchaba a galope firme: corceles de batalla con bardas de acero pulido, capas escarlatas ondeando al viento y el estandarte del Sol Radiante de Solaria coronando la vanguardia.",
        "Eran los **Cien Caballeros Nobles de la Corona**.",
        "Al frente cabalgaba un Lord Paladín de la corte real con armadura blanca y dorada, deteniendo su montura al divisar la silueta de Morvath. El emisario real contempló las murallas intactas, las banderas de la Academia y de Julian ondeando en las almenas, a los ochenta aldeanos de la milicia custodiando las puertas con orgullo, y a los siete cadetes de la Sección Uno junto al joven Regente emergiendo victoriosos de las minas.",
        "El caballero paladín desmontó, se quitó el yelmo con una mirada de asombro y respeto reverente, y desenvainó su espada para saludar a los muchachos:",
        "—¿Quién está al mando de esta plaza? —preguntó con voz clara que resonó por todo el valle.",
        "Julian de Morvath dio un paso al frente con la frente en alto y el sello condal en su mano, mientras Vlad y yo nos colocábamos a sus flancos con sonrisas cansadas pero satisfechas.",
        "—Morvath está en paz, mi señor —proclamó Julian con la voz firme de un verdadero líder—. La herejía ha sido purgada, los tiranos están encadenados y el pueblo es libre."
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
