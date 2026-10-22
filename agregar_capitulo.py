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
    "num": 76,
    "header": "Capítulo 76",
    "title": "Vientos de Retorno: Lazos Forjados, Despedidas y el Camino a Solaria",
    "paragraphs": [
        "La última semana en Morvath pasó con la serenidad de una marea calma tras una tormenta despiadada.",
        "Los ecos de la gran fiesta de liberación dieron paso a un ritmo de trabajo disciplinado y meticuloso. No podíamos darnos el lujo de marcharnos dejando cabos sueltos; los siete integrantes de la Sección Uno nos volcamos por entero en los preparativos de marcha, la puesta a punto de nuestro armamento y la consolidación de las defensas locales.",
        "Cada amanecer comenzaba con el estruendo de la fragua comunal en el patio bajo del castillo.",
        "Borgar y Garret se habían vuelto inseparables frente al yunque. El maestro enano, fascinado por la pureza del mineral extraído de los nuevos filones limpios, dedicó tres días enteros a forjar piezas a medida para cada uno de nosotros antes de partir.",
        "—Toma esto, muchacho —me dijo Borgar, entregándome un par de brazaletes y grebas de hierro negro reforzado con inserciones de cobre templado—. Las correas de cuero común no soportan la densidad de tu maná telúrico cuando golpeas a fondo. Estas piezas distribuyen el choque por toda la estructura ósea; podrás partir rocas con las tibias sin desgastar las junturas.",
        "Me ajusté las grebas a las piernas: eran densas, perfectamente balanceadas y se amoldaban a mis pantorrillas como una segunda piel.",
        "Para Faelar, Garret y Borgar forjaron tres docenas de puntas de flecha con estrías helicoidales de cobre que estabilizaban el vuelo con vientos cruzados. A Vlad le bruñeron y reesmaltaron el león heráldico de su escudo; a Vance le entregaron una daga de parada a juego con su estoque de esgrima; y para Akame y Kaelan elaboraron cotas de malla ligera de escamas de acero que no entorpecían sus movimientos sigilosos.",
        "* * *",
        "A media mañana, el patio de armas se convertía en nuestro centro de entrenamiento.",
        "Vlad y Vance cruzaban aceros en duelos de técnica exquisita, mientras Kaelan y Akame practicaban desplazamientos a ciegas esquivando péndulos de madera. En una esquina despejada, Julian traía a su hermano menor Leo para observar la rutina.",
        "El niño de doce años aún mostraba timidez tras los horrores del secuestro, pero sus ojos brillaban con admiración pura cada vez que mis puños impactaban contra los postes de madera acolchada.",
        "—Ven aquí, Leo —le llamé con un gesto tranquilo, secándome el sudor de la frente.",
        "El pequeño dio unos pasos vacilantes hacia mí. Me agaché a su altura y le corregí la postura:",
        "—Separa los pies al ancho de tus hombros. Flexiona las rodillas y mantén la barbilla recogida. Nunca bajes la mirada, Leo. No importa cuán grande sea el enemigo o cuán oscura parezca la noche; un señor de Morvath siempre mira hacia adelante con la frente en alto.",
        "Leo apretó sus pequeños puños, imitó mi guardia y lanzó un golpe directo al aire con determinación infantil. Julian, observando desde la barandilla con una sonrisa conmovida, me dedicó un asentimiento cargado de gratitud.",
        "Poco después, Vance tomó a Julian bajo su tutela para instruirlo en la defensa de estoque y desvío de mandoble, puliendo los reflejos del joven Regente para que pudiera defenderse sin depender eternamente de escoltas.",
        "* * *",
        "A mitad de semana realizamos la última inspección de campo junto al contingente real.",
        "Recorrimos los pueblos de Sauceda y Fresnedas escoltados por una escuadra de caballeros de la Corona. El cambio era conmovedor: las empalizadas habían sido reparadas y reforzadas; las familias compartían pan caliente en mesas comunales; y el Sargento Roderick, luciendo con orgullo su capa de Comandante General de la Guardia Cívica, mantenía relevos de guardia impecables en cada atalaya.",
        "Luego descendimos hacia las minas del sur. Donde días atrás ardían los altares de Khorne, ahora se escuchaba el estrépito sano del trabajo honesto. Tarek nos recibió en la bocamina con una sonrisa tiznada de carbón:",
        "—Los tiros de ventilación que diseñó el señor Borgar funcionan a la perfección —nos informó, mostrándonos las vagonetas repletas de mineral reluciente—. El aire en el fondo es fresco y puro. Estamos extrayendo dos toneladas diarias de cobre de primera ley. Los barqueros del Consorcio ya esperan el primer embarque en el muelle fluvial.",
        "Morvath no solo había sobrevivido: estaba floreciendo.",
        "* * *",
        "La noche previa a nuestra partida, nos reunimos en la terraza oriental del castillo bajo un firmamento estrellado y límpido.",
        "Hacía frío, pero el brasero encendido y la complicidad de las últimas semanas nos envolvían en una calidez reconfortante. Vlad escribía con esmero una carta sellada con cera roja, tarareando para sí mismo con una sonrisa boba que no pasó desapercibida.",
        "—¿Escribiéndole a Lady Zemina? —bromeó Borgar, dándole un trago a su bota de licor.",
        "—A ella misma —admitió Vlad sin ruborizarse, guardando el pliego en su morral—. Dos semanas y media sin verla parecen una eternidad. Muero por volver a la capital y contarle todo lo que vivimos aquí.",
        "—Y por cobrar varios abrazos pendientes, me imagino —añadió Kaelan moviendo la cola con sorna, provocando las risas de Vance y Faelar.",
        "Vance von Holst miró las luces titilantes de la ciudad a nuestros pies y suspiró con serenidad:",
        "—Cuando me asignaron a la Sección Uno al inicio del semestre, creí que era un castigo por haber caído de la Clase A. Pensaba que compartir fila con plebeyos y semihumanos era una humillación. Dioses... qué ciego y estúpido fui. Esta misión me enseñó lo que de verdad significa la palabra honor. Gracias a todos por no dejarme atrás.",
        "Vlad le dio una palmada en la hombrera y chocamos nuestras copas en un brindis silencioso.",
        "Me aparté unos pasos hacia el pretil de piedra caliza para contemplar el valle en silencio. Un instante después, Akame se acercó a mi lado con dos manzanas recién cosechadas. Me ofreció una y apoyó los codos en la baranda, comiendo a pequeños mordiscos:",
        "—¿Pensando en las chicas de la F-2? —preguntó con su tono calmo y perspicaz.",
        "—Un poco —sonreí de lado, mordiendo la fruta dulce—. En Kaoruko, en Yami, en Subaru y los demás. Me pregunto cómo les habrá ido en sus respectivas misiones de largo alcance.",
        "—Conociéndolos, habrán dejado su propia huella —respondió Akame, mirándome de reojo con sus ojos carmesíes brillando bajo la luz lunar—. Pero dudo que hayan tenido que derribar un culto al Caos y salvar un feudo entero como nosotros. Hiciste un gran trabajo liderándonos, Seier.",
        "—No lo hice solo. Lo hicimos juntos —repliqué mirándola a los ojos—. Somos un buen equipo, Akame.",
        "La asesina sonrió con esa ternura discreta que reservaba solo para quienes confiaba con su vida, dejando que el viento de otoño meciera su cabello negro.",
        "* * *",
        "El amanecer del día siguiente tiñó los picos de las montañas con destellos rosados y dorados.",
        "En la explanada del Portón del Sur, más de mil habitantes de Morvath se habían congregado para despedirnos. Las carretas de viaje estaban cargadas con raciones de marcha, telas abrigadas y barriles de agua limpia. Una escolta de veinte caballeros nobles de la Corona, comandada por un lugarteniente del Lord Paladín, se alineó junto a nuestros corceles para acompañarnos a través del desfiladero montañoso.",
        "Julian y el pequeño Leo se adelantaron hacia nosotros.",
        "El joven Regente lucía la espada de plata ceñida al cinto y una capa con los nuevos colores de la Casa Morvath. Se cuadró frente a la Sección Uno y nos entregó un cilindro de plata lacrado con el nuevo sello oficial:",
        "—Este salvoconducto declara a cada miembro de la Sección Uno 'Huéspedes de Honor Perpetuos' del Condado de Morvath. Nuestras puertas, nuestros graneros y nuestras espadas estarán a su disposición siempre que lo necesiten.",
        "Leo se acercó a mí con paso firme, sacó de su bolsillo un pequeño colgante de madera tallada que representaba un león rugiendo y me lo puso en la mano:",
        "—Lo tallé para ti, señor Seier. Para que no olvides que te prometí ser fuerte.",
        "Cerré la mano sobre el amuleto con un nudo cálido en la garganta:",
        "—Lo llevaré siempre conmigo, Leo. Cuida de tu hermano y de este pueblo.",
        "Montamos a caballo. Vlad, Vance, Borgar, Faelar, Kaelan, Akame y yo nos alineamos al frente de la columna.",
        "Alcé mi brazo enguantado hacia el cielo matinal. Las trompetas de la milicia sonaron al unísono, acompañadas por los vítores atronadores de los aldeanos y el ondear de cientos de pañuelos blancos.",
        "—¡Sección Uno, en marcha! —ordené con voz firme.",
        "Los caballos rompieron al trote, cruzando el puente levadizo hacia el camino real. Dejábamos atrás un feudo renacido, una tiranía destruida y un lazo de hermandad forjado en sangre.",
        "El largo viaje de regreso a la Gran Academia Militar y Mágica de Solaria acababa de comenzar."
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
