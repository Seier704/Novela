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
    "num": 78,
    "header": "Capítulo 78",
    "title": "Crónicas del Despliegue: Ecos de Fuego, Sombras y Acero",
    "paragraphs": [
        "El fuego del hogar de *La Jarra Plateada* crepitaba con fuerza, proyectando destellos dorados y sombras alargadas sobre las vigas de roble y las jarras de peltre.",
        "A medida que la cena avanzaba y los oscuros relatos de Morvath dejaban paso al tintineo de los cubiertos, apoyé los codos sobre la mesa de la cabecera. Recorrí con la mirada a cada uno de mis hermanos de armas: Ken, Caesar, Subaru, Kaoruko, Lucy, Erza, Yami y los demás. Una media sonrisa retadora se dibujó en mis labios al dar un trago a mi cerveza:",
        "—Vlad, Akame y yo ya les contamos cómo la Sección Uno le arrancó las fauces a la herejía en el este... —dije con voz serena pero provocadora—. Pero la reforma del Instructor Gerald nos dispersó por todo el reino precisamente para que nadie se acomodara. Separaron a las parejas y nos mezclaron en secciones aleatorias para romper camarillas y probar si podíamos valernos por nosotros mismos. Así que no crean que se librarán de rendir cuentas. ¿Qué hicieron ustedes durante este mes mientras nosotros nos arrastrábamos por fundiciones y minas de cal viva?",
        "Ken fue el primero en responder. El artista marcial golpeó la mesa con su jarra, haciendo bailar los platos entre carcajadas:",
        "—¡Ja! ¡Nosotros no nos quedamos rascándonos la panza, hermano! ¡Nuestra misión de la Sección Cuatro en la Garganta de Fuego fue puro infierno desde el primer día!",
        "* * *",
        "Ken se acomodó la bandana roja en la frente y comenzó su relato con los ojos chispeantes de ardor marcial.",
        "La Sección Cuatro había sido conformada por una mezcla improbable: Ken, Lady Elysea de Vane, Piper Wheel y tres antiguos cadetes nobles de la vieja Clase A-1. Al inicio, aquellos aristócratas estirados miraban con abierto desdén a Ken por pelear a puño limpio sin coraza de acero, y trataban a Piper como a una 'mocosa callejera engrasada' por comer galletas sobre los mapas rúnicos mientras ajustaba sus llaves mecánicas.",
        "—Lady Vane los puso en su sitio la primera noche —relató Ken entre risas—. Desenvainó su estoque gélido *Escarcha Silenciosa* y le congeló la copa de vino a uno en la mano a un milímetro de su nariz, advirtiéndoles que si volvían a faltarnos al respeto los dejaría pegados al suelo como estatuas de hielo.",
        "La verdadera prueba llegó en el desfiladero de basalto de la Garganta de Fuego. Una caravana de refugiados mineros fue emboscada por un Drake de Magma Alfa: una bestia colosal de escamas volcánicas que exhalaba llamaradas a más de dos mil grados. En aquel cañón asfixiante, las pesadas armaduras de placas de los nobles se convirtieron en hornos mortales que casi los cocinan vivos por la radiación térmica.",
        "—No había tiempo de formar filas —continuó Ken levantándose un poco de su banco—. Activé el *Paso del Fénix*, salté sobre las rocas al rojo vivo y me metí bajo la mandíbula del reptil. Lady Elysea coordinó el ataque con una precisión aristocrática impecable: lanzó una estocada de escarcha concentrada directo a la placa pectoral de la bestia, enfriando el basalto candente en una milésima de segundo.",
        "El brutal choque térmico fracturó la coraza de piedra. Fue en ese instante cuando Ken descargó su técnica marcial: un gancho ascendente imbuido en fuego solar puro que penetró la grieta, pulverizando la quijada del Drake y arrojándolo al fondo del abismo en una llamarada dorada. Mientras tanto, Piper Wheel reparó sobre la marcha un cabrestante averiado con su motor de vapor de alta presión, arrastrando las carretas con los niños fuera del alcance de los derrumbes.",
        "—Al final de la misión —concluyó Ken sonriendo de oreja a oreja—, los mismos nobles que se burlaban de nosotros terminaron cargándome la mochila y sirviéndole el té a Piper con la cabeza gacha. ¡Respeto ganado a golpe limpio!",
        "—¡Así se habla, muchacho! —celebró Borgar chocando ruidosamente su jarra con la de Ken—. ¡A las alimañas de roca se les quiebra el espinazo combinando calor y frío! ¡Una técnica digna de un enano!",
        "* * *",
        "Caesar King soltó una carcajada vibrante, mostrando una mella fresca en el borde de su gigantesco escudo de titanio:",
        "—¡Pues la Sección Cinco no se quedó atrás en acción bruta! Aunque admito que me tocó hacer de niñera de dos bichos raros: este estratega de aura plateada y una científica que casi nos desintegra el campamento.",
        "Subaru Natsuki, sentado frente a Caesar, suspiró teatralmente mientras se frotaba los hombros con una mueca de dolor cómico:",
        "—¡Sobrevivir a la Sección Cinco fue una hazaña peor que derrotar al culto! Caesar me despertaba a las cinco de la mañana a base de patadas para hacer trescientas flexiones en el barro diciendo que 'mi chasis necesitaba temple'. ¡Y mientras tanto, Agnes Tachyon no paraba de mezclar matraces humeantes que olían a azufre y maná inestable junto a mi bolsa de dormir!",
        "La Sección Cinco había sido enviada al Paso del Trueno y los valles mineros septentrionales, plagados de nidos de mantícoras de tormenta que cazaban a los mineros imperiales. Subaru tomó las riendas tácticas del despliegue: en lugar de cargar a ciegas por los riscos, diseñó un sistema de silbatos de alta frecuencia y bengalas de fósforo para saturar los sentidos de las bestias y forzarlas a canalizarse por un desfiladero estrecho donde no pudieran desplegar las alas.",
        "—Y ahí fue donde entró la capitana —añadió Subaru con sincera admiración—. Caesar se plantó en la boca del cañón como una cordillera viva. Aguantó la embestida de la mantícora alfa —un monstruo de cuatro toneladas cayendo en picada— con una risotada salvaje y un golpe de escudo que le partió las costillas en el aire.",
        "El golpe de gracia lo dio Agnes Tachyon: arrojó sus nuevos frascos de reactivos termodinámicos inestables al centro de la manada, desatando una explosión sónica de frío hiperbárico que congeló las alas de las bestias restantes, rindiéndolas sin necesidad de derramar más sangre.",
        "—¡Buen cerebro y agallas de titanio! —rugió Caesar dándole a Subaru una palmada en la espalda que casi lo hace escupir la cerveza—. ¡Tienes madera de líder, novato!",
        "Desde el otro lado de la mesa, Lucy (que había estado en la Sección Dos) fingió desdén cruzándose de brazos, pero sus mejillas sonrojadas y sus ojos dorados no pudieron ocultar el alivio y el orgullo al ver a Subaru sano y salvo:",
        "—Más te valía no hacerte matar por los frascos de Agnes, idiota... —murmuró la rubia con un mohín tierno que provocó sonrisas en toda la mesa.",
        "* * *",
        "A mi lado, Kaoruko Waguri sonrió con serenidad, rozando el dorso de mi mano con sus dedos cálidos antes de tomar la palabra animada por la mirada atenta del grupo.",
        "—A la Sección Seis nos asignaron una tarea muy distinta, pero que tocó profundamente nuestros corazones —explicó con su voz suave y melodiosa—. Yami, Lighter y yo fuimos enviados a custodiar la ruta médica del Valle de los Lirios, protegiendo un convoy con semillas, sanadores y trigo para las aldeas periféricas.",
        "En mitad de los bosques templados, una banda de más de cincuenta mercenarios desertores intentó tender una emboscada para saquear los cargamentos medicinales.",
        "—Yami dormía en el techo del carro con el puro encendido diciendo que su sola presencia espantaba a las moscas —relató Kaoruko con gracia contenida—. Pero cuando los jinetes cargaron disparando ballestas, Lighter saltó al frente como un rayo: sus puñetazos de combustión rápida desarmaron a los exploradores a una velocidad inverosímil, noqueándolos sin causarles heridas fatales.",
        "Cuando el grueso de los asaltantes cargó a caballo, Kaoruko intervino con maestría: desplegó una inmensa barrera sagrada entrelazada con zarzas y raíces florales de maná. La cúpula detuvo la lluvia de saetas en seco y las enredaderas atraparon suavemente las cinchas y las patas de las monturas, desarmando a la tropa completa sin herir a los animales ni quitar una sola vida humana.",
        "—Y cuando el líder de la banda sacó un orbe explosivo oscuro para volar el convoy —añadió Kaoruko mirando de reojo al espadachín de las sombras—, Yami apareció tras él en un parpadeo de miasma y le apoyó la vaina de su katana en la nuca. Le bastó una mirada para que el hombre se desmayara del puro terror.",
        "Kaoruko no solo salvó la caravana, sino que pasó semanas curando a los campesinos enfermos de las aldeas y compartiendo panes de miel recién horneados, recibiendo coronas de lirios blancos de los ancianos agradecidos.",
        "Apreté suavemente la mano de Kaoruko bajo la mesa, mirándola a los ojos con devoción y orgullo sincero. Ella bajó la vista con un dulce sonrojo, apretando mis dedos en correspondencia.",
        "* * *",
        "Erza Scarlet fue la siguiente en tomar la palabra, imponiendo con su mera presencia la disciplina marcial que la caracterizaba como líder de la Sección Dos.",
        "—Nuestra asignación nos llevó a los Pantanos de Niebla-Negra y las Ruinas Anegadas del norte —explicó Erza con tono solemne—. Dante Vance, Lucy y yo tuvimos que rastrear una infestación de acechadores de ciénaga y rezagados Skavens que habían atrapado a una patrulla de exploradores de tercer año en arenas movedizas.",
        "Lucy resopló divertida, haciendo girar un tenedor entre sus dedos:",
        "—Erza casi nos mata de agotamiento antes de llegar al pantano. Marchas forzadas de treinta kilómetros diarios en lodo hasta las rodillas desde las cuatro de la mañana, y cada vez que yo suspiraba pensando en la academia, ¡me caía una reprimenda por falta de concentración militar!",
        "—La disciplina salva vidas en el terreno, Lucy —replicó Erza con severidad afectuosa—. Y demostraste con creces el temple de una verdadera guerrera.",
        "Erza narró cómo Dante Vance neutralizó centinelas a cuatrocientos metros entre la bruma con su mosquete rúnico, mientras Lucy ejecutaba piruetas acrobáticas letales con sus dagas dobles y patadas ígneas para cortar las sogas de los prisioneros antes de que fueran sumergidos. Cuando una Hidra del Cieno emergió del lodo profundo, Erza invocó su *Armadura de la Emperatriz del Agua* para congelar y estabilizar el terreno pantanoso, cambiando en un destello de luz a su *Armadura de la Rueda del Cielo* para decapitar a la criatura con una danza de espadas flotantes.",
        "—Misión cumplida con cero bajas y los cadetes rescatados intactos —concluyó Erza, dedicándole a Lucy una sonrisa de genuina aprobación marcial que hizo resplandecer a la rubia de los Calydon.",
        "* * *",
        "Los ecos de las otras secciones también resonaron en la mesa: Rukia y Burnice en la Sección Tres, donde el hielo sagrado de la shinigami y los cócteles chispeantes de pimienta de la rubia calmaron tensiones en los valles rúnicos con el sigilo felino de Pulchra y el ojo certero de Hoshina; y la Sección Siete, donde Renji Abarai desató a *Zabimaru* contra golems de roca mientras Lortelle Kecheln aseguraba lucrativos acuerdos mercantiles de abastecimiento con los gremios mineros locales, respaldada por la firmeza de Freya y Zemina.",
        "Cuando el último relato se apagó en la taberna, una atmósfera de profunda admiración y camaradería envolvió la mesa.",
        "Miré a cada uno de ellos. Ken con sus nudillos ardientes; Caesar y Subaru sonriendo con complicidad táctica; Kaoruko y Lucy reflejando fortaleza y ternura; Erza y Vlad como pilares de acero; Yami disfrutando de su cerveza en paz. Ninguno era el mismo cadete novato e impulsivo que había comenzado el semestre.",
        "Teníamos diecisiete años. Habíamos sido separados deliberadamente para que aprendiéramos a luchar sin nuestras muletas afectivas, enfrentando monstruos, hambre, bandidos y tormentas con compañeros desconocidos. Y el resultado era incuestionable: la vieja división entre clases y estamentos se había disuelto para siempre en el crisol de la experiencia real.",
        "Alcé mi jarra rebosante hacia el centro de la mesa. Todos respondieron al unísono, haciendo chocar el metal con un estruendo formidable que retumbó en las vigas del local:",
        "—Nos separaron creyendo que dependeríamos de nuestras parejas o de nuestras antiguas aulas —dije con voz firme y llena de convicción—. Y lo único que lograron fue convertir a cada uno de ustedes en una fuerza imparable.",
        "—¡POR LA HERMANDAD DE LA F-2 Y LA GENERACIÓN DORADA! —rugió Ken alzando su puño envuelto en chispas solares.",
        "El choque de jarras, los cantos y las risas se prolongaron hasta el filo de la madrugada bajo el cielo estrellado de Astraea. La gran prueba del despliegue había concluido, pero la leyenda de nuestra generación apenas comenzaba a dar sus primeros pasos."
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
    for base in [NOVELA_DIR, PROYECTO_DIR]:
        if os.path.exists(base):
            print(f"[*] Procesando en: {base}")
            guardar_capitulo_individual(CAPITULO, base)
            compilar_libro_completo(base)
    print("===============================================================")
    print(" ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("===============================================================")

if __name__ == "__main__":
    ejecutar()
