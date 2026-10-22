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
    "num": 73,
    "header": "Capítulo 73",
    "title": "Asedio a la Cumbre Oscura: Sangre en las Almenas y el Rescate de Leo",
    "paragraphs": [
        "El redoble de los tambores de Khorne retumbaba desde la cima de la montaña como el latido de un corazón monstruoso.",
        "A lo lejos, suspendido sobre el abismo de piedra afilada en el flanco este de la ciudadela alta, la silueta diminuta del pequeño Leo de Morvath oscilaba en el vacío, sostenida por una cadena herrumbrosa. Los sollozos del niño apenas se distinguían entre el estrépito de la lluvia y los cánticos guturales de los cultistas.",
        "Nos reagrupamos bajo el arco del Portón del Sur para trazar el plan en menos de dos minutos. El tiempo corría en nuestra contra: el alba despuntaba en el horizonte con una tenue franja carmesí.",
        "—Una embestida frontal por la rampa principal es un suicidio para Leo —analizó Vlad con el ceño fruncido, señalando el camino empinado flanqueado por saeteras y aspilleras de brea caliente—. Si Valerius nos ve subir en masa, soltará la polea y el niño se estrellará contra las rocas del fondo antes de que alcancemos el primer rellano.",
        "—Por eso nos dividiremos —dictaminé con firmeza—. Vlad, tú liderarás la fuerza de distracción junto a Vance, Borgar y treinta milicianos de Garret. Suban por la calzada golpeando los escudos, encendiendo antorchas y fingiendo un asalto masivo al portón levadizo superior. Atraigan toda la atención de Valerius y sus ballesteros hacia las troneras.",
        "—¿Y nosotros? —preguntó Akame, ajustando el tahalí de *Murasame* a su espalda.",
        "—Nosotros escalaremos el Risco del Buitre —respondí señalando la pared vertical de granito negro que caía a plomo justo detrás de la batería este—. Akame, Kaelan, Faelar, Julian y yo. Mientras Vlad absorbe el fuego, nosotros treparemos por el acantilado ciego y sacaremos a Leo de la soga.",
        "Julian apretó la empuñadura de su espada con nudillos ensangrentados:",
        "—Cuenten conmigo. No descansaré hasta arrancar a mi hermano de las garras de ese bastardo.",
        "* * *",
        "El plan se ejecutó con sincronía militar.",
        "Desde la rampa de acceso, un rugido de guerra sacudió la noche: Vlad activó el fulgor dorado de su maná ducal, alzando su escudo pesado como una muralla de bronce brillante, mientras Borgar golpeaba con su martillo una plancha de acero templado y los milicianos gritaban al unísono. Desde las almenas del castillo alto, los mercenarios entraron en pánico creyendo que la turba entera asaltaba la entrada:",
        "—¡Por aquí vienen los gusanos! —gritaban los arqueros desde las aspilleras—. ¡Preparen el aceite y disparen a la falange!",
        "La lluvia de virotes y pez ardiente comenzó a llover sobre el caparazón blindado de Vlad y Vance, absorbiendo cada proyectil con disciplina marcial.",
        "Entretanto, en la cara norte del acantilado, reinaba el silencio sepulcral.",
        "Kaelan lideró la escalada con una agilidad pasmosa, clavando sus garras semihumanas en las grietas húmedas de la roca. Faelar disparó dos flechas con amarras de seda élfica reforzadas con maná de viento que se anclaron profundamente en las cornisas superiores. Akame ascendía como un felino nocturno, liviana y sin rozar una sola piedra suelta.",
        "Detrás subíamos Julian y yo. Mis dedos, condicionados por diez años de trabajo físico pesado y canalización telúrica, se aferraban a las aristas del granito como tenazas de hierro. La fuerza isométrica de mis hombros y dorsales impulsaba mi peso hacia arriba palmo a palmo en medio del vendaval.",
        "* * *",
        "Alcanzamos el borde de las almenas orientales justo cuando el tercer toque del tambor resonaba en el patio de armas.",
        "Sobre la barbacana, dos fanáticos de Khorne con túnicas empapadas en brea sostenían la palanca del cabrestante donde colgaba la cadena de Leo. Valerius, enfundado en su armadura de hierro negro con picos rituales, se acercaba al borde con su montante desenvainado:",
        "—¡Se acabó el tiempo, Julian! —vociferó Valerius hacia la rampa vacía, ajeno a nuestra presencia a sus espaldas—. ¡El Dios de la Sangre reclama a su cachorro!",
        "Valerius alzó su espadón para partir el eslabón de un tajo.",
        "—¡Ahora! —rugí.",
        "Faelar soltó la cuerda de su arco recurvo: una flecha imbuida en un vórtice cortante de viento impactó de lleno la hoja del montante de Valerius en pleno descenso, desviando el acero con un chispazo que le arrancó un juramento furioso.",
        "En la misma fracción de segundo, la sombra de Akame barrió los adoquines: dos destellos carmesíes de *Murasame* decapitaron en un parpadeo a los dos fanáticos que custodiaban el cabrestante. Pero al desplomarse los cadáveres, la palanca quedó liberada y el engranaje comenzó a girar descontrolado, soltando la cadena al vacío.",
        "—¡LEO! —gritó Julian arrojándose hacia el borde.",
        "Reaccioné por puro instinto neuromuscular. Me impulsé con un salto rasante sobre las almenas, me deslicé por la piedra mojada y atrapé la cadena de eslabones gruesos con ambas manos desnudas.",
        "El tirón casi me arranca los hombros de las cuencas. El peso del niño cayendo por el abismo amenazaba con arrastrarme hacia las rocas del fondo. Activé la canalización interna de maná al límite: mis fibras musculares se hincharon como cables de acero, los tendones crujieron y mis pies se anclaron contra el pretil de piedra caliza, pulverizando el mortero bajo las suelas de mis botas.",
        "—¡HAAAA! —bramé desde las entrañas, tirando de la cadena eslabón por eslabón con pura masa compacta y rabia concentrada.",
        "Julian y Kaelan se sumaron a la tracción, y con un esfuerzo coordinado logramos izar la cadena sobre el parapeto.",
        "Leo cayó sobre los adoquines empapados, tosiendo y temblando de terror. Julian lo envolvió de inmediato entre sus brazos, protegiéndolo con su cuerpo mientras las lágrimas brotaban de los ojos del muchacho:",
        "—Ya estás a salvo, Leo... Mírame, hermanito, estás conmigo... nadie volverá a tocarte.",
        "* * *",
        "—¡Malditas ratas de alcantarilla! —rugió una voz enloquecida a nuestras espaldas.",
        "Valerius se giró con los ojos inyectados en sangre negra. Del cuello del noble colgaba un amuleto de bronce que irradiaba miasma pestilente, y sus venas palpitaban oscuras y gruesas por el consumo de elixires prohibidos de Khorne. Detrás de él emergieron seis berserkers de élite, guerreros mercenarios transformados en autómatas asesinos con hachas dobles.",
        "—¡Creen que pueden arrebatarme lo que es mío! —bramó el primogénito, blandiendo su espadón con una fuerza sobrehumana potenciada por el Caos—. ¡Si no puedo ofrecer al chiquillo, me bañaré en la sangre de todos ustedes!",
        "Los berserkers cargaron contra nosotros como toros enfurecidos.",
        "En ese instante preciso, el portón levadizo principal del patio estalló en pedazos: Vlad irrumpió en el adarve secundario encabezando la carga junto a Vance y Borgar, cortando la retirada enemiga.",
        "—¡Por la Casa Silva y la Orden del León! —bramó Vlad barriendo a dos berserkers con su escudo bendecido, mientras Vance ensartaba a un tercero con una estocada certera entre las placas de la garganta.",
        "—¡Este loco es mío! —le grité a mis compañeros, encarándome directamente con Valerius.",
        "* * *",
        "Valerius descargó su montante en un arco diagonal descendente que silbó con la furia de una guillotina. La fuerza del golpe era tal que, si intentaba pararlo con un arma común, me habría partido en dos.",
        "No intenté bloquearlo.",
        "Utilicé el balanceo de hombros del Jeet Kune Do: me deslicé hacia el flanco exterior del ataque, permitiendo que la hoja pesada se incrustara diez centímetros en el suelo de piedra. Antes de que pudiera recuperar el equilibrio, cerré la distancia como una exhalación.",
        "Le encajé un golpe seco con la base de la palma en la barbilla para romper su postura, seguido de un gancho al hígado con maná concentrado que hizo retumbar sus costillas protegidas por el acero negro. Valerius soltó un alarido de aire viciado, pero su locura demoníaca le impedía sentir el dolor común: intentó derribarme con un manotazo cargado de fuego oscuro.",
        "Atrapé su brazo con un desvío circular, aseguré la nuca de su casco con ambas manos y cerré el agarre de Muay Thai.",
        "Descargué tres rodillazos consecutivos contra su costado y su esternón con la potencia de un martinete hidráulico. La coraza de placas negras se hundió en una abolladura profunda. El impacto final fue un codazo descendente que le partió la celada del casco, haciéndole morder el polvo entre espasmos de sangre negra.",
        "El espadón de Valerius rodó por los adoquines.",
        "* * *",
        "Julian se puso de pie, dejando a Leo al cuidado seguro de Faelar, y avanzó hacia su hermano derrotado con la espada desenvainada.",
        "Valerius se arrastró por el suelo, escupiendo bilis y veneno:",
        "—Mátame... hermanito cobarde... —siseó con una risa entrecortada—. Mi padre ya firmó el pacto... Las minas de Khorne se abrirán... ¡Nadie detendrá al Dios de la Sangre!",
        "Julian alzó su acero. Por un segundo, la ira amenazó con nublar su juicio. Pero el muchacho respiró hondo, bajó la punta de la espada y golpeó a Valerius en la sien con el pomo macizo, noqueándolo en el acto.",
        "—No me rebajaré a tu nivel —dijo Julian con voz solemne—. No serás un mártir para tus demonios. Serás juzgado como un traidor ante el Rey y ante la gente que desangraste.",
        "Borgar sacó gruesas cadenas de hierro frío y amarró las muñecas y los tobillos de Valerius con nudos enanos imposibles de romper.",
        "* * *",
        "Avanzamos con paso firme hacia las estancias privadas del conde, en el corazón del torreón de homenaje.",
        "Derribamos las puertas de madera tallada esperando una última emboscada, pero lo que encontramos dentro fue un espectáculo patético y decadente:",
        "El Conde Raymond de Morvath no era un guerrero fiero ni un estratega temible. Yacía en una cama con dosel carcomido, consumido por la peste y la degeneración física de los rituales oscuros que había financiado. Sus ojos estaban hundidos, su piel parecía pergamino podrido y apenas podía levantar los brazos cubiertos de llagas negras, rodeado de cofres vacíos y cálices de sangre seca.",
        "Al ver a Julian, a la Sección Uno armada y al pequeño Leo sano y salvo, el viejo conde comenzó a sollozar en un ataque de pánico miserable:",
        "—Julian... hijo mío... perdóname... fueron los emisarios... prometieron devolverme la juventud y las riquezas... no me dejes morir...",
        "Julian caminó despacio hacia el lecho. Con una mezcla de desprecio y dolor contenido, tomó la mano marchita de su padre y le arrancó el anillo con el sello condal de la Casa Morvath:",
        "—Tu gobierno ha terminado, padre —sentenció Julian con frialdad—. Has vendido la vida de tus vasallos y el honor de nuestro apellido. A partir de este momento, asumo la regencia del feudo hasta la llegada del tribunal real.",
        "* * *",
        "Salimos a las terrazas del castillo alto.",
        "El sol de la mañana rompió por fin las nubes grises, iluminando Morvath con una luz dorada y limpia. Abajo en la plaza, los aldeanos de la Guardia Popular y los artesanos vitoreaban al ver ondear el pendón de Julian en la torre más alta junto al estandarte de la Academia Imperial.",
        "La ciudad amurallada estaba liberada. Los tiranos estaban encadenados y el pueblo estaba a salvo.",
        "Sin embargo, Tarek el minero subió a la terraza con el rostro cubierto de preocupación y nos mostró los documentos confiscados del escritorio de Raymond:",
        "—Señores... el castillo está asegurado, pero en el fondo de las minas del sur el portal de Khorne aún sigue alimentándose de la energía residual. Si no sellamos el altar antes del mediodía, las criaturas del abismo podrían cruzar la brecha.",
        "Vlad y yo contemplamos las humaredas lejanas de las montañas.",
        "—Un último esfuerzo —dije sonriendo de lado mientras me ajustaba los guantes—. Cerremos ese portal y devolvámosle la paz completa a este feudo."
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
