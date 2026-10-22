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
    "num": 77,
    "header": "Capítulo 77",
    "title": "Retorno a las Torres Blancas: Abrazos, Fuego y Camaradería",
    "paragraphs": [
        "Ocho jornadas a lomo de caballo por la gran calzada imperial nos devolvieron finalmente al corazón del reino.",
        "Cuando cruzamos el último collado montañoso, las majestuosas agujas de mármol blanco y cristal de la Gran Capital de Solaria se alzaron contra el cielo del atardecer, refractando la luz dorada en destellos púrpuras y carmesíes. Las murallas monumentales de la Gran Academia Militar y Mágica se recortaban imponentes en el horizonte, anunciando el regreso a nuestro verdadero hogar.",
        "Apostados a la entrada de las puertas sur de la Academia, los centinelas de guardia se cuadraron de inmediato al divisar el estandarte de la Sección Uno y los veinte jinetes de la escolta real:",
        "—¡Abran paso a la Sección Uno! —gritó el oficial de guardia con reverencia marcial—. ¡La vanguardia de Morvath ha regresado!",
        "La noticia de nuestras hazañas nos había precedido por varios días a través de los despachos reales y el vuelo del espíritu de ave. A medida que cruzábamos los arcos del patio principal, decenas de estudiantes de diversas clases y grados se asomaban a las galerías, murmurando con asombro al ver las armaduras melladas por el combate, los sellos de honor condales y el porte imperturbable de nuestro escuadrón.",
        "* * *",
        "Apenas desmontamos en el patio adoquinado de la residencia estudiantil, el aire se llenó de pasos apresurados.",
        "—¡SEIER!",
        "Una voz dulce, cargada de una angustia contenida que se transformó en puro alivio, resonó bajo los arcos de piedra.",
        "Kaoruko Waguri corrió hacia mí a toda prisa, dejando caer a un lado la cesta de panecillos dulces que llevaba en brazos. No le importó el barro seco de mi abrigo ni el polvo del camino: se arrojó contra mi pecho con tanta fuerza que sus brazos rodearon mi cuello de golpe, enterrando su rostro en el hueco de mi hombro entre sollozos ahogados.",
        "—Estás aquí... estás vivo... por todos los dioses, volviste... —susurraba Kaoruko con la voz rota de emoción, apretándose contra mi cuerpo como si temiera que me desvaneciera si me soltaba.",
        "La envolví en un abrazo firme y cálido, sintiendo el aroma familiar a vainilla y flores silvestres de su cabello. Mis manos masivas acariciaron su espalda con ternura contenida:",
        "—Te dije que volvería, Kaoruko —le murmuré al oído con voz suave y profunda—. Estoy entero. No tienes de qué preocuparte.",
        "Kaoruko levantó el rostro empapado en lágrimas, pasando sus dedos delicados por mis mejillas, examinando cada cicatriz superficial y cada marca de cansancio, antes de apoyarse en puntillas y plantar un beso tembloroso y apasionado en mis labios, sellando semanas de silencio con una entrega pura y desinhibida que arrancó silbidos cómplices de los muchachos.",
        "* * *",
        "A pocos metros de nosotros, la escena se repetía con intensidad desbordante.",
        "Zemina, vestida con su elegante uniforme aristocrático, rompió toda compostura al ver a Vlad. El caballero de la Orden del León arrojó su morral al suelo, la levantó en vilo en medio del patio haciéndola girar en el aire y la besó con una pasión tan ardiente que la joven noble se aferró a sus hombros con las mejillas encendidas de un rubor carmesí.",
        "—¡Te extrañé como un loco, mi dama! —reía Vlad entre risas y besos robados.",
        "—¡Idiota presumido! ¡Casi me matas de preocupación con tus cartas! —reprochaba ella con dulzura, golpeándole el pecho sin ninguna fuerza.",
        "En ese momento, una ráfaga de calor abrasador barrió los adoquines del patio.",
        "—¡Miren a quién trajo el viento! —rugió una voz enérgica y desafiante.",
        "Ken emergió de las arcadas del gimnasio con el torso cubierto de sudor, los vendajes de combate humeantes y su clásica bandana roja atada a la frente. El discípulo de Hwa Young sonreía de oreja a oreja, con las llamas doradas danzando en sus nudillos:",
        "—¡Seier! ¡Escuché que le partiste la mandíbula al primogénito de un conde y reventaste un altar demoníaco en las minas!",
        "Me acerqué a él y chocamos los antebrazos con un impacto seco y contundente que resonó en el patio, liberando una pequeña onda de choque cinético que dispersó las hojas secas.",
        "—Y tú pareces haber estado quemando piedras todo el mes, Ken —respondí con una sonrisa ladeada, apreciando la densidad de su maná de fuego.",
        "—¡Hwa Young me hizo entrenar bajo cascadas hirvientes! —carcajeó Ken con los ojos brillando de entusiasmo marcial—. ¡Cuando descanses tenemos un duelo pendiente, hermano! ¡Quiero probar la dureza de esas nuevas grebas que traes puestas!",
        "* * *",
        "La reunión no tardó en completarse con la llegada del resto del grupo.",
        "Yami apareció apoyada contra una columna de granito, con los brazos cruzados y una sonrisa burlona cargada de orgullo:",
        "—Parece que el bruto de la F-2 aprendió a mandar tropas sin que se le caiga la espada. Buen trabajo, Seier.",
        "Detrás de ella, Caesar irrumpió con su porte imponente y su voz de trueno, propinándome una palmada en la espalda que habría tumbado a un novato:",
        "—¡Eso es tener agallas, muchacho! ¡Poner a raya a un conde traidor y regresar con laureles imperiales! ¡Esta noche la cerveza corre por mi cuenta!",
        "Natsuki Subaru y Lucy se acercaron sonrientes. Subaru extendió el puño con su habitual teatralidad:",
        "—¡El escuadrón estrella ha vuelto intacto! Nosotros tuvimos que lidiar con nidos de bestias en los pantanos del norte, pero lo tuyo suena a cantar de gesta medieval, Seier.",
        "Lucy, más relajada y con una mirada dulce, saludó con un asentimiento cómplice a Akame:",
        "—Me alegra que todos estén a salvo, Akame.",
        "—Fue una buena caza, Lucy —respondió la asesina, dedicándole una sonrisa serena.",
        "Erza Scarlet y Agnes Tachyon también se sumaron a la comitiva. Erza inspeccionó con mirada clínica los filos de nuestras armas, asintiendo con sincera aprobación militar:",
        "—Disciplina impecable. La Sección Uno ha dejado en alto el estándar de combate de nuestra generación.",
        "Agnes, por su parte, se ajustó los anteojos y comenzó a acribillar a Faelar a preguntas sobre las perturbaciones electromagnéticas del portal de Khorne.",
        "* * *",
        "No dejamos a nuestros nuevos hermanos fuera del calor del recibimiento.",
        "Seier y Vlad tomaron la palabra frente al grupo para presentar formalmente a los compañeros que pelearon codo a codo en Morvath: Borgar Yunque-Férreo, Faelar Vaelen, Kaelan Garra-Pardo y Vance von Holst.",
        "Borgar encajó de inmediato con Caesar y Ken al descubrir que compartían la pasión por el licor fuerte y los desafíos de fuerza bruta. Kaelan fue recibido con calidez por Subaru y Lucy, mientras Faelar compartía impresiones botánicas y tácticas con Erza.",
        "Pero el momento más significativo fue el recibimiento a Vance von Holst.",
        "El noble menor temía las miradas de rechazo de la antigua Clase F por su pasado aristocrático, pero Yami, Kaoruko y Subaru le ofrecieron una jarra y una palmada sincera de bienvenida:",
        "—Si sangraste junto a Seier y Vlad, eres de los nuestros, Vance —le dijo Yami con franqueza—. Olvida las etiquetas de salón; aquí lo que cuenta es la lealtad.",
        "Vance tragó saliva y asintió con una reverencia respetuosa, sintiéndose por primera vez en su vida parte de una verdadera hermandad sin falsas apariencias.",
        "* * *",
        "Esa misma noche, trasladamos la celebración a *La Jarra Plateada*, la posada y taberna donde comenzó nuestra historia en la capital.",
        "Unieron seis mesas de roble en el centro del local. Se sirvieron bandejas colmadas de costillas glaseadas, salchichas ahumadas, quesos de montaña y jarras humeantes de cerveza rubia y sidra espumosa. La taberna vibraba con canciones guerreras, carcajadas estruendosas y el relato detallado de cada combate en Morvath: desde la emboscada bajo la lluvia en Sauceda hasta el asalto vertical del Risco del Buitre y la demolición del Altar de Bronce.",
        "Ken y Borgar competían en pulsos sobre una mesa auxiliar entre los vítores de Caesar y Kaelan. Vlad y Zemina conversaban en una esquina en susurros cariñosos, con sus frentes unidas y miradas llenas de complicidad. Vance reía abiertamente con Subaru, mientras Faelar afinaba su flauta de madera al compás de las risas.",
        "Me quedé un momento sentado en la cabecera, observando a todos con una profunda sensación de plenitud.",
        "Teníamos diecisiete años recién cumplidos. Habíamos llegado a esta academia como marginados, huérfanos o aristócratas en declive; y ahora, forjados en la sangre, el sudor y el combate real, éramos un núcleo invencible capaz de desafiar a imperios y abismos por igual.",
        "Kaoruko se sentó a mi lado, deslizando su mano suave por debajo de la mesa hasta entrelazar sus dedos cálidos con los míos.",
        "Apoyó su cabeza con delicadeza en mi hombro robusto, cerrando los ojos al compás de la música:",
        "—Gracias por volver a casa, Seier.",
        "Apreté suavemente su mano y levanté mi jarra hacia la mesa rodeado de mis hermanos de armas:",
        "—Por la hermandad, por los caídos y por el camino que aún nos queda por recorrer.",
        "—¡POR LA HERMANDAD! —rugió la taberna al unísono, chocando el acero y la madera en una noche de gloria que jamás olvidaríamos."
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
