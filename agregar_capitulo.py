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
    "num": 72,
    "header": "Capítulo 72",
    "title": "Sombras de Acero y Muro Quebrado: La Apertura de las Puertas de Morvath",
    "paragraphs": [
        "La medianoche cayó sobre Morvath con una llovizna helada y pegajosa que amortiguaba el sonido de nuestros pasos.",
        "Apostados en la espesura del bosque que bordeaba el foso exterior, el equipo de asalto respiraba al unísono. Éramos ocho sombras listas para la brecha: Vlad, Akame, Borgar, Faelar, Kaelan, Vance, Julian y yo, guiados por el minero Tarek. Más atrás, ocultos entre los sauces del barranco, Garret y los ochenta aldeanos de la Guardia Popular aguardaban en silencio absoluto, empuñando sus lanzas recién forjadas con miradas fijas en las almenas.",
        "—Por aquí —susurró Tarek, señalando una abertura oculta entre raíces retorcidas y bloques de piedra cubiertos de musgo ceniciento—. Es el aliviadero de aguas negras de las fundiciones. Cruza bajo el baluarte y desemboca justo en el patio trasero de la guarnición baja.",
        "Kaelan se agachó primero. El joven semihumano dilatò sus pupilas felinas en la oscuridad y olfateó el aire fétido del ducto:",
        "—Huele a fango, óxido y grasa quemada —reportó en un murmullo—. Dos centinelas apostados a cincuenta metros, junto a la compuerta de retención. Están aburridos y medio dormidos.",
        "—Silencio absoluto —indiqué ajustando los vendajes de mis nudillos y la densidad de mi maná interior—. Akame, Kaelan, despejen el túnel. Vlad y Borgar, listos para forzar la reja.",
        "* * *",
        "Nos arrastramos por el conducto estrecho con el agua turbia empapándonos las rodillas. La piedra caliza goteaba sobre nuestras armaduras, pero nadie emitió un solo quejido.",
        "Al final del pasaje subterráneo, una luz anaranjada temblaba sobre el agua sucia. Dos mercenarios con cotas oxidadas roncaban apoyados contra la pared junto a una reja de hierro macizo. Antes de que pudieran pestañear, la penumbra cobró vida: Kaelan cayó sobre uno de ellos quebrando su tráquea con un zarpazo enguantado, mientras el filo silencioso de *Murasame* destellaba como un relámpago carmesí en la garganta del segundo.",
        "Ambos cuerpos se desplomaron en el fango sin emitir más que un gorgoteo ahogado.",
        "Borgar se acercó a la reja con una barra de desencofrar forjada por él mismo:",
        "—Hierro viejo de fundición barata —masculló el enano entre dientes—. Mis antepasados escupirían sobre esta soldadura.",
        "Haciendo palanca con sus hombros anchos de piedra, Borgar aplicó presión. Los pernos cedieron con un chasquido sordo que el rumor de la lluvia tapó por completo. La reja se abrió, dándonos acceso directo a los adoquines húmedos del corral de la guarnición baja.",
        "* * *",
        "A través de los ventanucos del cuartel bajo se escuchaban carcajadas roncas, el tintineo de dados de hueso y el chocar de jarras de estaño.",
        "La guardia personal de Valerius, confiada tras el cierre de los portones principales, se emborrachaba con el vino requisado a los campesinos.",
        "Julian se asomó por la rendija de madera. Al ver a los hombres que habían aterrorizado a Fresnedas y Sauceda bebiendo sobre mesas repletas de comida robada mientras los niños pasaban hambre, sus ojos se encendieron de furia:",
        "—Son los oficiales de Malrik que no marcharon a Sauceda —susurró el noble apretando la empuñadura de su espada—. Unos doce mercenarios y un cabecilla.",
        "—Entramos en formación de choque —ordené con voz gélida—. Vlad al frente. Vance cubre la escalera. Nadie toca la campana de alarma.",
        "Vlad dio un paso al frente, concentró su aura ducal en el escudo de la Orden del León y lanzó una patada demoledora que reventó el cerrojo de roble de par en par.",
        "—¡En nombre de la justicia imperial y de la Casa de Morvath, arrojen las armas! —rugió Vlad como un león en celo.",
        "* * *",
        "El caos estalló en una fracción de segundo.",
        "Los mercenarios volcaron las mesas entre juramentos y desenvainaron sus alfanjes. Un bruto corpulento de más de cien kilos se abalanzó sobre mí con un hacha de carnicero:",
        "—¡Mocoso insolente! ¡Te abriré de arriba a abaj...!",
        "Ni siquiera pestañeé.",
        "Di un paso en diagonal por dentro de su guardia (técnica de entrada de Jeet Kune Do), desvié el hachazo con el antebrazo blindado de maná telúrico y le encajé un codazo en punta ascendente que le pulverizó el pómulo y la mandíbula. El tipo se tambaleó con la boca llena de esquirlas óseas, pero antes de que tocara el suelo, cerré la distancia, le atrapé la nuca en un agarre cerrado de Muay Thai y le hundí un rodillazo devastador en pleno plexo solar.",
        "La cota de malla crujió hacia adentro como una lata aplastada. El hombre salió disparado contra una alacena de peltre, completamente inconsciente.",
        "A mi derecha, Akame era una danza de sombras y estocadas letales: tres mercenarios que intentaron cercarla cayeron al suelo con heridas limpias que detuvieron sus corazones al instante por la ponzoña maldita de su hoja. Vlad embistió con su escudo como un ariete viviente, aplastando a dos tiradores contra los pilares de carga del salón.",
        "Vance von Holst demostró la refinada letalidad de su esgrima aristocrática: con su estoque plateado, desarmó a dos oponentes con ripostes quirúrgicos en las muñecas, atravesando los huecos de las hombreras sin desperdiciar un solo gramo de energía.",
        "—¡Buen estoque, muchacho! —le gritó Borgar mientras barría las piernas de otro guardia con el reverso de su martillo.",
        "—Disciplina de academia, señor enano —respondió Vance con una inclinación cortante y elegante.",
        "* * *",
        "El capitán de la guarnición —un fanático calvo con runas de ocho puntas tatuadas en el cráneo— comprendió que la posición estaba perdida y corrió desesperado hacia la escalera de caracol del torreón:",
        "—¡Valerius! ¡Nos atacan desde adentro! —aulló mientras sacaba de su cinto un tubo de bengala de señales alquímicas.",
        "Subió tres peldaños y alzó el percutor hacia la aspillera para alertar al castillo alto.",
        "Un silbido de viento cortó el aire con precisión milimétrica.",
        "La flecha de Faelar, envuelta en un torbellino esmeralda de maná silvano, atravesó la mano del capitán de lado a lado, clavándola contra el pasamanos de roble macizo. La bengala cayó al suelo rodando apagada.",
        "Julian subió la escalera en dos zancadas, colocó la punta de su espada corta contra el cuello del hereje y de un golpe con el pomo le arrancó la máscara de bronce:",
        "—La campana no sonará esta noche, traidor —sentenció Julian con voz implacable.",
        "El cuartel bajo estaba completamente neutralizado.",
        "* * *",
        "Subimos de inmediato al adarve del Portón del Sur.",
        "La enorme maquinaria de engranajes y poleas de madera negra gobernaba las cadenas del rastrillo levadizo. Borgar y Vance encajaron las barras de madera en el cabrestante principal:",
        "—¡A la de tres! —bramó Borgar haciendo crujir sus músculos dorsales—. ¡Uno... dos... y ARRIBA!",
        "Con un chirrido ensordecedor de hierro y piedra, las cadenas de eslabones gruesos comenzaron a enrollarse. Las toneladas del rastrillo de hierro comenzaron a ascender palmo a palmo de las ranuras de la barbacana.",
        "Las pesadas puertas de roble blindado de Morvath se abrieron de par en par a la noche.",
        "Julian corrió hacia las almenas que daban al exterior, tomó el gran cuerno de guerra de cobre grabado con el león de su casa y sopló con toda la fuerza de sus pulmones.",
        "El bramido grave y ancestral del cuerno resonó sobre los campos mojados, rebotando en los riscos y rasgando la lluvia.",
        "La respuesta fue instantánea.",
        "A lo largo de la línea de árboles, decenas de antorchas se encendieron como un río de fuego. Garret el herrero, alzando su mazo de batalla, encabezó la carga al frente de los ochenta combatientes de la Guardia Popular:",
        "—¡Por el señor Julian! ¡Por nuestros hijos! ¡A las puertas!",
        "Los aldeanos armados cruzaron el foso en tromba y tomaron control del portal sur en cuestión de minutos, estableciendo una falange de lanzas y barricadas de carretas para sellar el perímetro y proteger las barriadas bajas.",
        "Los habitantes del distrito artesano comenzaron a asomarse desde las ventanas, atónitos al ver a sus propios parientes victoriosos ondeando el blasón de plata de Julian. Las lágrimas y los gritos de júbilo se mezclaron con el rumor de la lluvia: ¡el pueblo había recuperado su muralla!",
        "* * *",
        "Pero la victoria en las puertas apenas era el preludio.",
        "Desde la cumbre escarpada de la colina, donde el castillo ancestral de la Casa Morvath se erguía como una fortaleza oscura, las campanas de alarma comenzaron a redoblar con un tañido fúnebre y desquiciado. Junto al repique de bronce, un estruendo profundo y cavernoso sacudió los cimientos de la montaña: los tambores de piel humana del Culto a Khorne habían comenzado a batir.",
        "Sobre las altas almenas del castillo alto, iluminadas por antorchas de fuego verde y carmesí, apareció una figura envuelta en armadura negra de placas bruñidas:",
        "Valerius de Morvath.",
        "A su alrededor se apiñaban decenas de cultistas encapuchados con túnicas empapadas en sangre ritual. Y frente a ellos, suspendido en el vacío sobre el precipicio por una cadena de hierro atada a sus muñecas, colgaba un niño de doce años que sollozaba aterrorizado:",
        "El pequeño Leo.",
        "—¡Julian, hermano cobarde! —la voz de Valerius resonó amplificada por un artefacto mágico oscuro, cargada de una demencia fría y calculadora—. ¿Crees que has ganado algo abriendo las puertas a este rebaño de campesinos andrajosos? ¡Si no deponen las armas y me traen la cabeza de tus amiguitos de la academia antes del amanecer, soltaré la cadena de Leo y regaré las piedras del foso con sus sesos para complacer al Dios de la Sangre!",
        "Julian cayó de rodillas sobre las almenas, clavando las uñas en la piedra mojada con el rostro desencajado por la agonía:",
        "—¡Leo...! ¡Maldito monstruo... es tu propio hermano menor!",
        "Di un paso al frente y puse mi mano enguantada con firmeza sobre el hombro del muchacho, obligándolo a levantar la mirada. Vlad se colocó a mi lado, desenvainando su espada con un brillo dorado de maná que desafió la oscuridad de la montaña.",
        "—Levántate, Julian —dije con voz serena, áspera e inconmovible—. Ya quebramos la puerta de la ciudad. Ahora vamos a derribar su castillo y a recuperar a tu hermano.",
        "Akame desenvainó a *Murasame* un centímetro de su saya, y sus ojos carmesíes brillaron con una sed de justicia implacable:",
        "—La noche aún no termina. Asaltemos la cima."
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
