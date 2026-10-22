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
    "num": 75,
    "header": "Capítulo 75",
    "title": "El Renacer de Morvath: Justicia, Honor y el Juicio de las Sombras",
    "paragraphs": [
        "Dos semanas transcurrieron desde que el portal de las minas del sur fue tragado por el magma incandescente, transformando el destino de Morvath para siempre.",
        "El hedor a azufre y sangre rancia que durante tres años asfixió al feudo se disipó bajo los vientos limpios del otoño. Donde antes reinaban el miedo, las cortinas cerradas y los pasos apresurados de los recaudadores, ahora resonaban los martillazos de los carpinteros, el murmullo bullicioso de los mercados y las risas de los niños jugando en las plazas.",
        "Fueron catorce días de labor implacable, reconstrucción febril y justicia solemne.",
        "* * *",
        "El primer gran acto de purga ocurrió en el Gran Salón del castillo alto, convertido en sede del Tribunal Militar Extraordinario presidido por el Lord Paladín de la Corona y los magistrados reales.",
        "Los traidores comparecieron despojados de toda pompa aristocrática. Valerius de Morvath, con el rostro marcado por los hematomas de mi combate en las almenas y las muñecas aprisionadas en grilletes de hierro frío, escuchó su sentencia sin un ápice de arrepentimiento. Su blasón fue quebrado ritualmente por un heraldo real y sus bienes confiscados: fue condenado a prisión perpetua en las fosas subterráneas de la fortaleza de Roca-Negra, bajo vigilancia mágica constante para evitar que su pacto demoníaco volviera a manifestarse.",
        "Su padre, el Conde Raymond, no llegó a ver el final del juicio. La gangrena mágica provocada por la maldición de Khorne consumió sus órganos vitales durante la primera semana de custodia; su deceso fue registrado con la pérdida póstuma de todos sus honores nobiliarios.",
        "En cuanto a los mercenarios y cómplices del Capitán Malrik, los tribunales fueron severos pero justos: los cabecillas responsables de secuestros y abusos fueron sentenciados a la horca en la colina del viento, mientras los soldados rasos que no cometieron crímenes de sangre fueron destinados a trabajos forzados para desescombrar los túneles colapsados de las minas bajo la estricta custodia de los caballeros de la Corona.",
        "* * *",
        "Durante esas dos semanas, los Cien Caballeros Nobles, la Sección Uno y los milicianos barrieron palmo a palmo cada rincón del condado.",
        "Se registraron sótanos, graneros abandonados y cavernas boscosas, capturando a más de cuarenta fanáticos remanentes que intentaban huir hacia las fronteras. En la plaza de Fresnedas se levantó una gigantesca pira purificadora donde Faelar y los sacerdotes paladines redujeron a cenizas cada ídolo de bronce, pergamino blasfemo y estandarte de la estrella de ocho puntas, sellando espiritualmente la tierra con bendiciones de luz y viento.",
        "Con la herejía extirpada de raíz, comenzó la verdadera reconstrucción.",
        "Sauceda y Fresnedas renacieron de sus cenizas. Gracias a los fondos recuperados de los cofres privados del conde y a los suministros de emergencia traídos por la caravana de la Corona, se repararon los techos de paja quemados, se distribuyeron sacos de grano para la siembra de invierno y se entregaron herramientas a cada familia campesina.",
        "Borgar Yunque-Férreo lideró los trabajos en las minas del sur: con la maestría de la ingeniería enana, rediseñó los conductos de ventilación, reforzó los puntales de madera con vigas de hierro forjado y eliminó todo peligro de gases tóxicos. Tarek fue nombrado oficialmente Gran Inspector de Minas por decreto condal, garantizando salarios dignos, jornadas justas y un ambiente de trabajo seguro para los mineros que antes eran tratados como ganado.",
        "* * *",
        "En el ámbito económico, la astucia táctica y diplomática dio sus mayores frutos.",
        "Julian, asesorado por Vlad y por mí, se reunió en el castillo con los emisarios del influyente Consorcio Mercantil del Río Sol. Con el respaldo del sello ducal de la Casa de Vlad y la garantía del edicto real del Lord Paladín, firmaron un tratado comercial histórico:",
        "Morvath reanudó de inmediato la exportación legal de cobre, plata y grano con aranceles reducidos, recibiendo a cambio un flujo constante de moneda de oro, medicinas y telas finas. En menos de diez días, las caravanas de comerciantes abarrotaron los muelles fluviales y el mercado central, devolviendo la prosperidad y la autosuficiencia a un pueblo que semanas atrás moría de hambre.",
        "Julian demostró con creces que no era un señor de despacho: recorría las calles a diario vestido con cota de cuero sencilla, escuchando a los artesanos, supervisando las obras de la fragua junto a Garret y llevando a su hermano menor Leo a cada acto civil para que el niño aprendiera que el verdadero deber de un gobernante es servir a su pueblo.",
        "La milicia de campesinos que empuñó horquillas en Fresnedas fue legalizada mediante decreto real bajo el nombre oficial de la **Guardia Cívica de Morvath**: el veterano Sargento Roderick fue ascendido a Comandante General de la plaza, mientras Garret fue nombrado Maestro Mayor de Forjas y Blindajes.",
        "* * *",
        "Al finalizar la segunda semana, el día previo a nuestra partida hacia la capital, la plaza mayor de Morvath fue el escenario de una magna ceremonia pública.",
        "Bajo un cielo despejado de otoño, frente a una multitud de más de dos mil ciudadanos congregados con lágrimas de gratitud y estandartes engalanados, el Lord Paladín de la Corona subió al estrado de honor acompañado por Julian, quien lucía la túnica formal de Regente.",
        "Por orden expresa del Rey de Solaria, se confirieron los máximos reconocimientos militares y académicos a los integrantes de la Sección Uno:",
        "—Por su liderazgo estratégico inquebrantable, su maestría marcial y su valor indomable al arrancar a este pueblo de las garras del Caos —proclamó el Lord Paladín con voz solemne—, la Corona otorga a **Seier Silva** la *Gran Cruz de Plata con Sol Naciente* y el título vitalicio de **Caballero Protector de Morvath**.",
        "La multitud estalló en una ovación atronadora que sacudió los muros de la ciudad. Di un paso al frente y saludé con la mano en el corazón, recibiendo la condecoración con la sobriedad y la firmeza que mi padre Ronald me inculcó.",
        "Vlad fue condecorado con el *Laurel Dorado de la Orden Real*, consolidando el prestigio militar de su linaje ducal. Akame, Faelar, Borgar y Kaelan recibieron las *Medallas al Mérito de Primera Clase de la Corona*, acompañadas de certificados de rango S para sus expedientes académicos en la Academia Imperial.",
        "Un momento especialmente emotivo fue la distinción a Vance von Holst:",
        "El noble menor recibió la *Orden de la Espada Justa*. Al escuchar su nombre ovacionado por plebeyos y soldados por igual, los ojos de Vance se humedecieron de orgullo contenido; todo rastro de arrogancia y resentimiento de su pasado en la Clase A-2 quedó borrado para siempre, reemplazado por el honor genuino de un auténtico caballero.",
        "* * *",
        "Al caer la noche, las campanas tocaron un tañido pausado y reverente.",
        "En el corazón de la plaza, donde antes ardía el pánico de las redadas, se develó un gran monumento de mármol blanco y piedra de río: el **Memorial de los Inocentes y Defensores de Morvath**.",
        "Cientos de velas blancas y lámparas de papel flotaron sobre el agua del estanque comunal. En la piedra estaban grabados los nombres de los campesinos asesinados, de las mujeres y niños desaparecidos durante los tres años de opresión, y de los milicianos leales que dieron su vida defendiendo las trincheras de Sauceda.",
        "Julian, con Leo tomado firmemente de su mano, se arrodilló ante el monumento y derramó una copa de vino de honor sobre la base de granito:",
        "—No olvidaremos su sacrificio —juró Julian en voz alta ante su pueblo—. Mientras la Casa Morvath se mantenga en pie, jamás permitiremos que la sombra de la tiranía vuelva a tocar esta tierra.",
        "Nosotros siete nos mantuvimos en guardia de honor a los lados de la piedra, con las cabezas gachas en señal de respeto absoluto.",
        "* * *",
        "Pero la solemnidad dio paso, horas más tarde, a la fiesta más grande que el condado había visto en medio siglo.",
        "Se encendieron hogueras gigantescas en cada esquina. Los barriles de sidra dorada y cerveza de malta rodaron sin descanso; se asaron jabalíes con romero y miel, y las mesas se colmaron de hogazas de pan blanco recién horneado.",
        "La gente de Morvath nos rodeaba con un cariño abrumador: campesinas ancianas que bendecían mis manos callosas, niños que imitaban los golpes de boxeo de Seier, jóvenes que bailaban con Borgar al son de flautas y tambores enanos, cazadores que brindaban con Faelar y Kaelan, y aldeanos que llenaban la jarra de Vlad y Vance.",
        "Akame comía un pincho de carne asada apoyada contra una columna de piedra, observando el regocijo con una sonrisa serena y cálida que rara vez mostraba.",
        "Me acerqué a ella con dos jarras de sidra:",
        "—Buen trabajo, Akame.",
        "La asesina de ojos carmesíes chocó suavemente su jarra con la mía, bebió un sorbo y contempló a Julian, quien reía alzando a Leo sobre sus hombros entre los vítores de la multitud:",
        "—Valió la pena cada gota de sudor, Seier —susurró ella con suavidad—. Salvamos vidas reales. Esto es lo que significa ser un guerrero.",
        "Asentí en silencio, saboreando el licor dulce y la brisa fresca de la noche.",
        "La misión de largo alcance de la Sección Uno en el Condado de Morvath había concluido con una victoria absoluta. Mañana emprenderíamos el camino de regreso a la Gran Academia de Solaria, no como simples cadetes de segundo semestre, sino como héroes forjados en la fragua de la guerra."
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
