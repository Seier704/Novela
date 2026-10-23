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
    "num": 80,
    "header": "Capítulo 80",
    "title": "El Nuevo Orden Académico: Crisol de Magia, Acero y Destino",
    "paragraphs": [
        "El Gran Anfiteatro Rúnico de la Academia de Astraea vibraba con una energía sobrecogedora, amplificada por el murmullo de más de doscientos cadetes que colmaban sus gradas circulares de mármol y basalto.",
        "Bajo la monumental cúpula de cristal encantado, donde los rayos del mediodía se refractaban en destellos violetas y esmeraldas, se congregaba la totalidad del Segundo Año. La estampa era histórica: veteranos curtidos de la antigua Clase F-2 y de la disuelta A-1 compartían bancos con los nuevos cadetes becados, forasteros de las estepas, artesanos enanos y los emisarios del lejano oriente. Las charreteras de plata bruñida brillaban sobre los hombros de los veteranos de diecisiete años, proyectando una presencia de peso marcial ineludible.",
        "En el estrado central, flanqueado por caballeros rúnicos de la guardia imperial y maestros arcanos, el Instructor Jefe Gerald dio un paso al frente. Su armadura de cuero reforzado con placas de mithril resonó con un chasquido seco que impuso un silencio absoluto en el recinto.",
        "—¡Cadetes de Segundo Año! —tronó la voz de Gerald, amplificada por las runas de resonancia del púlpito—. El año pasado los vimos ingresar divididos por murallas absurdas de cuna y sangre. Esas murallas fueron demolidas por la realidad de la guerra. Su despliegue en las fronteras y en el feudo de Morvath demostró una verdad inapelable: la complacencia engendra cadáveres, y la rigidez teórica es el mayor pecado de un oficial.",
        "Gerald recorrió las gradas con su mirada penetrante de halcón antes de continuar con tono severo:",
        "—Por orden directa de la Corona y del Alto Mando, el Segundo Año no tolerará especialistas ciegos. Se acabó la separación entre la 'División Marcial' y la 'División Mágica'. Un espadachín que no comprende las matrices de dispersión arcana es despedazado por el primer conjuro de área en el campo de batalla; y un mago que entra en pánico cuando un enemigo se le acerca a menos de tres metros es un estorbo para su escuadrón. A partir de hoy, entra en vigor el **Sistema de Formación Integral y Sinergia Táctica**.",
        "Un murmullo de asombro y expectativa recorrió el anfiteatro. Gerald alzó su guantelete hacia el gigantesco cristal rúnico que flotaba en el centro de la sala, el cual estalló en un resplandor dorado revelando las listas de las nuevas Divisiones Tácticas:",
        "—Todas las asignaturas y maniobras de este semestre serán combinadas e interdisciplinarias —proclamó Gerald—. Magia de combate aplicada, transmutación de blindajes, barreras de choque, alquimia de asedio, esgrima rúnica y condicionamiento físico extremo a puño limpio. Nadie se graduará como oficial de Astraea sin dominar tanto el acero como el maná.",
        "* * *",
        "Los nombres de los estudiantes comenzaron a destellar en la superficie pulida del cristal, organizando a los cadetes en divisiones equilibradas donde las habilidades se complementaran sin contemplaciones.",
        "**Kenichi Shirahama** estiró el cuello desde las gradas intermedias para buscar su nombre, ajustándose la correa de su petate con dedos temblorosos. Al ver la lista de su aula —la **División Táctica Beta**—, el pobre muchacho palideció como si acabara de ver a la muerte en persona:",
        "Estaba asignado junto a **Seier Silva**, **Erza Scarlet**, **Natsuki Subaru**, **Agnes Tachyon** y **Dante Vance**.",
        "—¡A-Ay, no puede ser...! —gimió Kenichi llevándose las manos al rostro en un ataque de cómica desesperación—. ¡Esa chica de la armadura escarlata tiene exactamente la misma mirada sanguinaria que el maestro Sakaki cuando me persigue con un tronco al hombro! ¡Y esa chica de la bata blanca está mirando mis articulaciones como si quisiera disecarme en un matraz! ¡Mi esperanza de vida acaba de reducirse a tres semanas!",
        "Subaru, sentado dos filas más abajo, se dio la vuelta y le dio una palmada reconfortante en el hombro, sonriendo con complicidad:",
        "—Tranquilo, compañero de Hinomoto. Yo sobreviví un mes entero durmiendo junto a los explosivos de Agnes y las patadas de Caesar. Si no te asustas de las explosiones de azufre al desayuno, nos llevaremos de maravilla.",
        "—Y en cuanto a Erza —añadí con una sonrisa serena inclinándome hacia Kenichi—, su disciplina puede parecer aterradora, pero es el escudo más leal que jamás tendrás a tu lado. Si sobrevives a sus sesiones matutinas, ningún monstruo del reino podrá tocarte.",
        "Kenichi tragó saliva sonoramente, pero en el fondo de sus ojos castaños brilló esa chispa inquebrantable de tenacidad que solo los discípulos forjados en el dolor poseen.",
        "* * *",
        "A pocos metros de distancia, el cristal reveló la conformación de la **División Táctica Alfa**, donde **Miu Fūrinji** fue asignada junto a **Kaoruko Waguri**, **Ken**, **Caesar King**, **Vlad** y **Piper Wheel**.",
        "Caesar soltó una carcajada vibrante que retumbó en la bóveda, poniéndose de pie para palmear con entusiasmo la barandilla de piedra:",
        "—¡Ja! ¡Vaya cuadro nos ha tocado! ¡Una bailarina marcial del oriente, la señorita de las barreras sagradas y un loco que pelea a puño limpio envuelto en llamas! ¡Esa aula va a tener más pegada que un ariete de asedio!",
        "Kaoruko sonrió con infinita dulzura hacia Miu, dedicándole una reverencia amable con la cabeza:",
        "—Será un gran honor aprender de tus movimientos, Miu-san. He oído que las técnicas de Hinomoto fluyen con la armonía de la naturaleza.",
        "—El honor es mío, Kaoruko-san —respondió Miu con una sonrisa radiante y ojos violetas llenos de respeto sincero—. Mi abuelo siempre dice que la fuerza que protege a los demás sin derramar sangre es la más excelsa de todas las disciplinas.",
        "Ken, con su bandana roja encendida, se crujió los nudillos con ganas contenidas mientras miraba a Miu y luego a Kenichi en la otra división: la promesa de intercambiar golpes y estilos con herederos de un dojo legendario lo tenía al borde de la euforia marcial.",
        "* * *",
        "El silencio volvió a adueñarse de la sala cuando el Instructor Gerald alzó nuevamente la voz, adoptando un tono de gravedad solemne que tensó el ambiente:",
        "—Presten atención a lo que voy a decir, porque esto define el verdadero propósito de su Segundo Año.",
        "Gerald bajó los escalones del púlpito, plantándose en el centro de la arena rúnica:",
        "—Este entrenamiento conjunto de magia y combate no es para que presuman en torneos ni para inflar sus calificaciones académicas. En cuanto concluya el primer trimestre de acondicionamiento, cada una de estas divisiones se subdividirá en **Escuadrones de Despliegue de Alto Riesgo** para misiones directas de la Corona en los rincones más hostiles del continente.",
        "La advertencia cayó con el peso del plomo sobre las gradas.",
        "—Hablamos de brechas de fallas arcanas activas, ruinas olvidadas de la Primera Dinastía, migraciones de bestias de pesadilla y amenazas fronterizas donde el peligro es real, despiadado y mortal —continuó el instructor jefe sin parpadear—. Pero escúchenme bien: ustedes ya no son los reclutas asustadizos de primer año. Son veteranos formados de diecisiete años que han vencido al culto, han derrotado a los Skavens y han sobrevivido a las ciénagas y desfiladeros del reino. Tienen el poder, la disciplina y el temple. Si aplican esta sinergia y confían en sus compañeros, no habrá peligro en este mundo capaz de doblegarlos.",
        "Un murmullo de determinación colectiva sacudió las filas. No había miedo en los ojos de mis hermanos; solo la certeza serena de quienes ya habían mirado a los ojos de la muerte y habían prevalecido.",
        "* * *",
        "Para sellar la inauguración del curso, Gerald convocó una sesión de calentamiento y demostración táctica sobre el suelo de basalto reforzado del anfiteatro.",
        "—¡Aspirante Shirahama! —llamó Gerald con voz de trueno—. ¡Al centro de la arena!",
        "Kenichi dio un respingo cómicamente desmesurado, soltando sus libros sobre el banco antes de tropezar un poco al bajar las gradas, agitando las manos avergonzado:",
        "—¡¿Y-Yo?! ¡P-Pero si yo apenas sé hacer reverencias, señor instructor!",
        "La consigna era clara: una prueba de reflejos y evasión frente a matrices de proyectiles mágicos no letales y dos cadetes de segundo año armados con bastones de choque rúnico. Cuatro orbes de entrenamiento flotaron en el aire, comenzando a disparar ráfagas consecutivas de energía balística hacia Kenichi a gran velocidad, mientras los dos guardias cargaban desde los flancos.",
        "Lo que siguió dejó a todo el anfiteatro en un silencio sepulcral.",
        "En el instante en que el primer rayo de maná rozó el aire junto a su sien, el aspecto tímido de Kenichi se desvaneció por completo. Sus ojos castaños se afilaron como hojas de afeitar; su cuerpo entró en el estado de supervivencia absoluta que solo los infiernos de Ryōzanpaku podían forjar.",
        "Sin retroceder un solo paso, Kenichi ejecutó el *Seikūken*: una esfera defensiva perfecta de micro-movimientos milimétricos. Esquivaba los proyectiles mágicos ladeando la cabeza por fracciones de pulgada, girando el torso con suavidad como si el aire mismo lo guiara. Cuando los dos guardias intentaron asestarle sendos golpes con los bastones, Kenichi desvió ambos impactos simultáneamente con las palmas abiertas, cruzó los brazos en una rotación impecable y desarmó a los atacantes en un parpadeo, dejándolos mirando sus manos vacías mientras las chispas de los proyectiles estallaban inofensivas contra el suelo a su espalda.",
        "Kenichi parpadeó, miró a su alrededor y volvió a encogerse de hombros con una risita nerviosa:",
        "—¡L-Lo siento! ¡No quise romperles la postura!",
        "Desde las gradas, Ken soltó una carcajada atronadora mientras aplaudía con fuerza:",
        "—¡Ese es el espíritu, novato de Hinomoto! ¡Ni siquiera tocó el maná y los dejó en ridículo!",
        "Erza Scarlet asintió con gravedad marcial, cruzándose de brazos con una mirada de sincero respeto:",
        "—Reflejos extraordinarios. Con la instrucción adecuada sobre matrices arcanas, será un baluarte invencible en la vanguardia.",
        "Apoyé la espalda contra el asiento, cruzando miradas con Vlad, Akame y Kaoruko. Una profunda satisfacción me llenó el pecho. La Gran Academia de Astraea había dejado atrás el pasado oxidado de los linajes para convertirse en una forja implacable de guerreros.",
        "El Segundo Año apenas comenzaba, pero la llama de nuestra hermandad ardía con más fuerza que nunca, lista para afrontar cualquier misión que el reino nos pusiera por delante."
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
