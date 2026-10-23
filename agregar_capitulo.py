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
    "num": 79,
    "header": "Capítulo 79",
    "title": "El Albor del Segundo Año: Los Discípulos de Hinomoto",
    "paragraphs": [
        "El sol matutino bañaba las torres de mármol blanco de la Gran Academia Militar y Mágica de Astraea, encendiendo destellos dorados sobre los estandartes que ondeaban en el Gran Patio de Armas.",
        "Un nuevo ciclo comenzaba. Con la llegada del Segundo Año, los cambios decretados tras la abolición de las castas habían dejado de ser una simple reforma administrativa para convertirse en una realidad vibrante. Las viejas murallas invisibles entre la cuna noble y el origen campesino habían sido demolidas por completo; el mérito, la disciplina y el honor en batalla eran ahora la única moneda de curso legal.",
        "Más aún: las puertas del imperio se habían abierto de par en par. La academia recibía por primera vez a cadetes de reinos aliados y regiones remotas del continente sin distinción de linaje ni especie. Elfos silvanos de las arboledas del norte, artesanos enanos de las ciudadelas montañosas, ágiles semihumanos de las estepas y jóvenes plebeyos que habían ganado becas imperiales por su talento con la espada o el maná llenaban las calzadas adoquinadas, cargando baúles de viaje y uniformes nuevos.",
        "Desde las arcadas superiores del patio este, varios veteranos de la antigua Clase F-2 contemplábamos el bullicio con los codos apoyados en la barandilla de piedra.",
        "Llevábamos sobre los hombros las nuevas charreteras de segundo año: placas de plata bruñida que acreditaban nuestro rango como veteranos de campaña tras las duras operaciones en Morvath y las fronteras. A nuestros diecisiete años, el aura que desprendíamos ya no era la de estudiantes inexpertos que buscaban evitar castigos, sino la de guerreros fogueados cuya reputación precedía cada paso.",
        "—Miren a toda esa carne fresca —comentó Lighter recostado contra una columna, quitándose un guantelete para ajustarse las vendas de lino en los nudillos—. Me recuerda a nosotros cuando cruzamos ese mismo portón... aunque creo que nosotros teníamos caras de tener menos futuro.",
        "A su lado, Ken soltó una carcajada limpia mientras se anudaba la clásica bandana roja en la frente. Había terminado una sesión intensa de calistenia matutina y el vapor ascendía de su piel bronceada:",
        "—Habla por ti, Lighter. Yo desde el primer día sabía que les partiría la cara a los nobles presumidos. Pero hay que admitir que el ambiente se siente distinto. Ya no hay miraditas de asco ni carruajes de oro pisoteando a los reclutas de a pie. Gerald se lució con esta apertura.",
        "—La fuerza de una nación no reside en los árboles genealógicos de sus generales, sino en la voluntad de quienes están dispuestos a sangrar por su gente —dije con calma, apoyando una mano sobre el pomo de mi espada mientras examinaba el flujo de personas en el patio.",
        "Fue en ese instante cuando un grupo peculiar que cruzaba el arco principal detuvo nuestra conversación.",
        "* * *",
        "Eran emisarios y cadetes de intercambio procedentes del lejano **Imperio de Hinomoto**: una mística potencia insular del oriente, famosa allende los mares por sus tradiciones milenarias, sus templos de madera sagrada entre cerezos y, por encima de todo, por albergar las escuelas de artes marciales a mano limpia más despiadadas y refinadas del mundo conocido.",
        "Entre la delegación oriental, dos figuras destacaban con una fuerza magnética irresistible para cualquiera que supiera leer las señales del combate.",
        "La primera era un muchacho de cabello castaño corto, complexión aparentemente modesta y estatura media. Llevaba el uniforme de aspirante holgado y cargaba una mochila de lona repleta de libros gruesos de botánica e historia, además de un petate de entrenamiento sobre el hombro. Caminaba con cierta timidez, encogiéndose de hombros cada vez que la marea de gente lo rozaba y pidiendo disculpas en voz baja con reverencias apresuradas. Su nombre, inscrito en la tablilla de madera de su petate, era **Kenichi Shirahama**.",
        "A su lado caminaba una joven cuya presencia cortaba el aliento: **Miu Fūrinji**. Poseía una deslumbrante melena rubia dorada recogida en una coleta alta que se balanceaba con gracia cadenciosa, ojos azul violáceo llenos de dulzura y una silueta atlética impecable que ni siquiera el uniforme holgado lograba disimular. Pero más allá de su evidente belleza, lo que paralizaba los sentidos era su andar: flotaba sobre los adoquines como un felino ingrávido, sin hacer crujir una sola brizna de hierba, con el centro de gravedad alineado a la perfección en cada paso.",
        "Ken entrecerró los ojos de inmediato. Su sonrisa burlona desapareció, reemplazada por la mirada afilada de un cazador que reconoce a un igual:",
        "—Oigan... miren a ese chico castaño —murmuró Ken, señalando discretamente con la barbilla—. Miren cómo camina.",
        "Lighter se enderezó de la columna, clavando sus ojos experimentados en el andar de Kenichi:",
        "—No apoya los talones. Distribuye el peso en los metatarsos y pivota las caderas de manera imperceptible en cada paso para absorber cualquier impacto imprevisto. Y fíjate en el cuello y los trapecios... ese muchacho parece tímido, pero sus músculos han aguantado más castigo del que un regimiento de infantería soportaría en toda su vida.",
        "Yo cerré los ojos por una fracción de segundo, sincronizando mi respiración para desplegar un barrido sutil de maná y percepción telúrica.",
        "Bajo la tela de su ropa, el cuerpo de Kenichi Shirahama vibraba con una densidad asombrosa. No era el volumen exagerado de quien levanta pesas por vanidad; eran capas superpuestas de fibras musculares comprimidas, tendones de acero y microfisuras óseas soldadas por un condicionamiento físico inhumano. El chico albergaba un núcleo de *ki* concentrado en su vientre bajo, templado en el fuego de un entrenamiento diario tan despiadado que rozaba la tortura.",
        "—No es un novato cualquiera —dije en voz baja—. Ese chico oculta un poder marcial descomunal. Y la chica que lo acompaña... su equilibrio corporal es casi divino.",
        "* * *",
        "Antes de que pudiéramos descender las escalinatas, un pequeño incidente estalló en el centro del patio.",
        "Tres cadetes novatos de complexión pesada —hijos de familias mercenarias adineradas del oeste que aún conservaban aires de matones callejeros— avanzaban a empujones entre la multitud, abriéndose paso a codazos. El líder del trío, un tipo fornido de casi dos metros con un mandoble enfundado a la espalda, chocó con fuerza deliberada contra el hombro de Kenichi, buscando intimidarlo.",
        "El choque hizo caer un par de libros de la mochila del castaño.",
        "—¡Fíjate por dónde caminas, enano debilucho! —bramó el bravucón, plantándose frente a él con el ceño fruncido—. Los campesinos extranjeros deberían arrodillarse antes de estorbar el paso.",
        "La reacción de Kenichi fue cómica y desconcertante. En lugar de encararlo, el chico abrió los ojos desmesuradamente, agitando las manos en el aire presa de un pánico casi cómico:",
        "—¡A-Ah! ¡Lo lamento tanto! ¡Iba distraído mirando los planos de la biblioteca! ¡Perdón, no fue mi intención estorbarle, señor!",
        "Kenichi se agachó apresuradamente para recoger sus tomos, pero el gigantón, envalentonado por la aparente sumisión del joven, soltó una carcajada despectiva y lanzó un manotazo tosco para agarrarlo del cuello de la camisa:",
        "—¡Las disculpas no bastan, forastero! ¡Vas a aprender modales en...!—",
        "Lo que ocurrió a continuación duró menos de un parpadeo.",
        "El cuerpo de Kenichi reaccionó mucho antes de que su mente pudiera entrar en conflicto. En cuanto la mano del agresor entró en su radio letal, los reflejos condicionados por años de castigo extremo tomaron el control absoluto: sin levantarse del todo del suelo, Kenichi giró la cadera en un ángulo milimétrico, desvió la muñeca del gigantón con una caricia circular de pura suavidad del estilo Jujitsu y redirigió la propia fuerza del golpe hacia el vacío.",
        "Con un movimiento fluido como el agua corriente, atrapó el brazo del agresor y dio un paso lateral. El coloso de dos metros perdió por completo el centro de gravedad, sobrevoló la espalda de Kenichi describiendo un arco perfecto en el aire y aterrizó de espaldas contra los adoquines con un estruendo sordo que levantó una nube de polvo.",
        "El hombre quedó tendido en el suelo, sin aliento y con los ojos desorbitados, sin haber sufrido una sola fractura pero completamente desarmado y aturdido por la inercia.",
        "Kenichi, aún sosteniendo su libro en la mano izquierda, miró al sujeto en el suelo y volvió a entrar en pánico cómico, llevándose las manos a la cabeza:",
        "—¡Aaaaah! ¡Lo siento mucho! ¡Mi cuerpo se movió solo! ¡¿Se encuentra bien?! ¡Por favor no me demande, se lo suplico!",
        "A su lado, Miu Fūrinji soltó un leve suspiro resignado, aunque sus ojos brillaban con una sonrisa divertida y tierna:",
        "—Kenichi-kun... los maestros te dijeron muy claramente que debías contener tus reflejos inconscientes. Vas a espantar a los compañeros de clase el primer día.",
        "* * *",
        "Desde el arco de piedra, Ken no pudo contenerse más y soltó un silbido prolongado, saltando la barandilla de tres metros para aterrizar con soltura felina en el patio de armas.",
        "Lighter y yo lo seguimos al instante, bajando con paso firme mientras la multitud se apartaba al reconocer las charreteras y los rostros de los héroes de Morvath. Los otros dos compinches del agresor dieron un paso atrás temblando al vernos aproximar.",
        "—Una técnica de proyección impecable, amigo —dijo Ken cruzándose de brazos frente a Kenichi, mostrando una sonrisa desafiante y entusiasta mientras pequeñas chispas de maná solar chisporroteaban en sus nudillos—. Redirigir el peso de un mastodonte sin usar maná bruto ni partirle un solo hueso... eso no se aprende en manuales de biblioteca.",
        "Lighter sonrió de lado, evaluando la postura del chico:",
        "—Reflejos de defensa absoluta grabados en la médula espinal. Apuesto a que si alguien te arroja una lanza mientras duermes, te despiertas con la punta entre dos dedos.",
        "Kenichi se puso de pie rápidamente, intimidado al principio por la estatura y la presencia física de Ken y Lighter, pero al mirarnos a los ojos, su instinto de artista marcial pareció captar la sinceridad y el respeto que emanábamos.",
        "—G-Gracias... —balbuceó rascándose la nuca con timidez—. Me llamo Kenichi Shirahama. Ella es mi compañera, Miu Fūrinji. Venimos del Imperio de Hinomoto... del **Dojo Fūrinji**.",
        "Miu dio un paso al frente y realizó una elegante y respetuosa reverencia tradicional oriental, con las manos entrelazadas al frente:",
        "—Es un gran honor conocer a los guerreros de Astraea. Hemos escuchado sobre las reformas de la academia y las hazañas de su escuadrón. Mi abuelo, el maestro del dojo, nos envió aquí para aprender sobre las artes de combate occidentales y ampliar nuestros horizontes.",
        "Di un paso adelante y extendí mi mano derecha hacia Kenichi con una sonrisa cálida y firme:",
        "—Mi nombre es Seier Silva, capitán de la antigua Clase F-2 y estudiante de segundo año. Ellos son Ken y Lighter. En esta academia las castas han muerto; no nos importan los títulos de nobleza ni las tierras de donde provengan. Aquí solo valoramos el temple, el esfuerzo diario y la hermandad.",
        "Kenichi tomó mi mano. El apretón fue revelador: su palma era áspera como la corteza de un roble milenario, pero su mirada transmitía una nobleza y una bondad incorruptibles.",
        "—Bienhallados en Solaria, discípulos de Fūrinji —añadí mientras el sol del mediodía iluminaba el patio—. Me parece que este segundo año va a ser mucho más interesante de lo que imaginábamos."
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
