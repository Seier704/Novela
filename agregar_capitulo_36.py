# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

CAPITULO_36 = {
    "header": "Capítulo 36",
    "title": "Plaga Subterránea",
    "paragraphs": [
        "El avance por el Segundo Estrato de las Cavernas Carmesíes nos condujo, tras cuatro horas de marcha subterránea, a la Gran Galería de los Pilares.",
        "Era una caverna colosal de más de cien metros de ancho, sostenida por gigantescas columnas naturales de basalto negro que se perdían en las tinieblas del techo. Por el centro de la estancia cruzaba un puente natural de piedra sobre una falla volcánica de la que emanaban ríos de lava incandescente, iluminando la bóveda con un resplandor cobrizo y sofocante.",
        "Allí, por designación del mando militar del fuerte, varios escuadrones mixtos habíamos convergido para cotejar mapas y reagrupar suministros: nuestro Escuadrón 7 se encontraba junto al Escuadrón 5 de Lighter y Subaru Hoshina, y el Escuadrón 2 liderado por la implacable vanguardia de Erza Scarlet y Akame.",
        "Los cadetes de dieciséis años descansaban sobre las piedras, bebiendo agua y curando rasguños menores. Todo parecía bajo control.",
        "Pero la tranquilidad en las entrañas del imperio era un lujo efímero.",
        "* * *",
        "El primer indicio no fue un rugido de bestia.",
        "Fue un hedor repugnante que atravesó incluso la cúpula de purificación de Kaoruko: una pestilencia nauseabunda a orina rancia, pelo podrido, grasa quemada y cloaca profunda.",
        "—Por los dioses del hielo... ¿qué es esa inmundicia? —gruñó Freya Skuld, arrugando la nariz mientras apretaba su lanza pesada.",
        "—No es miasma natural —advirtió Dante Vance en un susurro cortante, cargando el percutor de su mosquete rúnico—. El sensor térmico de mi mira está volviéndose loco. Hay cientos de fuentes de calor diminutas moviéndose detrás de las paredes.",
        "A pocos metros de nosotros, Subaru Hoshina dio un salto hacia atrás con sus dagas en guardia, señalando las fisuras de basalto en la parte superior del túnel.",
        "—¡Están en la roca! ¡Están cavando desde adentro!",
        "De pronto, un sonido aterrador y multitudinario comenzó a resonar por toda la galería. No era el crujido de la piedra, sino el chasquido frenético de millones de garras afiladas rascando las paredes, acompañado por un coro histérico y estridente de chillidos agudos que helaron la sangre de los cadetes más jóvenes:",
        "¡Chik-chik-chik! ¡SKREEEEE! ¡Mata-mata hombres-cosa!",
        "—¡A LAS ARMAS! —bramó Erza con su voz de mando de general, desenvainando dos espadas en un parpadeo.",
        "* * *",
        "Las paredes de la galería reventaron como cascarones de huevo.",
        "De las grietas, conductos de ventilación y fisuras de las columnas brotó una marea viviente y pestilente que cayó en cascada sobre el suelo de piedra.",
        "Eran **Skavens**.",
        "Hombres rata de complexión encorvada y sarnosa, cubiertos de pelaje grasiento y andrajos podridos. Sus ojos rojos y desorbitados brillaban con una locura febril en la penumbra. Aquella horda no estaba compuesta por guerreros de élite con armaduras rúnicas ni hechiceros de plaga; eran la casta más baja y numerosa de su especie: **Esclavos Skaven**.",
        "Cientos de criaturas esqueléticas y desesperadas, empujadas como carne de cañón dispensable por el hambre y el terror a sus amos ocultos, armadas con lanzas de madera astillada, cuchillos herrumbrosos y hondas de cuero.",
        "—¡¡¡CARNE FRESCA-COMIDA!!! ¡¡¡MATAR-CHIRRIAR HOMBRECITOS!!! —chillaban en su dialecto quebrado, abalanzándose en una estampida ciega que cubrió el suelo de la caverna en cuestión de segundos.",
        "Varios cadetes de las Clases B y C entraron en pánico ante la grotesca masa que se les venía encima.",
        "—¡MANTENGAN LA LÍNEA! —rugí con toda la fuerza de mis pulmones, dando tres pasos al frente para interponerme entre la horda y los heridos.",
        "Freya se plantó a mi lado, golpeando el suelo con su escudo de roble reforzado:",
        "—¡VENGAN POR MÍ, RATAS SARNOSAS! —bramó la norteña.",
        "* * *",
        "El choque fue un torbellino salvaje de sangre y acero.",
        "Los primeros veinte esclavos skaven se arrojaron sobre nosotros como una ola de dientes y picas oxidadas.",
        "Canalicé mi maná de **Tierra** a través de las caderas y barrí el aire con una patada baja de Muay Thai con la tibia de hierro: el impacto destrozó las patas y lanzas de los primeros cinco roedores en una explosión de astillas y huesos rotos. Sin frenar el impulso, entré en su masa compacta con una combinación demoledora de codos y rodillazos de Jeet Kune Do, pulverizando cráneos deformes y mandíbulas sarnosas con cada golpe.",
        "A mi derecha, Freya era un molino de guerra: su lanza de trueno descargaba ráfagas eléctricas continuas que freían a tres ratas por estocada, mientras su escudo aplastaba pechos esqueléticos contra la roca.",
        "A nuestro flanco izquierdo, Lighter y Subaru Hoshina demostraron una sincronización impecable.",
        "Lighter se movía como un boxeador en una reyerta callejera: deslizaba el torso bajo las lanzas enemigas y soltaba directos y ganchos quirúrgicos que reventaban cabezas de ratas en el aire, mientras Hoshina giraba a su alrededor como un torbellino de plata, degollando a los esclavos que intentaban trepar por sus piernas.",
        "Más atrás, Erza y Akame cortaban la marea con la precisión de guadañas en un campo de trigo: un destello de hojas plateadas y carmesíes dejaba pilas de cadáveres cercenados a cada segundo.",
        "Y desde la retaguardia, Elysea de Vane trazaba senderos de escarcha profunda que congelaban las patas de las criaturas, mientras Dante Vance disparaba perdigones de viento rúnico con su mosquete que barrían a grupos enteros hacia el abismo de lava.",
        "Los esclavos skaven morían por decenas.",
        "Eran débiles, cobardes y malnutridos.",
        "Pero el problema era que no dejaban de salir.",
        "* * *",
        "Por cada rata que aplastábamos contra la piedra, tres más saltaban de las grietas del techo, cayendo sobre nuestras armaduras como una lluvia hedionda.",
        "Y peor aún: el peso descomunal de la horda y las vibraciones de sus túneles apresurados comenzaron a resquebrajar la bóveda de la Gran Galería.",
        "¡CRACK! ¡KRAAAAASH!",
        "Gigantescas estalactitas de basalto de varias toneladas se desprendieron del techo, estrellándose contra el suelo y amenazando con sepultar a todos los estudiantes bajo una avalancha de roca volcánica.",
        "—¡SILVA! —gritó Erza a través del estruendo de la batalla—. ¡LA CUEVA NO VA A RESISTIR! ¡SI NOS QUEDAMOS A PELEAR EN CAMPO ABIERTO, EL TECHO NOS VA A APLASTAR A TODOS!",
        "Miré hacia el frente.",
        "Miles de ojos rojos seguían brillando en las profundidades de los túneles. Los esclavos skaven no tenían fin porque estaban siendo empujados desde atrás para desgastarnos.",
        "Combatir hasta la última gota de maná contra carne de cañón desechable en un terreno que colapsaba era una sentencia de muerte estúpida.",
        "—¡RETIRADA TÁCTICA! —ordené con voz atronadora—. ¡TODOS HACIA EL PUENTE DE PIEDRA! ¡EVACUEN AL PRIMER ESTRATO AHORA!",
        "La disciplina de la Clase F-2 se impuso de inmediato sobre el pánico.",
        "Kaoruko desplegó una barrera masiva de agua bendecida que empujó a la primera línea de ratas hacia atrás, mientras Dante guiaba a los cadetes heridos y asustados de las otras clases cruzando el puente natural hacia el túnel de escape.",
        "* * *",
        "La retaguardia final quedó en manos de los más experimentados.",
        "Erza, Lighter, Freya y yo nos apostamos en la boca del puente de piedra, formando un embudo infranqueable de diez metros de ancho.",
        "Los esclavos skaven se amontonaron como una marea frenética intentando cruzar la pasarela sobre la lava, chillando y pisoteándose entre ellos con tal de llegar a nosotros.",
        "—¡AGUANTEN DIEZ SEGUNDOS MÁS! —rugió Freya, empujando con su escudo una montaña de cuerpos chillones.",
        "Elysea desató un muro colosal de estacas de hielo que frenó la estampida en seco durante tres latidos.",
        "—¡Ya cruzaron todos! —gritó la voz de Dante desde el otro extremo de la sima volcánica—. ¡Silva, vuelen el puente ahora!",
        "—¡MUÉVANSE! —les grité a Erza, Freya y Lighter.",
        "Los tres saltaron hacia atrás cruzando el tramo final del puente en un segundo.",
        "Me quedé solo en el centro de la pasarela de roca.",
        "Cientos de esclavos skaven treparon sobre las estacas de hielo de Elysea, abalanzándose hacia mí con sus dientes afilados y ojos desorbitados por la sangre.",
        "Cerré los puños. Canalicé todo mi maná de **Tierra** hacia mis brazos y descargué un golpe de pulgada con ambos puños cerrados directo contra la base de las losas centrales del puente:",
        "—¡¡¡RUPTURA SÍSMICA!!!",
        "¡¡¡¡KRAAAAAAAAAAAAAAAAA-BOOOOOOOOOM!!!!",
        "La falla geológica respondió al impacto. El puente de piedra se partió en tres secciones colosales que se precipitaron hacia el abismo de lava ardiente.",
        "Salté hacia atrás en el último milisegundo, rodando sobre el suelo firme del túnel de escape mientras centenares de esclavos skaven caían al vacío entre chillidos agónicos, consumidos por el fuego líquido del magma.",
        "Dante activó de inmediato una carga de pólvora rúnica en la entrada del pasadizo, provocando un derrumbe controlado que selló la galería detrás de nosotros con toneladas de granito y polvo negro.",
        "* * *",
        "El silencio regresó de golpe, roto únicamente por las respiraciones agitadas y las toses ahogadas de los estudiantes.",
        "Estábamos a salvo en el corredor fortificado de subida hacia el Primer Estrato.",
        "Ningún cadete había muerto, pero los uniformes estaban rasgados, los escudos abollados y el hedor a rata quemada impregnaba el aire.",
        "Kaoruko se arrodilló a mi lado con lágrimas en los ojos, pasándome un frasco de agua y revisando mis nudillos con devoción temblorosa.",
        "—Seier-kun... estás bien... gracias a los cielos...",
        "—Estoy entero, Kaoruko —le aseguré con una sonrisa cansada, apoyando la espalda contra el muro de roca mientras recuperaba el aire.",
        "A unos metros, Erza, Elysea, Lighter y Dante se reunieron en círculo con rostros sombríos.",
        "—Esos no eran monstruos nacidos del miasma —dijo Elysea con tono gélido, limpiando la sangre negra de su estoque con un pañuelo de seda—. Eran Skavens. Una plaga subterránea organizada.",
        "—Y solo eran esclavos —añadió Lighter, mirando los escombros sellados con los ojos entrecerrados—. En las Tierras Yermas se dice que cuando los esclavos corren en masa hacia la superficie... es porque algo mucho más monstruoso y letal viene marchando detrás de ellos.",
        "Erza apretó la empuñadura de su espada con tanta fuerza que sus nudillos se pusieron blancos.",
        "La incursión de nivel medio acababa de dejar de ser un simple entrenamiento militar.",
        "Una sombra ancestral e infecta despertaba en las raíces del imperio, y la Clase F-2 estaba en la primera línea de la tormenta que se avecinaba."
    ]
}

def compilar_capitulo_36():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 36              ")
    print("==================================================")
    
    from compilar_novela import PROLOGO, CAPITULO_1
    from agregar_capitulo_2 import CAPITULO_2
    from agregar_capitulo_3 import CAPITULO_3
    from agregar_capitulo_4 import CAPITULO_4
    from agregar_capitulo_5 import CAPITULO_5
    from agregar_capitulo_6 import CAPITULO_6
    from agregar_capitulo_7 import CAPITULO_7
    from agregar_capitulo_8 import CAPITULO_8
    from agregar_capitulo_9 import CAPITULO_9
    from agregar_capitulo_10 import CAPITULO_10
    from agregar_capitulo_11 import CAPITULO_11
    from agregar_capitulo_12 import CAPITULO_12
    from agregar_capitulo_13 import CAPITULO_13
    from agregar_capitulo_14 import CAPITULO_14
    from agregar_capitulo_15 import CAPITULO_15
    from agregar_capitulo_16 import CAPITULO_16
    from agregar_capitulo_17 import CAPITULO_17
    from agregar_capitulo_18 import CAPITULO_18
    from agregar_capitulo_19 import CAPITULO_19
    from agregar_capitulo_20 import CAPITULO_20
    from agregar_capitulo_21 import CAPITULO_21
    from agregar_capitulo_22 import CAPITULO_22
    from agregar_capitulo_23 import CAPITULO_23
    from agregar_capitulo_24 import CAPITULO_24
    from agregar_capitulo_25 import CAPITULO_25
    from agregar_capitulo_26 import CAPITULO_26
    from agregar_capitulo_27 import CAPITULO_27
    from agregar_capitulo_28 import CAPITULO_28
    from agregar_capitulo_29 import CAPITULO_29
    from agregar_capitulo_30 import CAPITULO_30
    from agregar_capitulo_31 import CAPITULO_31
    from agregar_capitulo_32 import CAPITULO_32
    from agregar_capitulo_33 import CAPITULO_33
    from agregar_capitulo_34 import CAPITULO_34
    from agregar_capitulo_35 import CAPITULO_35
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25, CAPITULO_26, CAPITULO_27, CAPITULO_28, CAPITULO_29,
        CAPITULO_30, CAPITULO_31, CAPITULO_32, CAPITULO_33, CAPITULO_34,
        CAPITULO_35, CAPITULO_36
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 36
    txt_cap36 = os.path.join(capitulos_dir, "36_Capitulo_36.txt")
    with open(txt_cap36, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_36['header'].upper()}: {CAPITULO_36['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_36['paragraphs']))
    print(f"[OK] Texto plano Capítulo 36 guardado en:\n  {txt_cap36}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 36
    docx_cap36 = os.path.join(capitulos_dir, "36_Capitulo_36.docx")
    create_single_chapter_document(CAPITULO_36, docx_cap36, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 36 guardado en:\n  {docx_cap36}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 36)
    master_docx = os.path.join(NOVELA_DIR, "compilado", "El_Camino_del_Reencarnado.docx")
    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, master_docx)
        print(f"[OK] Libro Maestro actualizado en:\n  {master_docx}")
    except PermissionError:
        print(f"[AVISO] '{master_docx}' está abierto en Microsoft Word.")
        backup_master = os.path.join(NOVELA_DIR, "compilado", "El_Camino_del_Reencarnado_temp.docx")
        create_novel_document(titulo, subtitulo, autor, capitulos, backup_master)
        print(f"[OK] Guardado en copia auxiliar:\n  {backup_master}")

if __name__ == "__main__":
    compilar_capitulo_36()
