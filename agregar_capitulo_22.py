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

CAPITULO_22 = {
    "header": "Capítulo 22",
    "title": "El Llamado de la Arena",
    "paragraphs": [
        "Las semanas posteriores al viaje a las aguas termales transcurrieron a un ritmo implacable dentro de la Academia Real de Magia de Solaria.",
        "El descanso de las termas había recargado el espíritu de la Clase F-2, pero los instructores no tardaron en recordarnos que este lugar no era un hogar de retiro. Los entrenamientos de maná elemental se intensificaron, las clases de teoría militar llenaban pizarras enteras con esquemas de invasión bestial y el acondicionamiento físico nos dejaba a todos con los músculos ardiendo al final de cada jornada.",
        "En ese torbellino de sudor y disciplina, cada miembro del grupo continuaba afinando sus armas: Subaru Natsuki, a pesar de sus quejas constantes, iba ganando resistencia y control en su incipiente maná de oscuridad bajo el severo tutelaje de Yami; Subaru Hoshina se movía con un juego de pies mucho más fluido y seguro, compartiendo a menudo charlas tranquilas con Lighter y Kaoruko; y yo continuaba perfeccionando la integración de mis golpes de Muay Thai y Jeet Kune Do con la densidad aplastante del fuego y la tierra.",
        "Sin embargo, todo aquel régimen de preparación tenía una razón de ser.",
        "Un motivo de proporciones épicas que la academia se encargó de revelar una mañana de sol radiante mediante el tañido solemne de los grandes cuernos de guerra de la capital.",
        "* * *",
        "Por orden real, todo el cuerpo estudiantil fue convocado al Gran Coliseo Central de Solaria.",
        "El imponente anfiteatro de mármol blanco y piedra volcánica, con capacidad para miles de espectadores, hervía de expectación. En las gradas superiores se congregaban los nobles de las Clases A1, A2 y B, luciendo uniformes inmaculados con ribetes dorados y mirándonos por encima del hombro con su arrogancia habitual. En la arena principal, las distintas clases formaban en filas rigurosas.",
        "El Instructor Jefe, un veterano de las guerras fronterizas con cicatrices en el rostro y una armadura ceremonial reluciente, subió al estrado de piedra y alzó la mano, acallando el murmullo ensordecedor del coliseo.",
        "—¡Atención, cadetes de Solaria! —resonó su voz amplificada por un artefacto de maná sónico—. El primer semestre entra en su fase cumbre. La teoría y los ejercicios controlados quedan atrás. Ha llegado el momento de medir el verdadero filo del acero y el poder de la magia bajo los ojos del reino.",
        "Una ola de murmullos electrizó las gradas.",
        "—¡Queda formalmente inaugurado el Gran Torneo Interclases! —anunció el oficial con solemnidad marcial—. Las reglas son inflexibles: cada clase seleccionará y registrará a sus cuatro mejores peleadores. Nadie más podrá pisar la arena.",
        "El instructor clavó su mirada severa sobre nosotros antes de desgranar la estructura del torneo.",
        "—La competencia se disputará por puntos a lo largo de cuatro modalidades estratégicas: combate individual uno contra uno, duelo por parejas dos contra dos, choque de asalto tres contra tres, y la batalla campal de escuadrones cuatro contra cuatro. La clase que demuestre la mejor versatilidad, cohesión y poder destructivo será coronada campeona del semestre, con acceso a fondos de armamento imperial y méritos de ascenso militar directo.",
        "El coliseo estalló en vítores y aplausos enardecidos.",
        "A unos metros de nuestra formación, entre las filas aristocráticas de la Clase A-1, divisé a Cedric y al grupo de nobles que solían merodear a su alrededor. Sus rostros reflejaban una mueca de burla evidente; para ellos, que una clase de 'desechos' como la F-2 hubiera aprobado los exámenes teóricos era un chiste de mal gusto que estaban ansiosos por corregir a golpes en la arena.",
        "Sin embargo, más allá de los engreídos habituales, mis ojos se cruzaron con otra figura.",
        "Elysea de Vane, la elfa de hielo y capitana de la Clase A-2, permanecía erguida con su uniforme azul marino y su espada delgada en la cintura. Sus ojos celestes, fríos como la ventisca, se posaron directamente sobre mí. Cuando nuestras miradas se conectaron a través del foso de arena, la noble elfa ladeó la cabeza y me dedicó una media sonrisa sutil, desafiante y cargada de una expectación peligrosa.",
        "«Esto se va a poner interesante», pensé para mis adentros, sintiendo cómo la sangre comenzaba a hervir en mis venas.",
        "* * *",
        "En cuanto regresamos a nuestro salón de clases, el orden se disolvió en un segundo.",
        "—¡CUATRO PLAZAS! —bramó Ken de un salto sobre su pupitre, lanzando un derechazo al aire con una chispa dorada brotando de sus nudillos—. ¡Una de esas es mía! ¡Voy a destrozar a esos nobles presumidos en el uno contra uno hasta que aprendan a respetar a la F-2!",
        "—¡Cállate, idiota ruidoso! —le gritó Lucy agitando su vara de mando—. ¡No puedes saltar a la arena solo con ganas de romper cabezas! ¡Esto requiere estrategia, formación defensiva y disciplina táctica!",
        "—La capitana enana tiene razón en algo —intervino Erza, cruzada de brazos con una expresión tan solemne que imponía respeto a toda la habitación—. Quienes representen a nuestra clase cargarán con el honor de todos nosotros. No podemos permitirnos titubear. Si alguien sale a ese coliseo, debe estar dispuesto a dejar la piel.",
        "—A mí no me miren —bostezó Yami, apoyando los pies sobre la mesa mientras sacaba un cigarrillo y lo encendía con la chispa de su encendedor—. Demasiado papeleo ceremonial y reglas estúpidas. Si no hay recompensa en monedas de plata ni buen tabaco de por medio, prefiero dormir una siesta en las gradas mientras cuido que el mocoso de Natsuki no se atragante con los nervios.",
        "—¡Oye! ¡¿Por qué siempre me usas de ejemplo para cosas patéticas?! —chilló Natsuki con las manos en la cabeza, arrancando carcajadas en la fila trasera.",
        "Caesar King soltó una risa profunda y se cruzó de brazos, ajustándose la corona metálica.",
        "—En Calydon nos encanta una buena paliza pública —dijo la jefa de los motociclistas—. Pero Lucy y Piper necesitan apoyo logístico en las gradas para analizar las tácticas enemigas. Además, nosotros tenemos al mejor fajador del reino.",
        "Caesar le dio una palmada en el omóplato a Lighter, quien estaba recostado contra la pared junto a la ventana, jugando distraídamente con su chisquero de latón.",
        "—Si hay que subir al ring a callar bocas arrogantes, cuenten conmigo —dijo Lighter con una sonrisa relajada—. Mis puños están listos.",
        "Subaru Natsuki se aclaró la garganta y se colocó frente a la pizarra con una tiza entre los dedos, asumiendo su rol como analista táctico del grupo.",
        "—A ver, bestias del combate, escuchen al cerebro de la operación —dijo Natsuki dibujando cuatro casillas en el tablero—. Tenemos modalidades de 1v1, 2v2, 3v3 y 4v4. Eso significa que no solo necesitamos monstruos individuales; necesitamos combatientes capaces de complementarse a la perfección en cualquier formación.",
        "Erza dio un paso al frente y señaló la primera casilla con firmeza indiscutible.",
        "—El primer puesto no está en discusión. Silva debe liderar el equipo.",
        "Toda la clase asintió al unísono. Ken alzó el pulgar, Renji y Akame asintieron en silencio y Kaoruko me miró con una sonrisa radiante llena de orgullo sincero.",
        "—Silva es un muro inquebrantable —argumentó Erza—. Su potencia física, su resistencia al impacto y su dominio dual de fuego y tierra lo convierten en el ancla perfecta tanto para el combate individual como para la primera línea de choque.",
        "—Acepto —dije yo, dando un paso al frente con calma—. Dejaré el cuerpo en esa arena.",
        "—El segundo puesto le corresponde a Erza —propuso Natsuki de inmediato, escribiendo su nombre—. Es nuestra guerrera más versátil: armaduras mágicas para resistir asedios, dominio de cualquier arma blanca y capacidad de mando táctico en el fragor de la batalla.",
        "—Será un honor —respondió Erza con una reverencia marcial impecable.",
        "Quedaban dos puestos libres.",
        "La mirada de todos se repartió entre los contendientes restantes. Ken ardía de entusiasmo; Lighter sonreía con confianza tranquila; Renji y Akame mantenían la serenidad de guerreros experimentados.",
        "—Para el tercer puesto necesitamos movilidad, fuego ofensivo y presión continua —analizó Natsuki—. Alguien que rompa las líneas enemigas desde el aire y desorganice las formaciones mágicas.",
        "—¡Ese soy yo! —gritó Ken, frotándose los nudillos contra la palma abierta—. ¡Las técnicas del Dragón Ígneo van a quemar a quien se ponga enfrente!",
        "—No lo hace nada mal cuando no tropieza con su propia lengua —bromeó Burnice desde atrás, dándole un empujoncito animado—. ¡Ken-kun es una máquina de demolición!",
        "—De acuerdo. Ken ocupará la tercera plaza —dictaminó Erza, aprobándolo con la mirada.",
        "Natsuki anotó el nombre de Ken y miró la última casilla vacía en la pizarra.",
        "—La última plaza es crítica —advirtió Natsuki con seriedad—. Ya tenemos dos tanques de impacto frontal como Silva y Erza, y un asaltante aéreo explosivo como Ken. Nos falta un elemento clave: velocidad pura de contragolpe, lectura de ritmo y alguien capaz de desarmar especialistas veloces sin desgastarse.",
        "Todas las miradas se posaron sobre Lighter.",
        "El boxeador de los Hijos de Calydon empujó sus gafas de sol hacia la frente con un dedo y se despegó de la pared con pasos tranquilos.",
        "—Si están buscando un contragolpeador que apague las luces de los niños ricos antes de que terminen de conjurar sus hechizos... no busquen más —dijo Lighter con una sonrisa serena.",
        "A un lado del aula, Subaru Hoshina lo observó en silencio. Por primera vez, no había escepticismo ni desdén en su mirada verde; al contrario, un brillo tenue de confianza y apoyo silencioso se dibujó en sus ojos.",
        "—No te descuides ahí afuera, boxeador —dijo Hoshina en voz baja, con un tono que solo los más cercanos alcanzaron a escuchar—. Asegúrate de que esos tipos sientan lo que es un verdadero juego de pies.",
        "Lighter la miró de reojo, le dedicó una inclinación de cabeza cómplice y sonrió.",
        "—Descuida, señorita hielo. Haré que valga la pena.",
        "* * *",
        "Con los cuatro nombres sellados en la pizarra, el equipo oficial de la Clase F-2 quedó constituido:",
        "1. **Seier Silva** — Vanguardia pesada, impacto físico demoledor y ancla elemental.",
        "2. **Erza Scarlet** — Comandante de armas, armadura de asedio y versatilidad total.",
        "3. **Ken** — Ofensiva aérea explosiva y artes marciales del Dragón Ígneo.",
        "4. **Lighter** — Boxeo de élite, velocidad de contragolpe y lectura milimétrica de combate.",
        "Una alineación formidable. Cuatro estilos completamente distintos pero que, combinados en parejas de dos, tríos de asalto o en escuadrón completo de cuatro, formaban una máquina de combate aterradora.",
        "Kaoruko se acercó a mi pupitre, sosteniendo un pequeño frasco de ungüentos y una bolsita de té vigorizante que había preparado con sus propias manos.",
        "—Seier-kun... por favor, prométeme que no te sobreexigirás más de la cuenta —me pidió con su dulce voz teñida de preocupación cariñosa—. Voy a tener listos todos los remedios curativos en la enfermería de campaña para cuando terminen sus combates.",
        "—Estaré bien, Kaoruko —le aseguré con una sonrisa tranquila, colocando una mano sobre su hombro delicado—. Esta vez no solo peleamos para aprobar un examen. Vamos a demostrarle a todo el reino de qué está hecha la Clase F-2.",
        "La chica azabache se sonrojó levemente ante mi contacto y asintió con una determinación radiante.",
        "A nuestro alrededor, el resto de los compañeros se volcó en promesas de apoyo: Rukia y Renji se ofrecieron para estudiar las debilidades de los rivales desde las gradas; Akame preparó las piedras de afilar para las armas de Erza; y Lucy organizó una pizarra completa con cronogramas de calentamiento.",
        "Miré a mis tres compañeros de equipo: Erza ajustando las correas de sus guanteletes, Ken saltando de emoción con los ojos encendidos y Lighter vendándose los nudillos con calma quirúrgica.",
        "Los nobles de Solaria creían que este torneo sería su escenario para humillar a los plebeyos y marginados.",
        "No tenían la menor idea de la tormenta de fuego, hierro y puños que estaba a punto de caer sobre el coliseo."
    ]
}

def compilar_capitulo_22():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 22              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 22
    txt_cap22 = os.path.join(capitulos_dir, "22_Capitulo_22.txt")
    with open(txt_cap22, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_22['header'].upper()}: {CAPITULO_22['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_22['paragraphs']))
    print(f"[OK] Texto plano Capítulo 22 guardado en:\n  {txt_cap22}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 22
    docx_cap22 = os.path.join(capitulos_dir, "22_Capitulo_22.docx")
    create_single_chapter_document(CAPITULO_22, docx_cap22, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 22 guardado en:\n  {docx_cap22}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 22)
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
    compilar_capitulo_22()
