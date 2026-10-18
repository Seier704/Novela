# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_59 = {
    "header": "Capítulo 59",
    "title": "Mareas Doradas: Juegos, Sol y Ecos de Hermandad",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "costa_esmeralda_vacaciones.jpg"),
            "caption": "Ilustración: Seier, Kaoruko, Lortelle y Elysea en la Terraza de Costa Esmeralda"
        }
    ],
    "paragraphs": [
        "El cuarto día de vacaciones en la Costa Esmeralda amaneció con una luz deslumbrante que convertía el mar en un espejo infinito de oro y turquesa.",
        "Me desperté temprano, con el cuerpo descansado y una sensación de vigor renovado que hacía tiempo no experimentaba. Salí descalzo hacia la terraza de madera que conectaba los pabellones de la villa con la playa privada, apoyando los antebrazos sobre la barandilla para respirar el aroma a salitre, brisa limpia y flores tropicales.",
        "Apenas unos minutos después, unos pasos ligeros anunciaron compañía.",
        "Kaoruko apareció con un sombrero de paja decorado con pequeñas flores rosadas y un traje de baño blanco con volantes que resaltaba su silueta dulce y esbelta. Se colocó junto a mí en las escaleras de la terraza, dedicándome una sonrisa radiante y apoyando suavemente una mano en mi brazo.",
        "—Buenos días, dormilón... —susurró con ternura inmensa—. El mar está más hermoso que nunca hoy.",
        "Desde la tumbona acolchada, Lortelle se desperezó como un felino satisfecho, vistiendo su bikini esmeralda y un pareo translúcido que dejaba poco a la imaginación. Guiñó un ojo con picardía y soltó una pequeña risa:",
        "—Parece que el aire de la costa le sienta de maravilla al héroe de la F-2. Aunque con las caminatas nocturnas que se dan algunos por las terrazas, me sorprende que despiertes con tanta energía, Silva.",
        "Junto al barandal norte, Lady Elysea de Vane contemplaba las olas con un dos piezas azul hielo que realzaba su porte aristocrático y su piel blanquecina. Al escuchar el comentario de la pelirroja, un leve rubor tiñó las mejillas de la joven noble, pero lejos de mostrarse fría o esquiva, sonrió con una serenidad colmada y se acercó a nosotros con una copa de jugo fresco en la mano.",
        "—El mar tiene la virtud de calmar hasta los espíritus más inquietos —comentó Elysea con su voz melodiosa—. Y en este lugar, no hay máscaras que sostener ni batallas que librar.",
        "Contemplé a las tres mujeres extraordinarias que compartían mi vida. Entre ellas no había recelos, celos mezquinos ni rivalidades; solo existía una complicidad madura, un respeto infinito y la certeza inquebrantable de que nuestros corazones caminaban en la misma dirección.",
        "—Hoy va a ser un gran día —dije, rodeando con afecto los hombros de Kaoruko mientras le devolvía la sonrisa a Lortelle y a Elysea.",
        "* * *",
        "A media mañana, la arena de la playa se transformó en un improvisado coliseo deportivo.",
        "Caesar King, con un silbato metálico entre los labios y su actitud imponente de líder, trazó una cancha rectangular sobre la arena blanca con el talón de su sandalia, mientras Erza Scarlet ajustaba una red de pescador tensada entre dos postes de madera de balsa con la rigurosidad de un general preparando una trinchera.",
        "—¡Atención, holgazanes de la F-2! —bramó Caesar haciendo sonar el silbato con fuerza—. ¡Mucho descanso ablanda los músculos! ¡Vamos a jugar un torneo de voleibol con maná regulado! ¡El equipo perdedor lavará los platos del banquete de esta noche!",
        "El desafío encendió el orgullo competitivo de toda la clase.",
        "El primer enfrentamiento puso en cancha a dos titanes: por un lado, nuestro equipo formado por Ken, Burnice, Kaoruko y yo; por el otro, la escuadra compuesta por Lighter, Vlad, Caesar y Renji.",
        "El partido fue un espectáculo frenético de habilidad marcial, risas y polvo de arena.",
        "Ken servía el balón imbuyéndolo con una corriente de calor solar que silbaba en el aire; Caesar se lanzaba de plancha sobre la arena, bloqueando el remate con sus antebrazos como si interpusiera un muro blindado; Lighter colocaba la pelota con toques precisos y rápidos como jabs de boxeo; y Vlad saltaba con una elevación caballeresca impecable para rematar con fuerza milimétrica.",
        "En una de las jugadas decisivas, Kaoruko salvó una pelota imposible al borde de la red con un salto acrobático digno de su entrenamiento de esgrima, dejándome el balón suspendido en el aire a la altura perfecta.",
        "Aproveché la inercia, flexioné las piernas con el resorte elástico del Muay Thai y me elevé sobre la red. Concentré un leve pulso de maná de tierra en la palma de mi mano para darle un giro pesado al esférico y descargué un mate demoledor que levantó una nube de arena blanca al picar justo en la línea de fondo.",
        "—¡PUNTO! —gritó Erza con su brazo extendido como jueza implacable—. ¡Victoria para el equipo de Seier!",
        "Burnice saltó sobre la espalda de Ken celebrando entre risas y chispas de fuego, mientras Kaoruko me chocaba las dos manos con una alegría infantil contagiosa. En el otro lado de la red, Lighter se sacudía la arena del pecho quejándose con humor:",
        "—Ese salto tuyo no es humano, Seier. Deberían prohibir el Muay Thai en los deportes de playa.",
        "* * *",
        "Fuera de la cancha, la vida en la villa bullía con una calidez entrañable.",
        "Bajo una sombrilla amplia decorada con guirnaldas de conchas, **Natsuki** y **Lucy** compartían una mesa baja de madera.",
        "Lucy llevaba su sombrero de sol y unas gafas oscuras de carey, fingiendo observar el mar con total desinterés. Sin embargo, cada vez que Natsuki se acercaba a ofrecerle un vaso de té helado con fruta fresca, las mejillas de la pequeña artillera se encendían de rubor. Sus tres pequeños jabalíes mágicos correteaban alrededor de sus pies como perritos guardianes, mordisqueando alegremente las cáscaras de coco.",
        "—Ten, Lucy. Te preparé uno con extra de menta, como te gusta —dijo Natsuki con su sonrisa paciente.",
        "—M-más te vale que no esté demasiado dulce... —refunfuñó ella, pero tomó el vaso con ambas manos y dio un largo sorbo, dejando escapar un suspiro de satisfacción que no pudo ocultar—. Está... aceptable. Supongo que sirves para algo más que para recibir golpes en el calabozo.",
        "—Agradezco el cumplido, pequeña fiera —bromeó Natsuki.",
        "—¡Que no me llames así frente a todos! —espetó Lucy dándole un codazo juguetón en las costillas, pero sin poder evitar que una sonrisa tímida y radiante floreciera en sus labios, mientras disimuladamente rozaba su hombro contra el de él bajo la sombra.",
        "A pocos metros de allí, **Renji Abarai** intentaba impresionar a **Rukia Kuchiki** practicando equilibrio sobre una tabla de surf en la orilla. Una ola traicionera lo hizo caer de bruces sobre el agua poco profunda, provocando una risita cristalina de Rukia. La espadachina se acercó descalza al agua, extendiéndole una toalla blanca y sacudiéndole con ternura los granos de arena de su cabello rojo, haciendo que Renji se sonrojara hasta las orejas y rascara su nuca avergonzado.",
        "Cerca de las palmeras, **Akame** y **Pulchra** vigilaban una hilera de parrillas de carbón donde se asaban calamares gigantes, brochetas de pescado marinado y piñas caramelizadas. El aroma exquisito atrajo a **Yami Sukehiro**, quien finalmente se levantó de su hamaca, estiró los brazos perezosamente y tomó una brocheta directo de la parrilla.",
        "—Nada mal, chiquillas —murmuró Yami soltando una bocanada de humo—. Mientras haya comida de verdad y buena carne, estas vacaciones valen cada segundo.",
        "* * *",
        "Cuando el sol llegó a su cénit, toda la Clase F-2 y nuestros aliados nos reunimos alrededor de las largas mesas de madera bajo los porches de la villa para compartir el almuerzo.",
        "Entre brindis de sidra fría, risas estruendosas por las anécdotas de la batalla y platos que se vaciaban a una velocidad sobrehumana gracias al apetito voraz de Akame y Ken, me tomé un momento para contemplar a todos los presentes.",
        "Apenas unas semanas atrás, éramos considerados los marginados de la academia: los rechazados de la F-2 a los que nadie auguraba un futuro brillante. Hoy, tras sobrevivir a torneos despiadados, abismos olvidados y asedios de pesadilla, éramos una unidad invencible.",
        "Pero sobre todo, éramos una familia.",
        "Kaoruko me sirvió un plato de mariscos con una mirada dulce; a mi izquierda, Lortelle me rozó la mano por debajo de la mesa con complicidad ardiente; y al frente, Elysea alzó su copa hacia mí con una reverencia sutil cargada de afecto puro.",
        "El sol brillaba con una fuerza magnífica sobre las aguas de la Costa Esmeralda.",
        "Las sombras del pasado habían quedado atrás, y el verano nos regalaba el tesoro más preciado de todos: la paz ganada con sangre, compartida entre quienes jamás nos dejarían caer."
    ]
}

def compilar_capitulo_59():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 59              ")
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
    from agregar_capitulo_36 import CAPITULO_36
    from agregar_capitulo_37 import CAPITULO_37
    from agregar_capitulo_38 import CAPITULO_38
    from agregar_capitulo_39 import CAPITULO_39
    from agregar_capitulo_40 import CAPITULO_40
    from agregar_capitulo_41 import CAPITULO_41
    from agregar_capitulo_42 import CAPITULO_42
    from agregar_capitulo_43 import CAPITULO_43
    from agregar_capitulo_44 import CAPITULO_44
    from agregar_capitulo_45 import CAPITULO_45
    from agregar_capitulo_46 import CAPITULO_46
    from agregar_capitulo_47 import CAPITULO_47
    from agregar_capitulo_48 import CAPITULO_48
    from agregar_capitulo_49 import CAPITULO_49
    from agregar_capitulo_50 import CAPITULO_50
    from agregar_capitulo_51 import CAPITULO_51
    from agregar_capitulo_52 import CAPITULO_52
    from agregar_capitulo_53 import CAPITULO_53
    from agregar_capitulo_54 import CAPITULO_54
    from agregar_capitulo_55 import CAPITULO_55
    from agregar_capitulo_56 import CAPITULO_56
    from agregar_capitulo_57 import CAPITULO_57
    from agregar_capitulo_58 import CAPITULO_58
    
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
        CAPITULO_35, CAPITULO_36, CAPITULO_37, CAPITULO_38, CAPITULO_39,
        CAPITULO_40, CAPITULO_41, CAPITULO_42, CAPITULO_43, CAPITULO_44,
        CAPITULO_45, CAPITULO_46, CAPITULO_47, CAPITULO_48, CAPITULO_49,
        CAPITULO_50, CAPITULO_51, CAPITULO_52, CAPITULO_53, CAPITULO_54,
        CAPITULO_55, CAPITULO_56, CAPITULO_57, CAPITULO_58, CAPITULO_59
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 59
    txt_cap59 = os.path.join(capitulos_dir, "59_Capitulo_59.txt")
    with open(txt_cap59, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_59['header'].upper()}: {CAPITULO_59['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_59['paragraphs']))
    print(f"[OK] Texto plano Capítulo 59 guardado en:\n  {txt_cap59}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 59
    docx_cap59 = os.path.join(capitulos_dir, "59_Capitulo_59.docx")
    create_single_chapter_document(CAPITULO_59, docx_cap59, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 59 guardado en:\n  {docx_cap59}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 59)
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
    compilar_capitulo_59()
