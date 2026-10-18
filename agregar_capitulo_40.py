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

CAPITULO_40 = {
    "header": "Capítulo 40",
    "title": "El Retorno del Abismo",
    "paragraphs": [
        "La penumbra subterránea comenzó a disiparse tenuemente cuando las venas de cuarzo luminoso incrustadas en el techo del cañón cambiaron a un tono azul plateado, marcando la llegada del amanecer en la superficie.",
        "Desperté antes de que el primer destello de claridad tocara la gruta. En realidad, jamás había llegado a dormir por completo; mi cuerpo se mantuvo en una vigilia disciplinada, escuchando cada gota de agua y cada eco en el viento.",
        "Lortelle seguía profundamente dormida sobre mi pecho, con sus brazos delgados rodeando mi cintura y su rostro tibio enterrado contra mi piel. La pesadilla de la medianoche había quedado atrás; ahora su respiración era rítmica, serena y relajada.",
        "Con sumo cuidado para no sobresaltarla, deslicé una mano por su hombro.",
        "—Lortelle... —susurré con suavidad cerca de su oído—. Es hora de despertar. Nuestras ropas están secas y debemos ponernos en marcha.",
        "La joven pelirroja parpadeó con pereza. Sus párpados se abrieron despacio, revelando esos ojos verde esmeralda todavía empañados por el sopor. Tardó un par de segundos en procesar dónde estaba: acurrucada casi desnuda en el regazo de un chico, sostenida con firmeza protectora. Un suave rubor carmesí tiñó de inmediato sus mejillas, pero lejos de apartarse con timidez o fingir indignación, soltó un bostezo adorable y apoyó la barbilla en mi esternón, mirándome con una sonrisa cargada de complicidad.",
        "—Cinco minutos más, Seier... —murmuró con una voz ronca y seductora por el sueño—. Hacía años que mi cama de plumas en la capital no se sentía ni la décima parte de cómoda que este pecho tuyo.",
        "—Si nos quedamos cinco minutos más, las patrullas de rescate de la academia darán por sentado que fuimos devorados por los monstruos —repliqué con una sonrisa de medio lado, dándole una palmada suave en la espalda para animarla.",
        "Lortelle soltó una pequeña risita, estiró sus extremidades con la gracia de un felino y finalmente se incorporó.",
        "* * *",
        "Nos vestimos con rapidez. El uniforme y las capas, tras pasar la noche junto a la llama controlada de Lortelle sobre las piedras calientes, estaban completamente secos y conservaban un calor reconfortante que nos protegió de la humedad cortante del exterior.",
        "Aseguré mi daga al cinturón, ajusté las correas de combate en mis antebrazos y revisé el estado de la herida en el tobillo de Lortelle. Gracias a las compresas frías y al descanso, la hinchazón había disminuido notablemente, aunque todavía renqueaba al dar pasos en falso.",
        "—El ascenso no va a ser un paseo de compras en el distrito comercial, Lortelle —le advertí mientras examinaba la imponente pared de roca vertical que se alzaba hacia la grieta original—.",
        "—Tengo dieciséis años, Seier, pero he sobrevivido a tres atentados de gremios rivales y a una emboscada en las montañas del norte —respondió ella, ajustándose la chaqueta dorada de la Clase C con orgullo renovado—. Mientras tú pongas los músculos y los puños, yo me encargaré de despejar el aire con mis llamas. Ninguna grieta me va a impedir volver para cobrar mis inversiones.",
        "La determinación en su mirada me sacó una sonrisa genuina. Me agaché frente a ella, ofreciéndole mi espalda.",
        "—Sube. En las zonas empinadas te llevaré a cuestas. Necesitamos avanzar rápido.",
        "Lortelle no dudó. Se colgó de mis hombros con naturalidad, enlazando sus brazos alrededor de mi cuello y apoyando su barbilla junto a mi oreja.",
        "—Eres un excelente medio de transporte, Seier Silva. Muy resistente y con calefacción incluida.",
        "—Guárdate el sarcasmo para cuando salgamos, pelirroja —contesté, antes de concentrar mi maná de elemento tierra en las plantas de mis botas.",
        "* * *",
        "El trayecto de regreso fue un calvario de roca, sudor y acero.",
        "Las paredes del abismo eran húmedas y resbaladizas, cubiertas de líquenes negros y salientes traicioneras. Para sortear los desniveles verticales de más de treinta metros, canalicé mi maná telúrico, moldeando pequeñas plataformas y asideros de piedra en la pared con cada pisotón y cada agarre, escalando con la agilidad explosiva que me brindaba mi entrenamiento físico.",
        "Sin embargo, el abismo no nos dejaría marchar sin resistencia.",
        "A mitad del ascenso, en un saliente estrecho donde el viento silbaba con fuerza, un chillido gutural resonó entre las grietas. Cinco criaturas emergieron de las hendiduras laterales: eran acechadores Skavens, más ágiles y armados con dagas envenenadas de piedra bruja verde.",
        "—¡A la izquierda, Seier! —gritó Lortelle aferrada a mi hombro.",
        "Sin soltarla por completo, giré sobre mi eje con la guardia compacta del Jeet Kune Do. Un Skaven se lanzó en picada con las garras extendidas; descargué una patada lateral frontal cargada de fuerza bruta que impactó directo en su pecho, pulverizándole la caja torácica y arrojándolo al vacío.",
        "Los otros cuatro intentaron rodearnos en la cornisa estrecha.",
        "—¡Agáchate! —ordenó Lortelle.",
        "Incliné mi torso al instante. Desde mi hombro, la joven maga extendió su mano derecha. Esta vez no fue una chispa diminuta; con una concentración implacable, conjuró tres lanzas de fuego carmesí perfectamente comprimidas que salieron disparadas como saetas incandescentes. Los proyectiles atravesaron los cráneos de tres alimañas al unísono, incinerándolos antes de que tocaran el suelo.",
        "El último acechador intentó saltar hacia el rostro de Lortelle. Di un paso adelante, lo atrapé del cuello en pleno aire con mi mano izquierda y, con un giro violento de cadera, lo azoté contra la pared de basalto, rompiéndole la columna de un solo impacto seco.",
        "El silencio regresó al desfiladero.",
        "Lortelle respiraba con fuerza contra mi cuello, con los ojos brillantes por la adrenalina. Me miró de reojo, pasando suavemente una mano por mi cabello desordenado.",
        "—Buena coordinación, socio. Si alguna vez te cansas de pelear en la academia, te contrato como mi guardaespaldas personal.",
        "—Primero salgamos de aquí antes de hablar de contratos —respondí, saltando hacia la siguiente cornisa ascendente.",
        "* * *",
        "Dos horas de agotador ascenso después, divisamos la luz de bengalas mágicas filtrándose a través de la hendidura superior del Segundo Estrato.",
        "Voces conocidas resonaban entre los ecos de las cavernas superiores.",
        "—¡Seier! ¡Lortelle! ¿¡Hay alguien ahí abajo!? —la voz angustiada y potente de Kaoruko atravesaba la penumbra, acompañada por las órdenes firmes de Dante Vance y las invocaciones de rastreo de Freya Skuld.",
        "Aproveché una saliente de roca, concentré mi maná de tierra para crear un pilar que nos impulsara y salté hacia el borde de la cornisa principal.",
        "—¡Por aquí! —grité con voz clara y firme.",
        "Los miembros del Escuadrón 7 se giraron de golpe. Al verme salir de la grieta con Lortelle sobre mi espalda, la expresión de desesperación en el rostro de Kaoruko se transformó en un estallido de alivio inmenso.",
        "—¡SEIER! —gritó ella, soltando su espada al suelo sin importarle el protocolo ni las miradas ajenas.",
        "Kaoruko corrió hacia mí a toda prisa, con lágrimas en los ojos, arrojándose a mis brazos con tanta fuerza que casi nos hace caer hacia atrás. La rodeé con fuerza, sintiendo sus temblores y el latido desesperado de su corazón.",
        "—¡Idiota, estúpido, imprudente...! —sollozaba ella contra mi cuello, golpeándome el pecho con pequeños puños impotentes antes de besarme apasionadamente en los labios frente a todos—. Creí que te había perdido... vi cómo caías a ese maldito pozo...",
        "—Te dije que volvería, Kaoruko. No iba a dejarte sola en este calabozo —le susurré al oído, acariciándole la espalda hasta que sus sollozos se calmaron.",
        "Elysea de Vane se acercó suspirando con alivio, mientras Dante envainaba su espada corta y Freya cancelaba los sellos de búsqueda. Varios profesores y miembros de la guardia de la academia llegaron poco después, acordonando la zona tras comprobar que las escaramuzas de los Skavens habían sido finalmente repelidas por los escuadrones combinados.",
        "* * *",
        "Lortelle descendió de mi espalda con la ayuda de Elysea, recomponiendo su postura aristocrática con una rapidez asombrosa. Los representantes de la Clase C y varios mercaderes de la Compañía Elte asignados a la retaguardia corrieron hacia ella con camillas y pociones curativas.",
        "Sin embargo, antes de permitir que su comitiva la rodeara, Lortelle se detuvo. Caminó despacio hacia mí, ignorando a los guardias y a los demás estudiantes.",
        "Se detuvo a escasos centímetros de mi pecho. Sus ojos verdes centellearon con una mezcla exquisita de picardía, afecto profundo y soberbia.",
        "—Me salvaste la vida, Seier Silva —dijo en voz alta, asegurándose de que todos en el perímetro escucharan la seriedad de sus palabras—. Y la Compañía Elte jamás deja una deuda sin saldar.",
        "Antes de que pudiera responderle, Lortelle dio un paso más, tomó suavemente mi solapa con su mano derecha, se alzó en puntas de pie y plantó un beso cálido, tierno y deliberadamente prolongado sobre mi mejilla.",
        "Un murmullo de sorpresa recorrió a los estudiantes de la Clase C y a los presentes.",
        "—No creas que te librarás de mí tan fácil, salvador mío —me susurró al oído con una sonrisa cargada de promesas íntimas—. Vendré a buscarte a los dormitorios de la Clase F. Tenemos mucho de qué hablar... y muchas noches que recuperar.",
        "Me guiñó un ojo con coquetería descarada, dio media vuelta con el porte elegante de una reina mercante y se dejó escoltar por sus sanadores hacia la salida del calabozo.",
        "* * *",
        "El silencio que quedó tras su partida fue denso.",
        "Me giré lentamente hacia Kaoruko. La hermosa espadachina de la Clase F-2 estaba de pie con los brazos cruzados bajo el pecho, inflando las mejillas con un puchero inocultable y una ceja alzada que destilaba un fuego peligroso.",
        "—Vaya, vaya... —murmuró Kaoruko, caminando alrededor de mí como una pantera inspeccionando a su presa—. Caer por un abismo de cien metros con una de las chicas más codiciadas de la Clase C, pasar la noche entera a solas con ella en el fondo de una cueva... y regresar recibiendo besos y promesas de visitas nocturnas. Sí que eres todo un caballero andante, Seier.",
        "Elysea soltó una risita nerviosa y Dante fingió una repentina fascinación por el filo de su espada, dándonos espacio discretamente.",
        "—Kaoruko, la situación fue de vida o muerte —le expliqué con calma y sinceridad, mirándola directamente a los ojos—. Estaba herida y las cavernas estaban heladas. La protegí como habría protegido a cualquiera de mis compañeros.",
        "Kaoruko me miró fijamente durante unos largos segundos. El puchero en sus labios carnosos comenzó a suavizarse, dando paso a una mirada reflexiva y comprensiva.",
        "Dio un paso adelante y apoyó su mano en mi mejilla, justo sobre el lugar donde los labios de Lortelle habían dejado su marca tibia.",
        "—Sé quién eres, Seier. Y sé perfectamente que si ella está viva, es porque pusiste tu propio cuerpo como escudo —suspiró Kaoruko, con un tono mucho más maduro y sereno—. No puedo evitar sentir celos... verte tan cercano a otra mujer me quema un poco el pecho, no te lo voy a negar.",
        "Hizo una pausa, sonriendo con una ternura infinita mientras me tomaba de las manos.",
        "—Pero crecí en una familia de guerreros y nobles de la frontera, Seier. Sé muy bien que los grandes hombres, los que poseen un poder descomunal y un corazón capaz de proteger a los suyos, rara vez caminan solos. En este mundo, que un guerrero comparta su camino con más de una compañera no es una deshonra ni un problema... siempre y cuando haya honestidad absoluta, respeto mutuo y todas estemos de acuerdo.",
        "Kaoruko se acercó a mi rostro, rozando sus labios con los míos con una posesividad dulce y apasionada.",
        "—Solo ten en claro una cosa, mi amado idiota: yo fui la primera en tu corazón en esta academia, y no pienso ceder mi lugar ante ninguna mercader astuta.",
        "—Nadie va a quitarte tu lugar jamás, Kaoruko —le respondí, besándola con la certeza y la devoción que merecía.",
        "Con el calabozo a nuestras espaldas y el sol del mediodía esperando en la salida, el Escuadrón 7 emprendió la marcha de retorno, sabiendo que el abismo no solo no nos había destruido, sino que había fortalecido nuestros lazos para las batallas que aún estaban por venir."
    ]
}

def compilar_capitulo_40():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 40              ")
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
        CAPITULO_40
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 40
    txt_cap40 = os.path.join(capitulos_dir, "40_Capitulo_40.txt")
    with open(txt_cap40, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_40['header'].upper()}: {CAPITULO_40['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_40['paragraphs']))
    print(f"[OK] Texto plano Capítulo 40 guardado en:\n  {txt_cap40}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 40
    docx_cap40 = os.path.join(capitulos_dir, "40_Capitulo_40.docx")
    create_single_chapter_document(CAPITULO_40, docx_cap40, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 40 guardado en:\n  {docx_cap40}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 40)
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
    compilar_capitulo_40()
