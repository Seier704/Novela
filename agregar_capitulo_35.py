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

CAPITULO_35 = {
    "header": "Capítulo 35",
    "title": "El Despertar de la Penumbra",
    "paragraphs": [
        "Mientras el Escuadrón 7 descendía por la escalinata ciclópea hacia el Segundo Estrato, las entrañas de las Cavernas del Miasma Carmesí vibraban como un hormiguero en pie de guerra.",
        "Los diferentes sectores subterráneos se habían convertido en un laberinto de choques violentos, donde los escuadrones mixtos de dieciséis años luchaban por abrirse paso entre enjambres de bestias abisales. Pero en aquellos túneles oscuros, los cadetes de la Clase F-2 estaban dejando una marca imborrable.",
        "* * *",
        "En el Sector Tres, una galería gigantesca repleta de estalactitas de sulfuro parecía haberse transformado en un volcán en erupción.",
        "—¡FWOOOOOSH!",
        "Una llamarada colosal de trescientos grados barrió el pasillo de piedra, incinerando a una docena de ciempiés blindados que se retorcían en cenizas. Burnice White, con sus gafas rojas puestas y dos propulsores de ignición humeando en sus antebrazos, reía desatada dando saltitos entre las rocas calcinadas.",
        "—¡Abran paso, bichos feos! ¡El nitro está fresco y tenemos un calabozo entero por limpiar!",
        "Desde el techo de la caverna, Ken cayó en picado como un bólido solar, quebrando el caparazón de un monstruo acorazado con una patada de hacha que hizo retumbar la cueva entera: *¡BOOM!*",
        "El caballero rúnico de la Clase B-1 que los acompañaba contemplaba la masacre con la espada a medio desenvainar y la mandíbula desencajada, completamente incapaz de procesar cómo dos jóvenes de su misma edad podían desatar semejante nivel de demolición pirotécnica sin pestañear.",
        "A kilómetros de distancia, en el Sector Cinco, la batalla se libraba con una elegancia gélida y silenciosa.",
        "Subaru Hoshina se movía por un pasillo de cristales de cuarzo como un fantasma de plata. Aplicando el juego de pies que Lighter le había enseñado, pivotaba sobre sus talones con una soltura impecable, deslizando sus dos dagas por los tendones de los acechadores que intentaban saltar sobre ellos.",
        "Cuando un cuadrúpedo de escamas negras intentó abalanzarse sobre su espalda, un puño vendado como una maza de hierro entró en su campo de visión:",
        "¡PAAAAF!",
        "Lighter conectó un gancho de derecha demoledor a la mandíbula de la criatura, quebrándole el cuello en el aire y haciéndola estrellarse inerte contra la pared de piedra. El boxeador se acomodó las gafas de sol sobre la coronilla, le guiñó un ojo a Hoshina y ella le devolvió una sonrisa dulce y segura que dejó mudos a los dos nobles de la A-1 que intentaban seguirles el ritmo desde atrás.",
        "En otros sectores, Erza y Akame desmembraban manadas enteras con una danza coordinada de hojas relampagueantes en el Escuadrón 2; mientras que en el Escuadrón 4, el látigo dentado de Renji y las ráfagas de hielo puro de Rukia despejaban cavernas enteras con una sintonía forjada desde la infancia.",
        "Sin embargo, la prueba más dura y transformadora de aquella jornada se estaba librando en el Sector Seis.",
        "* * *",
        "El Sector Seis era una garganta estrecha y asfixiante, donde el miasma carmesí se mezclaba con emanaciones de maná de oscuridad pura.",
        "El Escuadrón 6 —compuesto por Yami Sukehiro, Subaru Natsuki, un joven clérigo de la Clase C y dos lanceros de la Clase B— había sido emboscado por una manada de **Gárgolas Umbrías** lideradas por un colosal **Gólem de Ceniza y Miasma** de cuatro metros de altura.",
        "Las criaturas aladas se movían entre las sombras del techo con una velocidad aterradora, atacando los flancos de la formación y desatando chillidos que aturdían la mente.",
        "—¡No puedo mantener la barrera sagrada! ¡El miasma está devorando mis oraciones! —gritaba el clérigo de dieciséis años, cayendo de rodillas con las manos temblando de pánico.",
        "Uno de los lanceros fue derribado por un coletazo del gólem y el segundo retrocedió arrastrando a su compañero herido, dejando la retaguardia completamente desprotegida.",
        "Al frente, Yami Sukehiro mantenía la línea con su postura pesada habitual, irradiando una presencia dominante. Con un cigarrillo apagado entre los labios y el torso cubierto únicamente por su camiseta blanca, desenvainó su katana negra con un silbido amenazante:",
        "—Vaya panda de escandalosos —gruñó Yami con su voz áspera—. Cierren la boca y no estorben.",
        "Una marea de maná de oscuridad densa envolvió la hoja de su espada. Yami saltó de frente hacia el gólem, bloqueando el puño de piedra de la bestia con un crujido sísmico y comenzando un intercambio brutal de tajos y estocadas.",
        "Pero la manada de gárgolas no atacó a Yami.",
        "Tres de las bestias aladas aprovecharon que el espadachín estaba ocupado con el coloso, plegaron las alas y se arrojaron en picado hacia el clérigo indefenso... y hacia la espalda ciega de Yami.",
        "* * *",
        "Subaru Natsuki estaba a cinco metros de distancia.",
        "Tenía las piernas temblando por el pánico instintivo. Sus manos sudorosas apretaban la empuñadura de una espada corta de infantería que pesaba como plomo. Durante semanas, se había sentido como el eslabón más débil de la Clase F: un chico común sin la fuerza brutal de Silva, sin la velocidad de Ken ni el linaje noble de los demás. Un bufón que solo servía para recibir palizas en los entrenamientos y cargar los petates de Yami.",
        "«¡Muévete, idiota! ¡Muévete!», se gritó a sí mismo por dentro.",
        "Vio las garras ensangrentadas de la primera gárgola descender hacia el cuello del clérigo.",
        "Y en ese segundo de desesperación absoluta, algo se quebró dentro del pecho de Natsuki.",
        "No fue rabia ciega. Fue la negativa absoluta a volver a ser un espectador inútil que solo sabe mirar cómo los demás sangran por él.",
        "—¡¡¡DEJEN EN PAZ A MI COMPAÑERO, MALDITAS BESTIAS!!! —bramó Natsuki con toda la fuerza de su garganta.",
        "Cerró los ojos y conectó con la chispa fría y profunda que Yami le había obligado a despertar a golpes durante los entrenamientos nocturnos.",
        "En sus venas, su maná de **Oscuridad** no fluyó como una cuchilla afilada ni como una llamarada destructiva.",
        "Fluyó como un manto de noche absoluta.",
        "Natsuki extendió la mano izquierda hacia el frente con los dedos abiertos:",
        "—¡¡¡MANÁ DE OSCURIDAD: VELO DE PENUMBRA... SHAMAK!!!",
        "* * *",
        "¡FWOOOSH!",
        "Una niebla negra como el alquitrán, densa y voraz, brotó de la palma de Natsuki y se expandió en una esfera de cinco metros que tragó por completo a las tres gárgolas en mitad de su vuelo.",
        "El *Shamak* no era un hechizo de daño físico.",
        "Era una maldición de privación sensorial absoluta.",
        "Dentro de la esfera de penumbra, las tres criaturas perdieron la vista, el oído y la noción del equilibrio en una milésima de segundo. Comenzaron a chillar enloquecidas, chocando entre sí y aleteando a ciegas contra las paredes de roca sin saber hacia dónde atacar.",
        "Natsuki no dudó.",
        "Recordó cada corrección de postura que Yami le había grabado en los músculos. Adelantó el pie izquierdo, flexionó las rodillas para bajar su centro de gravedad y se deslizó por debajo del ala desorientada de la bestia más cercana.",
        "Entró en su punto ciego, levantó su espada corta con ambas manos y clavó la hoja con todo el peso de su cuerpo en la juntura blanda debajo de su garganta:",
        "¡¡¡KRAAAK!!!",
        "La gárgola soltó un gorgoteo ahogado y se desplomó sobre la grava con el cuello quebrado.",
        "La segunda bestia intentó morder el aire a ciegas; Natsuki rodó por el suelo con agilidad felina, desenfundó su daga corta de cinto con la mano izquierda y le rebanó los tendones de la pata trasera, dejándola indefensa para que los dos lanceros de la Clase B la ensartaran con sus armas desde atrás.",
        "La tercera criatura intentó huir hacia el techo, pero Natsuki concentró los últimos restos de su maná oscuro en la punta de su bota y le lanzó una piedra imbuida en sombra que golpeó su ala, haciéndola estrellarse contra el suelo antes de rematarla.",
        "La retaguardia estaba completamente a salvo.",
        "* * *",
        "Al mismo tiempo, al frente del túnel, la batalla de Yami llegaba a su clímax.",
        "El espadachín de dieciséis años vio por el rabillo del ojo la nube de oscuridad de Natsuki y una chispa de aprobación brilló en sus ojos.",
        "—Vaya... parece que el mocoso por fin aprendió a morder —murmuró Yami.",
        "Con un movimiento de muñeca fluido, Yami concentró una cantidad monstruosa de maná negro en el filo de su katana, haciendo que la hoja pareciera doblar el espacio a su alrededor:",
        "—¡ARTE DE LA OSCURIDAD: CORTE DEL ABISMO!",
        "¡¡¡ZAAAAASH!!!",
        "Una medialuna de oscuridad pura de seis metros cortó el aire a la velocidad del sonido. El colosal Gólem de Ceniza fue partido limpiamente por la mitad a la altura del torso, colapsando en una avalancha de rocas y polvo negro que se desvaneció en el aire.",
        "El silencio regresó a la caverna.",
        "Yami sacudió la katana con un chasquido impecable y la enfundó en su vaina de madera.",
        "Caminó despacio hacia donde estaba Natsuki.",
        "El chico de pelo negro estaba sentado sobre una roca, empapado en sudor, con la ropa desgarrada y respirando con agitación, sosteniendo su espada corta manchada de sangre negra. A su lado, el clérigo le agradecía entre lágrimas por haberle salvado la vida.",
        "Yami se detuvo frente a Natsuki. Sacó un cigarrillo de su bolsillo, lo encendió con la chispa de su pulgar y soltó una bocanada de humo hacia el techo de la cueva.",
        "Natsuki levantó la mirada con timidez, esperando el habitual regaño por haberse quedado sin aire.",
        "Pero Yami no lo regañó.",
        "El espadachín bajó la mano pesada sobre la cabeza de Natsuki y le revolvió el cabello con una fuerza tosca pero indiscutiblemente afectuosa.",
        "—Nada mal para un novato llorón —dijo Yami con una sonrisa serena y cómplice—. Usaste el cerebro para taparles los ojos y clavaste el filo donde correspondía. Esa oscuridad tuya... empieza a parecerse a un arma de verdad.",
        "Yami sacó una cantimplora de agua fresca de su cinturón y se la arrojó al regazo.",
        "—Bébete eso y descansa cinco minutos. Aún nos quedan muchos pisos por bajar, mocoso... y no pienso dejarte atrás.",
        "Natsuki atrapó la cantimplora con dedos temblorosos. Una sonrisa amplia, radiante y cargada de un orgullo infinito se dibujó en su rostro mientras bebía el agua fresca.",
        "Por primera vez desde que pisó la academia de Solaria, Subaru Natsuki sintió que ya no era una carga para nadie.",
        "El fuego de la Clase F-2 ardía en cada uno de sus rincones.",
        "Y en las profundidades de aquel calabozo, cada uno de sus guerreros de dieciséis años estaba forjando su propio camino hacia la grandeza."
    ]
}

def compilar_capitulo_35():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 35              ")
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
        CAPITULO_35
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 35
    txt_cap35 = os.path.join(capitulos_dir, "35_Capitulo_35.txt")
    with open(txt_cap35, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_35['header'].upper()}: {CAPITULO_35['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_35['paragraphs']))
    print(f"[OK] Texto plano Capítulo 35 guardado en:\n  {txt_cap35}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 35
    docx_cap35 = os.path.join(capitulos_dir, "35_Capitulo_35.docx")
    create_single_chapter_document(CAPITULO_35, docx_cap35, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 35 guardado en:\n  {docx_cap35}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 35)
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
    compilar_capitulo_35()
