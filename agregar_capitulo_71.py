# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"c:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_71 = {
    "header": "Capítulo 71",
    "title": "Raíces de Insurrección: Sombras de Corrupción y el Juramento de los Pueblos",
    "paragraphs": [
        "La niebla matinal cubría como un sudario gris los campos desolados entre Sauceda y Fresnedas, el segundo pueblo del feudo de Morvath.",
        "Avanzábamos en formación de vanguardia rápida. Kaelan se deslizaba entre los matorrales con sus pasos felinos inaudibles, mientras Akame se movía por las copas de los árboles como una sombra espectral. Detrás marchábamos Vlad, Julian, Borgar, Faelar, Vance y yo, acompañados por diez milicianos seleccionados por el Sargento Roderick.",
        "Apenas habíamos recorrido tres leguas cuando Kaelan emergió del follaje con la respiración contenida y los ojos felinos dilatados:",
        "—Carretas en el camino viejo del peaje —susurró el lince señalando hacia una curva rocosa—. Huelo grano recién molido, pero también sudor rancio, orina y miedo. Mucho miedo.",
        "—¿Mercenarios? —preguntó Vlad desenvainando un palmo de su hoja.",
        "—Seis hombres armados con ballestas y cotas de malla con el blasón de Morvath —confirmó Akame, descendiendo de una rama de pino sin quebrar una sola brizna—. Escoltando dos carromatos pesados y una jaula cubierta con lonas de brea.",
        "Julian apretó los puños hasta que sus nudillos se tornaron blancos. El blasón que aquellos hombres llevaban en el pecho era el de su propia sangre.",
        "—Emboscada rápida —ordené en un susurro—. Borgar y Vlad bloquean el camino por el frente. Akame inutiliza a los ballesteros. Vance, cuida la retaguardia. No dejen que ninguno escape hacia la ciudad.",
        "* * *",
        "El asalto duró menos de un minuto.",
        "Borgar saltó desde una pendiente rugiendo con la furia de una avalancha de piedra. Su martillo de guerra impactó el eje de la carreta delantera, partiéndolo en astillas y deteniendo el convoy en seco. Antes de que los escoltas pudieran encajar sus virotes, dos sombras letales barrieron el fango: la hoja carmesí de *Murasame* destelló dos veces degollando a los tiradores de flanco, mientras Vlad aplastaba al capitán con su escudo pesado contra la rueda rota.",
        "—¡Ríndanse o sus cabezas rodarán antes de que toquen el suelo! —bramó Vlad con una autoridad que paralizó a los supervivientes.",
        "Los cuatro mercenarios restantes arrojaron sus armas al lodo, temblando al ver las capas rojas y doradas de la Orden del León y el uniforme militar de la Academia Imperial.",
        "Me adelanté cortando de un tajo las amarras de la lona negra del carromato trasero.",
        "Lo que encontramos dentro nos heló la sangre.",
        "* * *",
        "Hacinadas en una jaula de barrotes de hierro oxidado, seis personas yacían atadas con cadenas de argollas reforzadas: dos mujeres jóvenes con los vestidos desgarrados, un anciano con moretones en el rostro y tres muchachos adolescentes que apenas alcanzaban los catorce años. Todos mostraban quemaduras superficiales en los tobillos: marcas en forma de ocho puntas, el estigma preliminar de Khorne.",
        "—Por todos los dioses... —susurró Vance von Holst, retrocediendo con el rostro desencajado y una mano temblorosa sobre su estoque—. Están traficando con su propia gente...",
        "Julian se abalanzó hacia el carromato, desenvainando su daga para cortar las sogas de los prisioneros:",
        "—¡Tranquilos! ¡Están a salvo, se los juro! —gritaba con la voz quebrada por el horror y la culpa—. ¡Soy Julian de Morvath! ¡Nadie volverá a ponerles una mano encima!",
        "Al escuchar el apellido, una de las mujeres soltó un alarido de terror y el anciano intentó interponerse con desesperación:",
        "—¡No nos lleven a las fundiciones! ¡Por piedad, joven amo, pagamos el diezmo! ¡Entregamos hasta la última fanega de cebada!",
        "—¡No voy a lastimarlos! —insistió Julian cayendo de rodillas en el barro frente a la jaula, arrojando su daga al suelo para demostrar que estaba desarmado—. Mírenme... por favor, mírenme. No soy como mi padre. No soy como Valerius. He venido a detener esta locura o a morir intentándolo.",
        "El anciano observó los ojos llorosos del joven noble, su respiración agitada y la presencia protectora de los cadetes imperiales. Lentamente, la histeria se transformó en un silencio atónito.",
        "* * *",
        "Mientras Faelar y los milicianos alimentaban a los cautivos con raciones secas y agua limpia, Seier y Vlad revisaron los cofres de hierro transportados en la primera carreta.",
        "No solo contenían sacos de trigo marcados con los sellos comunales de Fresnedas; también hallaron libros contables confidenciales firmados por el Conde Raymond y cartas con el sello personal de Valerius.",
        "Vlad leyó los pergaminos a la luz de una antorcha, y su expresión se tornó cada vez más sombría:",
        "—Es peor de lo que imaginamos —dijo Vlad pasándome las páginas—. Los silos de grano de Morvath no están vacíos por malas cosechas. Raymond y su primogénito provocaron una hambruna artificial confiscando el setenta por ciento de la producción agrícola bajo supuestos 'impuestos de guerra'.",
        "—¿Y a dónde iba ese trigo? —preguntó Akame con frialdad.",
        "—Lo contrabandeaban fuera del condado a través de los puertos fluviales para comprar armas negras y pagar a los mercenarios de Malrik —respondí con rabia tras examinar los recibos—. Y lo más atroz: a los campesinos que no podían pagar el tributo en moneda, les embargaban a sus hijos bajo el eufemismo de 'leva laboral obligatoria'. En realidad, los entregaban como ganado al culto de Khorne para alimentar el portal de las minas a cambio de lingotes de plata demoníaca.",
        "Borgar escupió sobre las cenizas del camino:",
        "—Malditas ratas traidoras. Gobernar exprimiendo a tu propia sangre y vender niños a los demonios... Ni los peores orcos de las cordilleras se atreven a tanta bajeza.",
        "* * *",
        "Decidimos marchar de inmediato hacia Fresnedas, escoltando a los rescatados y a los mercenarios capturados.",
        "Fresnedas era un pueblo sombrío, de techos de paja carcomidos por la humedad y calles desiertas. Cuando la comitiva entró en la plaza, los aldeanos se asomaron desde las ventanas trancadas armados con horquillas y guadañas, temiendo otra redada de los recaudadores.",
        "Pero al ver a sus propios familiares bajar de las carretas con vida, los cerrojos comenzaron a abrirse.",
        "Una multitud de casi doscientos aldeanos rodeó la plaza central entre llantos de reencuentro y miradas llenas de rabia contenida. Julian subió a los escalones del pozo comunal de piedra. El muchacho respiró hondo, sacó el sello de plata de la Casa Morvath y lo alzó frente a todos:",
        "—¡Pueblo de Fresnedas! —proclamó Julian con una voz que resonó firme sobre el viento frío—. Sé lo que han sufrido. Sé que durante tres años mi padre, el Conde Raymond, y mi hermano mayor Valerius han saqueado sus graneros, han dejado morir de hambre a sus ancianos y han entregado a sus hijos a una secta abominable en las minas.",
        "Un murmullo de ira sorda recorrió a la multitud. Varios hombres corpulentos aferraron sus herramientas de labranza con nudillos tensos.",
        "—¡No vengo a pedirles obediencia ciega! —continuó Julian, golpeándose el pecho con el puño—. Vengo a decirles que el blasón que llevo no representará más la tiranía ni la cobardía. He traído a la Sección Uno de la Academia Imperial. La orden de auxilio ya vuela hacia la Corona y cien caballeros reales vienen en camino. Pero hasta que lleguen, ¡no permitiré que saquen a una sola persona más de sus hogares!",
        "Julian desenvainó su espada corta y la clavó en la tierra entre los adoquines:",
        "—Si alguien aquí quiere pelear por su pan, por sus hijos y por su libertad, juro por mi honor y por los dioses que estaré en la primera fila de la trinchera. ¡Morvath no le pertenece a los demonios ni a los tiranos! ¡Morvath es de su gente!",
        "* * *",
        "Las palabras de Julian no eran la retórica vacía de un cortesano mimado; eran el grito desgarrado de un joven que elegía la justicia sobre su propio linaje podrido.",
        "Un herrero corpulento de barba gris, apodado Garret, dio un paso al frente arrojando su delantal de cuero:",
        "—El conde se llevó a mi sobrina el mes pasado —dijo Garret con voz ronca—. Nos dijeron que murió de peste en las barracas. Si hay una oportunidad de aplastar a esos bastardos y recuperar lo que queda de nuestro pueblo, cuente con mi mazo, joven señor Julian.",
        "—¡Y con mis hachas! —gritó un leñador tuerto.",
        "—¡Y con nuestros arcos de caza! —corearon tres cazadores jóvenes de las colinas.",
        "Uno a uno, más de sesenta hombres y mujeres dieron un paso al frente, alzando guadañas, hachas de tala, lanzas de caza y martillos. La chispa de la rebelión había prendido en el corazón del campesinado.",
        "* * *",
        "No perdimos un solo segundo.",
        "Bajo las instrucciones tácticas de Seier y Vlad, convertimos la plaza de Fresnedas en un cuartel de adiestramiento militar y logística de guerra.",
        "Borgar se adueñó de la fragua comunal con Garret: con fuelle y carbón al máximo, el enano comenzó a transformar azadones y rejas de arado en puntas de lanza reforzadas, filos de alabarda y planchas de acero para blindar los carros de transporte.",
        "—¡El hierro templado no discute con la carne! —reía Borgar entre martillazos ensordecedores—. ¡En cuatro horas les tendré cincuenta lanzas capaces de perforar cotas de malla mercenarias!",
        "Faelar y Akame reunieron a los cazadores y exploradores locales para enseñarles señales de silbato, emboscadas en maleza y vigilancia perimetral con arcos largos. Kaelan los instruyó en el rastreo de centinelas nocturnos.",
        "Mientras tanto, Vance von Holst instruía a los campesinos más robustos en la formación de falange y disciplina de escudos:",
        "—Mantengan la línea recta —ordenaba Vance con porte impecable pero sin altanería—. Cuando un mercenario a caballo cargue, no corran; claven el talón de la lanza en la tierra y usen el peso del bruto en su contra.",
        "Julian, vestido con una cota de cuero reforzada por Garret, entrenaba espada mano a mano junto a sus nuevos soldados, ganándose el respeto sincero de cada aldeano con sudor y golpes limpios.",
        "* * *",
        "Al caer la tarde, la rebelión popular de Morvath ya no era una multitud desordenada; era la recién nacida **Guardia Popular de Julian**, integrada por casi ochenta combatientes decididos a recuperar su hogar.",
        "Además, los carros interceptados nos proveyeron de seis sacos de harina que fueron distribuidos de inmediato entre las familias hambrientas, devolviendo la esperanza a Fresnedas.",
        "En ese momento, un hombre sucio de tizne y con las ropas desgarradas corrió hacia la fragua escoltado por Kaelan:",
        "—¡Señor Julian! ¡Caballeros! —jadeó el fugitivo cayendo de rodillas—. Soy Tarek, minero del sector sur... Logré escapar por los antiguos canales de drenaje de las fundiciones.",
        "Seier se arrodilló junto al minero:",
        "—Toma agua, Tarek. Háblanos de las minas.",
        "El hombre desplegó sobre una caja de municiones un trozo de tela manchada de carbón donde había dibujado los túneles subterráneos:",
        "—Los cultistas han traído altares de bronce negro a la galería principal. Valerius estuvo allí esta mañana furioso porque Malrik no reportó... Está reuniendo a todos los mercenarios en el castillo de la colina y ha ordenado cerrar los portones de la ciudad amurallada al anochecer.",
        "Tarek señaló un conducto que cruzaba por debajo de los cimientos de la muralla de Morvath:",
        "—Este canal de desagüe desemboca directamente detrás del cuartel de la guardia baja. Si entramos por aquí al caer la noche, podemos abrir las puertas principales desde adentro sin que los centinelas de las almenas nos detecten.",
        "Vlad y Seier cruzaron miradas cargadas de resolución táctica.",
        "—Es la oportunidad que necesitábamos para limpiar Morvath desde adentro hacia afuera —sentenció Seier, poniendo su mano sobre el hombro de Julian—. Esta noche quebramos el cerco del conde."
    ]
}

def compilar_capitulo_71():
    print("[*] Iniciando compilacion del Capitulo 71...")
    
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
    from agregar_capitulo_59 import CAPITULO_59
    from agregar_capitulo_60 import CAPITULO_60
    from agregar_capitulo_61 import CAPITULO_61
    from agregar_capitulo_62 import CAPITULO_62
    from agregar_capitulo_63 import CAPITULO_63
    from agregar_capitulo_64 import CAPITULO_64
    from agregar_capitulo_65 import CAPITULO_65
    from agregar_capitulo_66 import CAPITULO_66
    from agregar_capitulo_67 import CAPITULO_67
    from agregar_capitulo_68 import CAPITULO_68
    from agregar_capitulo_69 import CAPITULO_69
    from agregar_capitulo_70 import CAPITULO_70
    from agregar_capitulo_71 import CAPITULO_71
    
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
        CAPITULO_55, CAPITULO_56, CAPITULO_57, CAPITULO_58, CAPITULO_59,
        CAPITULO_60, CAPITULO_61, CAPITULO_62, CAPITULO_63, CAPITULO_64,
        CAPITULO_65, CAPITULO_66, CAPITULO_67, CAPITULO_68, CAPITULO_69,
        CAPITULO_70, CAPITULO_71
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 71
    txt_cap71 = os.path.join(capitulos_dir, "71_Capitulo_71.txt")
    with open(txt_cap71, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_71['header'].upper()}: {CAPITULO_71['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_71['paragraphs']))
    print(f"[OK] Texto plano Capítulo 71 guardado en:\n  {txt_cap71}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 71
    docx_cap71 = os.path.join(capitulos_dir, "71_Capitulo_71.docx")
    create_single_chapter_document(CAPITULO_71, docx_cap71, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 71 guardado en:\n  {docx_cap71}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 71)
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
    compilar_capitulo_71()
