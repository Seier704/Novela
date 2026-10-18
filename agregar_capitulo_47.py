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

CAPITULO_47 = {
    "header": "Capítulo 47",
    "title": "El Rugido del León y el Fin del Abismo",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "erza_scarlet.png"),
            "caption": "Ilustración: Erza Scarlet (Defensa Implacable del Flanco)"
        },
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "akame.png"),
            "caption": "Ilustración: Akame (El Filo Silencioso de la Victoria)"
        }
    ],
    "paragraphs": [
        "El rugido de la batalla en el cuello de botella de la Roca Carmesí alcanzó un punto de ebullición absoluta.",
        "Con Lortelle y los demás prisioneros a salvo tras los muros gracias al audaz asalto de los combatientes de Calydon de la Clase F-2, el peso de la contención se transformó en una embestida furiosa. Ya no había rehenes que proteger ni dudas que refrenaran nuestros golpes: era una guerra de exterminio entre los defensores de la superficie y las aberraciones nacidas en las entrañas de la tierra.",
        "En el centro del frente, la marea comenzaba a cambiar de bando.",
        "Los cadetes de la Academia, inspirados por la presencia inquebrantable de Vlad y la ferocidad de la Clase F-2, formaron una muralla de picas y escudos que resistía las embestidas con una firmeza marmórea. Sin embargo, en el ojo del huracán, una mole continuaba siendo prácticamente indestructible: el **Señor de la Guerra Skaven**.",
        "Montado sobre su monstruosidad de púas y piel correosa, el tirano del abismo no retrocedía ni un palmo. Su alabarda rúnica de tres metros silbaba con cortes mortales, despidiendo oleadas de vapor verde que carcomían la roca y el acero. A su alrededor, las armaduras de bronce brujo forjadas con piedra de disformidad repelían nuestras estocadas con un fulgor enfermizo.",
        "* * *",
        "—¡A la carga, muchachos! ¡No dejen que respire! —bramó Caesar King, avanzando con su escudo de torre en alto.",
        "Caesar se estrelló contra el flanco derecho de la bestia acorazada, absorbiendo con el titanio de su defensa una coz devastadora que habría quebrado las costillas de diez hombres ordinarios. La Reina de Calydon gruñó con los dientes apretados, clavando los talones en el barro y devolviendo un tajo salvaje con su gladius contra los tendones de la pata delantera del engendro.",
        "La criatura chilló de dolor, pero la coraza de hierro remachada en su carne apenas se abolló.",
        "Al mismo tiempo, Ken se lanzó por los aires envuelto en un torbellino de llamas doradas. Con el grito de un halcón de fuego, descargó una patada hacha sobre el hombro del Señor de la Guerra, provocando una explosión de calor que calcinó parte de los estandartes podridos que colgaban de su montura. El caudillo Skaven rugió, girando su alabarda con una velocidad aterradora para rebanar a Ken en el aire; Yami interceptó el filo a escasos centímetros con su katana envuelta en miasma negro, desviando la ponzoña con una onda gravitatoria que hizo temblar el suelo.",
        "—¡Maldita cucaracha blindada! —escupió Lighter, conectando una ráfaga fulminante de golpes de combustión directa sobre el pecho de la montura mientras Natsuki, con su aura plateada al límite, desestabilizaba a los guardias que intentaban cerrar filas alrededor de su amo.",
        "Corrí hacia el combate cerrado, canalizando todo el maná telúrico en mis piernas. Utilicé el juego de piernas del Jeet Kune Do para evadir un barrido letal de la alabarda, me deslicé bajo el vientre de la bestia y solté un rodillazo devastador de Muay Thai directo a su quijada acorazada.",
        "*¡¡CRACK!!*",
        "El impacto de masa compacta quebró una de las púas de hierro y levantó la cabeza del monstruo, haciéndolo retroceder tres pasos. Pero el Señor de la Guerra apenas pestañeó: con una vitalidad aberrante potenciada por el veneno en sus venas, el comandante rata aguantó el embate combinado sin caer. Giró la empuñadura de su arma como un molino y nos obligó a retroceder a todos, manteniendo a raya a seis combatientes de élite al mismo tiempo.",
        "Era un muro de desgaste puro. Cada herida que le infligíamos se cerraba con una pasta de pus y humo verde, y nuestros músculos comenzaban a resentir el veneno suspendido en el aire.",
        "Si no lográbamos quebrar su defensa de un solo golpe fulminante, la resistencia del Warlord terminaría agotándonos.",
        "* * *",
        "Y fue en ese instante de máxima tensión cuando el destino dio un vuelco ensordecedor.",
        "Desde las gargantas superiores del cañón del este, un sonido claro, vibrante y majestuoso cortó el estrépito de la carnicería.",
        "*¡¡¡TOOOORRRRRRR-TO-TOOOOOOOOO!!!*",
        "Eran trompetas de guerra de bronce pulido.",
        "El clamor no pertenecía a los cuernos huecos y cavernosos de los Skavens, ni al timbre improvisado de la guarnición estudiantil. Era el toque marcial de las fuerzas regulares de la corona.",
        "—¡Miren hacia la cresta del paso! —gritó un centinela desde lo alto del rastrillo.",
        "Sobre la colina este, recortados contra el cielo teñido de ceniza y humo, apareció una formación imponente que hizo retumbar los cimientos de la cordillera. Estandartes dorados con la figura de un león rampante coronado ondeaban con orgullo en la brisa.",
        "Eran los caballeros del Reino.",
        "El Alto Mando y la Corona habían despachado a **cuatrocientos caballeros de la Orden del León** para socorrer y proteger el bastión de la Roca Carmesí. No eran paladines imbuidos en plegarias celestiales ni magos de salón cortesano; eran caballeros de hierro puro, una orden de caballería pesada forjada en las guerras de frontera, montados en colosales corceles de combate cubiertos de bardas de acero remachado y armados con lanzas de justa de tres metros con puntas de acero templado.",
        "—¡¡POR EL REINO Y EL LEÓN!! —resonó el grito unísono del comandante de la orden en lo alto de la colina.",
        "*¡¡¡BOOOOOOM-BOOOOM-BOOOOM!!!*",
        "Cuatrocientos corceles blindados descendieron por la pendiente como una avalancha de metal y furia.",
        "El choque fue descomunal. La caballería de la Orden del León se estrelló contra la retaguardia y los flancos desprotegidos de los Skavens a una velocidad arrolladora. Las lanzas ensartaban a tres y cuatro alimañas a la vez antes de quebrarse; los cascos herrados de los caballos aplastaban cráneos y corazas menores en el fango, y los jinetes desenvainaron pesadas espadas de arzón para comenzar a masacrar a los asaltantes con una disciplina quirúrgica e implacable.",
        "La masacre fue instantánea y abrumadora. Las formaciones de esclavos y alimañas, que dependían de la superioridad numérica para abrumar a las fortalezas, fueron barridas como hojas secas bajo la tormenta.",
        "* * *",
        "El impacto en la moral del enemigo fue letal.",
        "Y para el Señor de la Guerra Skaven, fue su sentencia de muerte.",
        "Al escuchar el rugido de la caballería pesada y contemplar con ojos desorbitados cómo su retaguardia era despedazada en cuestión de minutos, un terror visceral se apoderó de su semblante. Por una fracción de segundo crítica, el monstruo desvió la mirada de nuestro duelo, titubeó en su guardia y bajó la punta de su alabarda para buscar una vía de retirada.",
        "En un combate contra artistas marciales y guerreros templados en el infierno, distraerse un segundo equivale a entregar el cuello.",
        "—¡¡AHORA!! —rugí con toda la fuerza de mis pulmones.",
        "Nadie necesitó una segunda orden.",
        "Caesar King cargó con el borde inferior de su escudo directo contra las rodillas de la montura, partiéndole las articulaciones con un crujido seco que hizo que la gigantesca bestia colapsara de costado sobre el barro. El Señor de la Guerra trastabilló perdiendo el equilibrio.",
        "Vlad apareció a su diestra con un tajo ascendente de su espada bastarda que cercenó limpiamente el antebrazo derecho del Skaven, haciendo caer la alabarda envenenada al suelo. Casi en el mismo milisegundo, Ken se impulsó con una deflagración ígnea y le clavó un derechazo ardiente en el pecho que fundió las placas de bronce brujo y pulverizó su caja torácica.",
        "El caudillo rata soltó un alarido agónico, tambaleándose de rodillas.",
        "Di tres zancadas explosivas, cerrando la distancia con la cadencia perfecta del Fa Jin. Comprimí todo el maná telúrico y el calor ardiente de mis venas en mi antebrazo derecho y en mi rodilla.",
        "Salté hacia adelante, clavando mi rodillazo en el esternón del monstruo para quebrar su postura restante, y con el giro descendente de mi cuerpo descargué un codazo demoledor en el puente de su cráneo con todo el peso de mis noventa kilos reforzados con masa mágica.",
        "*¡¡¡KRRR-AAAACKKKK!!!*",
        "El cráneo del Señor de la Guerra Skaven estalló contra los adoquines ensangrentados.",
        "Su enorme cuerpo se sacudió una última vez en espasmos involuntarios antes de quedar completamente inerte en un charco de sangre negra y humo esmeralda.",
        "* * *",
        "El efecto fue inmediato y absoluto.",
        "Con la cabeza decapitada y la mente del comandante destruida, el vínculo de la mente colmena colapsó al instante. Cientos de ratas sobrevivientes soltaron sus armas con chillidos de pánico enloquecido, pisoteándose entre sí mientras intentaban huir hacia las grietas y túneles del desfiladero para escapar de la matanza.",
        "Sin embargo, el peligro no había terminado por completo.",
        "Un contingente de élite Skaven —cerca de doscientos Acechadores y Alimañas supervivientes armados con hojas ponzoñosas— intentó una maniobra desesperada de repliegue por el flanco oeste de la fortaleza, buscando desbordar los parapetos bajos donde aún se encontraban las barricadas secundarias y el acceso a los túneles médicos.",
        "No pasaron de los primeros metros.",
        "Una lluvia de acero resplandeciente cayó del cielo como un meteoro de espadas.",
        "—¡Nadie toca las murallas que mis camaradas defendieron con su sangre! —la voz imponente y regia de **Erza Scarlet** retumbó como un trueno sobre el desfiladero.",
        "Erza apareció plantada en mitad del flanco roto. Su figura brillaba con el fulgor sagrado de su *Armadura de la Rueda del Cielo*: decenas de espadas de plata pura flotaban en un círculo perfecto a su alrededor, obedeciendo a su voluntad como extensiones de su propia alma. Con un movimiento elegante de sus brazos, Erza liberó la ráfaga de acero: las espadas surcaron el aire a velocidades supersónicas, clavando a decenas de asaltantes contra las rocas y despedazando la vanguardia del flanco.",
        "Y entre las sombras que dejaba la lluvia de metal, una silueta oscura remató la masacre con una precisión quirúrgica.",
        "**Akame** emergió detrás de las líneas enemigas como una ráfaga de viento invisible. Su cabello negro se mecía con la brisa y sus ojos escarlata brillaban con un foco implacable. Cada paso suyo era un compás de muerte: su katana *Murasame* trazaba destellos carmesí que abrían gargantas y hendían corazones sin emitir el más mínimo sonido.",
        "—*Eliminar... eliminar... eliminar* —susurraba Akame con voz gélida.",
        "La combinación fue perfecta: la potencia destructiva de Erza Scarlet conteniendo el frente y la letalidad sigilosa de Akame cerrando las rutas de escape transformaron el intento de flanqueo de los Skavens en un callejón sin salida.",
        "Ambas guerreras protegieron los flancos con una sincronía magistral, repeliendo hasta la última alimaña que intentó acercarse a la fortaleza.",
        "* * *",
        "Un silencio denso y sobrecogedor descendió finalmente sobre el cañón de la Roca Carmesí.",
        "El humo de la pólvora rúnica y el hedor del azufre comenzaron a disiparse lentamente con el viento fresco de la alborada. En el horizonte, los primeros rayos del sol de la mañana despuntaban sobre las cumbres de la cordillera, iluminando las banderas ensangrentadas y los cuatrocientos caballeros de la Orden del León que terminaban de acordonar el perímetro exterior.",
        "Los cadetes de la fortaleza, cubiertos de fango, hollín y sangre enemiga, comenzaron a alzarse poco a poco.",
        "Primero fue un murmullo incrédulo. Luego, un alarido de júbilo ronco que creció hasta convertirse en un clamor triunfal que hizo retumbar las piedras de la fortaleza entera.",
        "Habíamos ganado.",
        "Con un esfuerzo sobrehumano, resistiendo al abismo, soportando el veneno y dejando cada gota de sudor en las piedras de la muralla, la victoria era nuestra.",
        "Me dejé caer sobre una roca al pie del portón, con la respiración entrecortada y los nudillos humeando por el calor del maná residual. A mi lado, Vlad apoyó su espada en la tierra suspirando con alivio, Caesar desató las correas de su escudo con una sonrisa exhausta y Ken se sentó en el fango secándose la frente con el dorso de la mano.",
        "Desde la entrada de la fortaleza, Kaoruko, Elysea y Lortelle asomaron la mirada hacia el patio. Al vernos en pie entre los restos de la horda, las lágrimas de alivio brillaron en sus ojos.",
        "La pesadilla de los Skavens había sido aplastada.",
        "Y la leyenda de los defensores de la Roca Carmesí apenas comenzaba a escribirse."
    ]
}

def compilar_capitulo_47():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 47              ")
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
        CAPITULO_45, CAPITULO_46, CAPITULO_47
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 47
    txt_cap47 = os.path.join(capitulos_dir, "47_Capitulo_47.txt")
    with open(txt_cap47, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_47['header'].upper()}: {CAPITULO_47['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_47['paragraphs']))
    print(f"[OK] Texto plano Capítulo 47 guardado en:\n  {txt_cap47}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 47
    docx_cap47 = os.path.join(capitulos_dir, "47_Capitulo_47.docx")
    create_single_chapter_document(CAPITULO_47, docx_cap47, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 47 guardado en:\n  {docx_cap47}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 47)
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
    compilar_capitulo_47()
