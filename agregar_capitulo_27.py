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

CAPITULO_27 = {
    "header": "Capítulo 27",
    "title": "Chispa y Ceniza",
    "paragraphs": [
        "La tarde posterior a nuestra victoria en parejas cerró la primera tanda de combates dos contra dos con una serie de choques violentos en el Gran Coliseo.",
        "Tal como esperábamos, las demás clases de la élite imperial no se quedaron de brazos cruzados:",
        "La Clase A-1, desesperada por lavar la humillación de sus derrotas individuales, desplegó a un dúo aterrador compuesto por un piromante de fuego oscuro y un asesino de sombras que destrozaron la formación de la Clase C-1 en menos de tres minutos, quemando la arena con crueldad despiadada para recordar a las gradas por qué ostentaban el rango supremo.",
        "Por su parte, la Clase B-1 logró redimirse tras la caída de Goran gracias a sus dos caballeros de asedio rúnico, quienes aplastaron a la pareja de la Clase D tras una agotadora guerra de desgaste con hachas de guerra y escudos de torre.",
        "Y en la Clase A-2, a pesar de que Erza y yo habíamos eliminado a Derek y Livia, la facción de Elysea de Vane demostró tener un banquillo profundo: sus otros dos representantes se impusieron con solvencia ante la Clase B-2, manteniendo a su grupo firmemente en la pelea por los primeros puestos de la tabla general.",
        "La noche cayó sobre la capital dejando el tablero al rojo vivo. Al día siguiente se disputaría la segunda y última jornada de los duelos de dos contra dos.",
        "Y el reglamento del torneo era inflexible: ningún cadete que hubiera combatido el día anterior en la misma modalidad podía repetir.",
        "Erza y yo estábamos automáticamente inhabilitados.",
        "* * *",
        "A primera hora de la mañana, en nuestra sala de preparación, la tensión se palpaba en el aire.",
        "—Nuestros rivales para hoy serán los Hermanos Valenwood, de la Clase B-1 —explicó Natsuki frente a la pizarra táctica—. Dos hermanos gemelos que combaten con alabardas rúnicas de dos metros imbuidas en maná de relámpago y viento. Tienen una sincronización casi telepática: uno barre el suelo con cortes bajos y el otro decapita en el aire. Son especialistas en despedazar formaciones desorganizadas.",
        "—Silva y yo estamos fuera por reglamento —dijo Erza, sentada con los brazos cruzados junto a la mesa—. Eso significa que Lighter debe liderar la pareja.",
        "Lighter estaba apoyado contra la pared, con los puños vendados y sus gafas oscuras puestas, asintiendo con una media sonrisa tranquila.",
        "—Por mí no hay problema —dijo el boxeador—. Pero sigo necesitando a alguien que me cubra la espalda en la arena.",
        "Todas las miradas se desviaron hacia la camilla de la enfermería.",
        "Allí estaba Ken, de pie sobre el suelo de piedra, flexionando los brazos y saltando sobre las puntas de los pies como un resorte cargado de pólvora.",
        "Kaoruko acababa de retirarle las últimas vendas del hombro y del torso. La piel del muchacho aún mostraba marcas rojizas de las quemaduras de su primer combate, pero los cortes estaban completamente sellados gracias a la devoción y la magia restauradora de la chica azabache.",
        "—Estás al ochenta por ciento, Ken-kun... —advirtió Kaoruko con una mezcla de orgullo y preocupación en sus ojos oscuros—. Tus canales de maná ya no están saturados, pero tus músculos todavía necesitan recuperarse del esfuerzo del Fénix.",
        "—¡El ochenta por ciento de mí es más que suficiente para hacer temblar a este coliseo! —rugió Ken con una sonrisa salvaje, atándose la bandana roja sobre la frente con un tirón firme—. ¡Llevo dos días sentado viendo cómo se divierten! ¡Si no salgo a pelear hoy, voy a explotar de pura frustración!",
        "Burnice se acercó a él de un salto y le ajustó con fuerza las correas de cuero de sus brazaletes de hierro con púas.",
        "—¡Más te vale no hacer el ridículo, héroe! —le dijo la rubia con un guiño alegre y las mejillas encendidas—. ¡Si ganas este combate junto a Lighter, te cocino una bandeja entera de pastelillos de manzana con extra de canela!",
        "—¡Trato hecho, capitana del nitro! —rió Ken chocando los puños con ella.",
        "Lighter se despegó de la pared con pasos lentos y se colocó al lado del chico de pelo azul.",
        "—Escucha, fósforo con patas —dijo Lighter con tono sereno—. Esos dos tipos de las alabardas van a querer mantenernos a distancia con sus dos metros de acero. Tú eres el aire y el fuego; yo soy el suelo y el contragolpe. Tú creas el caos por arriba y yo les apago las luces por abajo.",
        "—Me gusta ese plan, boxeador —sonrió Ken, haciendo crepitar una chispa dorada en sus nudillos—. Vamos a enseñarles cómo pelea la calle.",
        "* * *",
        "Cuando las compuertas de hierro se abrieron hacia el Gran Coliseo, el clamor del público sacudió las vigas del estadio.",
        "—¡EN LA ESQUINA NORTE! —bramó el anunciador imperial—. ¡La guardia de alabardas gemelas de la Clase B-1! ¡Los hermanos Torsten y Kaelen de la Casa Valenwood!",
        "Por el túnel norte avanzaron dos guerreros idénticos, de cabello castaño y corazas doradas, portando sendas alabardas de acero rúnico cuyas hojas vibraban con un zumbido eléctrico y torbellinos de viento cortante.",
        "—¡Y EN LA ESQUINA SUR! —rugió la voz mágica con entusiasmo desbordante—. ¡El regreso del Fénix Indómito, Ken... acompañado por el maestro del contragolpe de Calydon, Lighter!",
        "Treinta mil personas estallaron en aplausos ensordecedores.",
        "En nuestro balcón, Kaoruko apretaba las manos contra su pecho rezando en silencio; Subaru Hoshina observaba a Lighter desde la barandilla con una mirada atenta y serena; y Erza y yo manteníamos la vista fija en la arena, confiando ciegamente en nuestros dos compañeros.",
        "Torsten Valenwood hizo girar su alabarda con una maestría visual impecable y apuntó la punta hacia nuestros chicos.",
        "—Pensamos que enviarían a Silva o a Scarlet —dijo el noble con desdén—. En cambio, nos mandan a un boxeador sin armas y a un salvaje herido que apenas puede sostenerse en pie. La Clase B-1 limpiará el suelo con ustedes en un minuto.",
        "Lighter se acomodó las gafas oscuras sobre el puente de la nariz y adelantó el pie izquierdo en guardia clásica.",
        "—Tienen dos metros de ventaja de alcance, muchachos —dijo Lighter con calma glacial—. Será mejor que no pestañeen, porque cuando crucemos esa distancia, el acero no les va a servir de nada.",
        "—¡COMIENCEN! —rugió el árbitro oficial.",
        "* * *",
        "Los hermanos Valenwood cargaron al unísono con una coordinación terrorífica.",
        "—¡TÉCNICA DE ALABARDA GEMELA: MOLINO DE VIENTO Y TRUENO!",
        "Torsten barrió el suelo con un corte circular rasante cargado de viento comprimido, buscando amputar las piernas de ambos; mientras Kaelen saltaba un metro en el aire descargando una estocada descendente electrificada directo al pecho de Lighter.",
        "Era una pinza de doble altura diseñada para no dejar espacio de maniobra.",
        "Pero Lighter y Ken no peleaban con libros de texto.",
        "—¡KEN, ARRIBA! —ordenó Lighter con voz seca.",
        "—¡DÉJAMELO A MÍ!",
        "Ken no intentó esquivar hacia atrás. Flexionó las piernas y se disparó en vertical con un salto acrobático descomunal, pasando limpiamente por encima del barrido bajo de Torsten.",
        "En el aire, Ken interceptó la estocada descendente de Kaelen con una patada de hacha envuelta en fuego puro:",
        "¡CLAAAANG!",
        "La suela reforzada de Ken golpeó el asta de la alabarda electrificada, desviándola en el aire y forzando a Kaelen a descender de golpe contra la arena con los brazos temblando por la vibración térmica.",
        "Al mismo tiempo, en el suelo, Lighter se encargó de Torsten.",
        "Con un paso deslizante milimétrico hacia adelante, Lighter esquivó el segundo giro del arma del noble inclinando el torso por debajo de la hoja. Entró directo en la distancia corta —el punto ciego absoluto de un arma de dos metros— y conectó un jab de izquierda seco y cegador en la nariz de Torsten:",
        "¡PAF!",
        "La cabeza del noble se sacudió violentamente hacia atrás, soltando un chorro de sangre y haciéndolo trastabillar dos pasos.",
        "—¡Maldito...! —rugió Torsten, intentando retroceder para usar la punta de su alabarda.",
        "Pero Lighter no lo dejó respirar.",
        "Avanzó pegado a su pecho como una sombra, soltando una combinación demoledora de tres golpes al cuerpo: costillas flotantes, plexo y estómago. Cada impacto resonaba como un martillazo contra el metal de su coraza.",
        "* * *",
        "Mientras Lighter castigaba a Torsten, Kaelen intentó socorrer a su hermano.",
        "Canalizó todo su maná de relámpago en el asta de su alabarda y lanzó una estocada traicionera hacia la espalda de Lighter:",
        "—¡SUÉLTALO, BASTARDO!",
        "Lighter ni siquiera se giró.",
        "Confiaba plenamente en su compañero.",
        "Desde el cielo, una estela carmesí cayó en picado como un bólido ardiente.",
        "—¡TU PELEA ES CONMIGO, CABEZA HUECA! —bramó Ken.",
        "Ken descendió en barrena girando sobre su propio eje, envuelto en un torbellino de llamas doradas: *¡Técnica del Dragón Ígneo: Taladro Solar!*",
        "El talón en llamas de Ken impactó directo sobre la punta de la alabarda de Kaelen en el aire, pulverizando la carga eléctrica en una lluvia de chispas y estampando el arma del noble contra el suelo con tanta violencia que la madera rúnica se astilló en pedazos.",
        "Kaelen cayó de rodillas sobre la arena, aturdido por la onda de choque térmica y con las manos quemadas por el calor residual.",
        "La formación de los hermanos Valenwood estaba completamente desmantelada.",
        "—¡Ahora, Lighter! ¡El golpe final! —gritó Ken, tomando impulso en la arena para su último asalto.",
        "—A la de tres, muchacho —respondió Lighter con una sonrisa serena.",
        "Los dos guerreros de la Clase F-2 coordinaron su remate en perfecta sintonía:",
        "Por la izquierda, Ken se elevó dos metros en el aire con un giro acrobático brutal, concentrando todo su fuego en su pierna derecha en un barrido ascendente: *¡Furia del Dragón Aéreo!*",
        "Por la derecha, Lighter hundió el talón en la tierra, flexionó la cintura y desató el golpe más devastador de su repertorio de pugilismo: un *Derechazo Cruzado de Calydon* lanzado con todo el peso de su cuerpo y una velocidad sobrehumana.",
        "Ambos impactos conectaron al mismo tiempo sobre los dos nobles acorralados:",
        "¡¡¡BOOOOOOOOOOOOOOM!!!",
        "Ken impactó en el mentón de Kaelen con una llamarada explosiva que lo mandó a volar diez metros hacia el fondo de la arena.",
        "Lighter conectó en la mandíbula de Torsten con un golpe seco que le arrancó el yelmo y lo dejó completamente noqueado en el aire antes de que su cuerpo cayera pesadamente contra la arena blanca.",
        "* * *",
        "El humo de la pólvora y el calor se disiparon bajo la brisa fresca del coliseo.",
        "Los dos hermanos Valenwood yacían profundamente inconscientes en extremos opuestos del foso de batalla, con sus alabardas quebradas sobre la tierra.",
        "Y en el centro de la arena, respirando con fuerza pero sin una sola herida de gravedad, estaban Lighter y Ken.",
        "Ken se secó el sudor de la frente con el dorso del brazo, soltó una carcajada alegre y chocó los cinco con Lighter con una palmada sonora que resonó por todo el estadio.",
        "—¡Te dije que ese ochenta por ciento bastaba, viejo! —festejó Ken entre risas.",
        "—No estuvo mal para un chiquillo con vendajes —bromeó Lighter, acomodándose las gafas oscuras y dedicándole una media sonrisa satisfecha.",
        "El árbitro principal levantó ambos brazos hacia el cielo sin dudar un solo segundo:",
        "—¡G-GANADORES POR NOCAUT TOTAL... KEN Y LIGHTER, DE LA CLASE F-2!",
        "El Gran Coliseo Central rugió en una ovación atronadora.",
        "Las gradas estaban de pie celebrando una victoria que nadie creía posible: la Clase F-2 había ganado los dos combates de la modalidad por parejas, manteniendo su récord absolutamente perfecto en el torneo.",
        "En nuestro palco, Kaoruko saltaba de alegría con las manos en el pecho; Burnice aplaudía a rabiar agitando una bandera improvisada; y Subaru Hoshina sonreía abiertamente, mirando a Lighter con un orgullo radiante en sus ojos verdes.",
        "Erza y yo nos miramos con una sonrisa cómplice.",
        "Las dos primeras fases del torneo estaban selladas: invictos en duelos individuales y e invictos en combates de dos.",
        "Ahora, los temidos combates de escuadrón de tres y cuatro guerreros estaban por comenzar.",
        "Y la Clase F-2 ya no era una sorpresa: éramos los reyes indiscutidos de la arena."
    ]
}

def compilar_capitulo_27():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 27              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25, CAPITULO_26, CAPITULO_27
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 27
    txt_cap27 = os.path.join(capitulos_dir, "27_Capitulo_27.txt")
    with open(txt_cap27, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_27['header'].upper()}: {CAPITULO_27['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_27['paragraphs']))
    print(f"[OK] Texto plano Capítulo 27 guardado en:\n  {txt_cap27}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 27
    docx_cap27 = os.path.join(capitulos_dir, "27_Capitulo_27.docx")
    create_single_chapter_document(CAPITULO_27, docx_cap27, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 27 guardado en:\n  {docx_cap27}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 27)
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
    compilar_capitulo_27()
