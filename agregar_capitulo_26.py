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

CAPITULO_26 = {
    "header": "Capítulo 26",
    "title": "Danza de Acero y Fuego",
    "paragraphs": [
        "El Gran Coliseo Central de Solaria amaneció bajo una expectación que rozaba la histeria colectiva.",
        "Con tres victorias consecutivas en los duelos individuales, la Clase F-2 se había convertido en el fenómeno absoluto del torneo. Los vendedores callejeros de la capital comerciaban con banderines improvisados con nuestra insignia y las gradas populares coreaban los nombres de Ken, Lighter y el mío como si fuéramos gladiadores consagrados.",
        "Sin embargo, el torneo entraba ahora en una etapa completamente diferente: la modalidad por parejas, el combate dos contra dos.",
        "Aquí ya no bastaba con la brillantez solitaria ni con el instinto salvaje de un solo hombre. Las batallas en parejas exigían sincronización táctica, cobertura de puntos ciegos y una lectura instantánea de los movimientos del compañero.",
        "* * *",
        "En nuestra enfermería de campaña subterránea, Ken estaba sentado sobre una camilla de madera con el ceño fruncido y los brazos cruzados con terquedad infantil.",
        "Tenía el torso vendado con gasas frescas impregnadas en ungüento de aloe y Kaoruko le sostenía el hombro con firmeza para impedir que hiciera fuerza.",
        "—¡Les juro que ya no me duele nada! —protestaba Ken, intentando mover el brazo derecho antes de que un pinchazo en las costillas le hiciera soltar un bufido doloroso—. ¡Déjenme salir a pelear hoy! ¡Quiero romper escudos con ustedes!",
        "—¡Te quedas exactamente donde estás, cabeza de fósforo! —le espetó Burnice, amenazándolo con una llave inglesa en la cara mientras le acomodaba una almohada a la fuerza—. Si vuelves a abrirte los puntos de la espalda antes del combate de cuatro, ¡te juro que te ato a la pata de la cama con alambre de púas!",
        "—Hazle caso a la niña del fuego, Ken —dijo Lighter con una sonrisa relajada, recostado contra el marco de la puerta mientras saboreaba una manzana—. Yo ya me quité las ganas de boxear ayer y tú diste el espectáculo del siglo el primer día. Tómate un respiro. Los combates de tres y cuatro aún están por delante y te necesitaremos al cien por cien.",
        "Erza Scarlet dio un paso al frente.",
        "Llevaba su armadura plateada de combate perfectamente pulida, con el emblema imperial cubriéndole el pecho y su larga melena escarlata recogida en una coleta alta que caía como una cascada sobre su espalda. En su cintura descansaba su espada larga habitual y su mirada irradiaba una determinación marcial tan pura que el aire de la habitación pareció adquirir peso.",
        "—Ken debe descansar —dictaminó Erza con su voz de mando inapelable—. La alineación para el combate de parejas está decidida desde anoche.",
        "Erza se giró hacia mí y me clavó sus intensos ojos castaños.",
        "—Silva. Tú y yo seremos la punta de lanza de la F-2 el día de hoy.",
        "Me puse de pie con calma, ajustándome los vendajes de cuerda con polvo de hierro en los antebrazos.",
        "—A tu señal, comandante —respondí con una media sonrisa tranquila—. Vamos a mostrarles cómo pelea la vanguardia.",
        "Kaoruko se acercó a nosotros dos, juntando las manos sobre el pecho con una mirada llena de admiración y cariño.",
        "—Seier-kun, Erza-san... por favor, cuídense el uno al otro —nos pidió con su voz dulce—. Sé que hacen una combinación invencible, pero no se confíen.",
        "—Descuida, Kaoruko —le aseguró Erza, dedicándole una sonrisa protectora—. Con Silva cubriendo mi flanco, ninguna espada en este reino puede alcanzarnos.",
        "* * *",
        "El sorteo de los duelos de dos contra dos nos emparejó contra la formidable Clase A-2, el escuadrón élite liderado por Elysea de Vane.",
        "Cuando cruzamos las rejas de hierro hacia la arena, el rugido de las gradas fue demoledor.",
        "—¡EN LA ESQUINA NORTE! —anunció la voz mágica del oficial—. ¡La formación táctica de asedio de la prestigiosa Clase A-2! ¡El caballero de la lanza pesada, Derek de Montclair... acompañado por la virtuosa del arco rúnico, Livia von Falken!",
        "Por el túnel norte emergieron dos combatientes cuya coordinación era legendaria en los registros de la academia.",
        "Derek era un guerrero robusto enfundado en una coraza pesada de bronce, armado con una gigantesca lanza de torneo de tres metros y un escudo redondo imbuido con runas de maná de gravedad. A su lado caminaba Livia, una arquera elfa de porte elegante que sostenía un arco largo de madera plateada, con un carcaj repleto de flechas de viento y luz que centelleaban con energía letal.",
        "—¡Y EN LA ESQUINA SUR! —rugió el anunciador con un tono vibrante—. ¡La imparable vanguardia de la Clase F-2! ¡La comandante de armas Erza Scarlet... y el demoledor invicto Seier Silva!",
        "La multitud estalló en vítores.",
        "En el palco superior de la Clase A-2, divisé a Elysea de Vane sentada en primera fila con las piernas cruzadas. La capitana elfa nos observaba con los ojos entrecerrados y una sonrisa calculadora; este combate era su prueba de fuego para medir la capacidad de adaptación de nuestro equipo.",
        "Derek plantó su lanza pesada en la tierra con un crujido sordo, apuntando la punta hacia nosotros.",
        "—Reconocemos su poder individual —dijo Derek con tono solemne y marcial—. Pero el combate por parejas es un arte de complementariedad. Un duelista de puños y una espadachina frontal no tienen forma de penetrar una muralla de gravedad mientras la lluvia de luz los diezma desde la distancia.",
        "Erza desenvainó su espada larga con un silbido metálico impecable.",
        "—Entonces será mejor que afiances bien tus talones, caballero —respondió Erza con serenidad cortante—. Porque mi acero y los puños de Silva no saben lo que es detenerse ante una muralla.",
        "—¡COMIENCEN! —bramó el árbitro.",
        "* * *",
        "La pareja de la A-2 ejecutó su estrategia con precisión militar.",
        "Derek golpeó el suelo con la base de su lanza, activando una red de runas púrpuras bajo nuestros pies:",
        "—¡CAMPO GRAVITATORIO: POZO DE PLOMO!",
        "La gravedad en nuestra mitad de la arena se multiplicó por tres en una milésima de segundo. El aire se volvió espeso y la arena pareció transformarse en lodo asfáltico, buscando clavar nuestras botas en el suelo.",
        "En el mismo instante, Livia saltó hacia atrás con la ligereza de una pluma, tensó la cuerda de su arco plateado y disparó una ráfaga continua de proyectiles rúnicos:",
        "—¡LLUVIA CELESTIAL: CIEN FLECHAS DE LUZ!",
        "El cielo de la arena se cubrió de una cortina dorada de proyectiles centelleantes que descendían en picado como meteoritos afilados directo sobre nuestras cabezas.",
        "Para cualquier dúo ordinario, aquella trampa de inmovilización combinada con bombardeo aéreo habría significado una derrota instantánea.",
        "Pero Erza y yo nos entendíamos sin necesidad de mediar palabra.",
        "—¡SILVA, EL FRENTE ES TUYO! —gritó Erza.",
        "—¡VOY!",
        "Erza concentró su maná mágico y su cuerpo fue envuelto en un destello de luz carmesí deslumbrante:",
        "—¡REQUIP: ARMADURA DE LA RUEDA DEL CIELO!",
        "En una fracción de segundo, su coraza de plata se transformó en un peto alado con cuatro alas de acero afilado a su espalda, y una corona de más de veinte espadas mágicas levitó a su alrededor en una formación circular perfecta.",
        "Erza saltó hacia el aire, desafiando el pozo gravitatorio con la fuerza de sus alas de acero, y desplegó su arsenal contra el bombardeo enemigo:",
        "—¡ESPADARES DE DISPERSIÓN!",
        "¡CLANG-CLANG-CLANG-CLANG!",
        "Las veinte espadas de Erza interceptaron cada una de las flechas de luz de Livia en el aire, pulverizándolas en una lluvia inofensiva de chispas doradas y despejando el cielo por completo.",
        "* * *",
        "Con el bombardeo neutralizado por Erza, llegó mi turno.",
        "Canalicé mi maná de **Tierra** a través de las piernas y las caderas. La gravedad artificial de Derek intentaba aplastarme, pero mi entrenamiento de diez años y la afinidad elemental me permitieron fusionar mi propio peso con el suelo, utilizando la presión añadida para cargar mis músculos como resortes de hierro comprimido.",
        "Exploté hacia adelante con un paso de rotación explosivo, reventando el campo gravitatorio bajo mis pies en una estela de arena:",
        "¡FWOOOOSH!",
        "Cruzé los quince metros que me separaban de Derek en un pestañeo.",
        "Derek abrió los ojos con pánico y reaccionó con reflejos admirables: levantó su escudo de gravedad y lanzó una estocada horizontal con su lanza de tres metros directo a mi garganta.",
        "No retrocedí.",
        "Con un desvío oblicuo de antebrazo de Jeet Kune Do, desvié la punta de la lanza por encima de mi hombro. El acero rozó el aire a centímetros de mi cuello, y en ese mismo segundo, penetré en su distancia corta con un rodillazo devastador de Muay Thai directo a su escudo redondo.",
        "¡¡¡BOOOOOOM!!!",
        "La fuerza de mis noventa kilos reforzados con maná de tierra hizo que el escudo de bronce crujiera y las rodillas de Derek se doblaran hacia el suelo.",
        "—¡Livia, dispárale ahora! —gritó el lancero, intentando sujetar mi pierna.",
        "Desde la retaguardia, Livia ya tenía una flecha perforante de viento cargada en el arco, apuntando directo a mi espalda descubierta.",
        "Pero no estaba descubierta.",
        "Una hoja de acero plateado cortó el aire a la velocidad del rayo, clavándose en la arena a escasos milímetros del pie de Livia y obligándola a cancelar el tiro para esquivar un corte descendente.",
        "Erza había descendido en picado detrás de mí, cubriendo mi retaguardia con una precisión matemática.",
        "—Tu oponente soy yo, arquera —dijo Erza con la espada en guardia alta.",
        "* * *",
        "La formación de la Clase A-2 estaba completamente rota.",
        "Derek intentó empujarme con el escudo para recuperar la distancia de su lanza, pero atrapé el borde superior de su defensa con la mano izquierda, tirando de él hacia adelante para descolocar su centro de gravedad.",
        "—Erza, ¡ahora! —le grité.",
        "Erza ejecutó un giro sobre sus talones, desarmó el arco de Livia con un tajo de revés que quebró la cuerda de plata y se posicionó a mi costado derecho en perfecta sincronía.",
        "Su armadura volvió a brillar con un resplandor ígneo:",
        "—¡REQUIP: EMPERATRIZ DEL FUEGO!",
        "Una coraza carmesí con alas de dragón y dos espadas gemelas incandescentes aparecieron en sus manos.",
        "A su lado, yo encendí el fuego concentrado en mi puño derecho, combinándolo con la masa aplastante de la tierra en mi talón.",
        "Ambos nos movimos al unísono, como si un solo cerebro gobernara nuestros cuerpos.",
        "—¡TÉCNICA COMBINADA: CRUZ CARMESÍ DE ACERO Y FUEGO! —rugimos al mismo tiempo.",
        "Erza descargó un corte cruzado con sus dos espadas ardientes en forma de 'X', mientras yo lanzaba una patada giratoria en salto con la tibia envuelta en una llamarada colosal.",
        "¡¡¡KRAAA-BOOOOOOOOOOOOOOOM!!!",
        "La explosión combinada de fuego elemental y choque de acero resonó como una bomba en el centro del coliseo.",
        "La lanza pesada de Derek se partió en dos pedazos y su escudo salió volando por los aires; Livia fue levantada por la onda expansiva térmica y ambos fueron proyectados hacia atrás a más de diez metros, impactando contra la base del muro perimetral antes de caer tendidos sobre la arena, completamente desarmados e inconscientes.",
        "* * *",
        "El humo carmesí se disipó bajo la brisa del mediodía.",
        "Erza enfundó sus espadas gemelas con un chasquido impecable y su armadura volvió a su aspecto plateado ceremonial. Yo sacudí el polvo de mis vendajes y me coloqué a su lado, con la respiración acompasada.",
        "El árbitro se quedó mirando los cuerpos de los dos campeones de la A-2 antes de alzar ambas manos hacia el cielo:",
        "—¡G-GANADORES POR NOCAUT TOTAL... LA CLASE F-2!",
        "El Gran Coliseo Central rugió como una bestia salvaje.",
        "Treinta mil personas aplaudían de pie, fascinadas por la demostración de sincronización y poderío marcial que acababan de presenciar. No había habido fisuras, no había habido errores: éramos un muro impenetrable de espada y puño.",
        "En nuestro balcón, Ken se levantó de la camilla ignorando los regaños de Burnice y gritó con el puño en alto; Lighter aplaudía con una sonrisa satisfecha y Kaoruko lloraba de alegría abrazada a Subaru Hoshina.",
        "En el palco superior, Elysea de Vane se puso de pie lentamente.",
        "Sus ojos celestes brillaban con una intensidad ardiente. Me miró a los ojos a través de la arena y me dedicó un aplauso pausado, reconociendo sin palabras que la Clase F-2 era la mayor amenaza a la que se enfrentaría jamás.",
        "Erza levantó el puño cerrado hacia mí.",
        "—Buen combate, Silva.",
        "Chocamos los puños en el centro de la arena con una sonrisa cómplice.",
        "—El mejor, comandante.",
        "Cuatro victorias consecutivas en cuatro días de competencia.",
        "La Clase F-2 estaba a un solo paso de la gloria absoluta, y los combates de tres y cuatro guerreros estaban a punto de definir el destino del imperio."
    ]
}

def compilar_capitulo_26():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 26              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25, CAPITULO_26
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 26
    txt_cap26 = os.path.join(capitulos_dir, "26_Capitulo_26.txt")
    with open(txt_cap26, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_26['header'].upper()}: {CAPITULO_26['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_26['paragraphs']))
    print(f"[OK] Texto plano Capítulo 26 guardado en:\n  {txt_cap26}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 26
    docx_cap26 = os.path.join(capitulos_dir, "26_Capitulo_26.docx")
    create_single_chapter_document(CAPITULO_26, docx_cap26, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 26 guardado en:\n  {docx_cap26}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 26)
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
    compilar_capitulo_26()
