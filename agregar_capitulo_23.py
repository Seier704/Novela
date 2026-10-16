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

CAPITULO_23 = {
    "header": "Capítulo 23",
    "title": "Ráfaga de Fénix",
    "illustration": os.path.join(NOVELA_DIR, "imagenes", "ken_fenix.png"),
    "illustration_caption": "Ilustración: Ken desatando el Espíritu del Fénix (Clase F-2)",
    "paragraphs": [
        "El sol matutino ascendió sobre la capital imperial de Solaria como un escudo de bronce al rojo vivo.",
        "El Gran Coliseo Central estaba a rebosar. Más de veinte mil espectadores abarrotaban las gradas escalonadas de mármol: nobles aristócratas con túnicas de seda, oficiales de la guardia real, mercaderes influyentes y todo el cuerpo estudiantil de la academia. El rugido colectivo de la multitud hacía vibrar el polvo de la arena y los estandartes heráldicos ondeaban bajo una brisa cargada de tensión mágica.",
        "El primer día del Gran Torneo Interclases estaba reservado exclusivamente para los duelos individuales: la modalidad uno contra uno.",
        "Y para inaugurar el torneo en nombre de la Clase F-2, no dudamos ni un segundo en elegir a nuestro primer guerrero.",
        "Ken había dado un paso al frente desde el primer minuto. Su deseo de volverse más fuerte no era vanidad ni orgullo ciego; era un fuego voraz que le quemaba en las entrañas, una necesidad primaria de medirse contra los más grandes del imperio para demostrar el valor de su disciplina.",
        "Sin embargo, el sorteo oficial del cuadro no tuvo piedad.",
        "Nuestro primer oponente pertenecía a la temida Clase A-1: la élite absoluta de la academia, el pináculo de la aristocracia marcial de Solaria.",
        "* * *",
        "—¡En la esquina norte, representando a la invicta Clase A-1! —atronó la voz del presentador a través de los amplificadores de maná—. ¡El prodigio del esgrima mágico, primogénito de la ilustre Casa Montfort... Julian de Montfort!",
        "Las gradas superiores estallaron en aplausos ensordecedores y vítores aristocráticos.",
        "Por el túnel norte emergió un joven esbelto y de porte altivo, envuelto en una capa blanca con bordados de zafiro. Su uniforme militar impecable relucía bajo la luz del sol y en su mano derecha sostenía un estoque ceremonial de acero rúnico. A su alrededor danzaban corrientes visibles de maná de viento comprimido, cortando el aire con un zumbido agudo y silbante.",
        "—¡Y en la esquina sur, representando a la Clase F-2...! —continuó el anunciador con un tono notablemente menos entusiasta—. ¡El artista marcial del fuego... Ken!",
        "A diferencia del recibimiento del noble, un coro de risas burlonas y silbidos despectivos resonó desde los sectores nobles.",
        "—¿Mandaron a un salvaje descalzo con una bandana roja? —se mofó un noble de la Clase A en las filas delanteras—. ¿Es un chiste de la clase de desechos?",
        "—Julian lo va a convertir en carne picada antes de que parpadee. Esto terminará en tres segundos.",
        "Ken no les prestó la menor atención.",
        "Caminó hacia el centro de la arena con pasos elásticos, descalzo sobre la arena caliente, con sus brazaletes de hierro con púas ajustados a las muñecas y su collar de grandes cuentas rojas colgando sobre su pecho desnudo. En su rostro no había miedo, ni titubeo, ni la habitual mueca cómica con la que solía bromear en los pasillos.",
        "Sus ojos carmesíes ardían con una intensidad feroz y concentrada.",
        "—¡PÁRTELO EN DOS, KEN-KUN! —gritó la voz desgarradora de Burnice desde nuestro palco, agitando las dos manos en el aire con un entusiasmo histérico—. ¡QUÉMALO HASTA QUE QUEDE COMO UNA TOSTADA!",
        "—¡Haz que respete nuestros puños, hermano! —secundó Lighter con los brazos cruzados y una sonrisa tranquila.",
        "—No bajes la guardia ni un milímetro, Ken —dijo Erza en voz baja, con la mano sobre la empuñadura de su espada en señal de respeto.",
        "Yo me mantuve en el borde de la barandilla con los brazos cruzados, observando la postura de mi amigo.",
        "Muchos en ese coliseo veían a Ken como un bruto ingenuo condenado al fracaso. Pero nosotros sabíamos lo que había dentro de ese chico: un espíritu inquebrantable que jamás retrocedía ante la adversidad.",
        "* * *",
        "Ambos combatientes quedaron frente a frente a diez metros de distancia en el círculo de batalla.",
        "Julian de Montfort desenvainó su estoque rúnico con un movimiento fluido y teatral. La hoja de plata vibró, envolviéndose en espirales de viento cortante y maná cian.",
        "—Es una lástima que la academia obligue a los caballeros de linaje a ensuciar su acero contra campesinos sin modales —dijo Julian con desdén gélido, apuntando la punta de su espada hacia la garganta de Ken—. Ríndete ahora. Mi técnica del 'Viento Castigador' no sabe tener compasión de la escoria.",
        "Ken no se inmutó. Adelantó la pierna izquierda, flexionó las rodillas bajando su centro de gravedad y levantó los puños vendados hacia el frente. Una chispa dorada crepitó entre sus dedos.",
        "—Hablas demasiado para ser un espadachín —respondió Ken con una sonrisa afilada—. Ven y comprueba si tu viento puede apagar mis llamas.",
        "—¡COMIENCEN! —rugió el árbitro oficial.",
        "El choque fue instantáneo y aterrador.",
        "Julian desapareció de su posición en una ráfaga de velocidad sobrehumana impulsada por maná eólico: *¡FSSSH!*",
        "Reapareció a un metro del flanco de Ken, descargando una estocada fulminante hacia su hombro. Ken esquivó por milímetros con un giro de torso, respondiendo con un gancho de derecha envuelto en fuego; pero Julian era un veterano del combate aristocrático: utilizó la corriente de su propia espada para desviar el puño de Ken en el aire y encadenó una combinación implacable de cinco estocadas de viento continuo.",
        "¡CLANG! ¡SWOOSH! ¡SWOOSH!",
        "Hojas de aire comprimido rasgaron el espacio. Los brazaletes de hierro de Ken bloquearon los dos primeros impactos con un estruendo metálico, pero las estocadas restantes le alcanzaron los muslos y el costado.",
        "La sangre brotó, manchando la tela de su pantalón.",
        "—¡Uf...! —soltó Ken, retrocediendo dos pasos sobre la arena mientras apretaba los dientes.",
        "Las gradas de los nobles estallaron en aplausos satisfechos.",
        "—¡Te lo dije! ¡No tiene oportunidad contra un espadachín de la A-1!",
        "Julian no le dio tregua. Avanzó con elegancia letal, trazando cortes geométricos que encerraron a Ken en una jaula de cuchillas de viento invisible.",
        "—¡Estilo Montfort: Tormenta de Cuchillas Espirales! —exclamó el noble, lanzando una estocada al pecho.",
        "Ken intentó contraatacar con una patada ascendente de fuego, pero el viento de Julian sofocó las llamas en el punto de contacto. La punta del estoque golpeó el hombro derecho de Ken con una fuerza demoledora reforzada por maná, desgarrándole la carne y mandándolo a estrellarse violentamente contra el muro de piedra del coliseo.",
        "¡BOOOM!",
        "Una nube de polvo y esquirlas de granito se levantó en el perímetro.",
        "* * *",
        "El silencio cayó sobre una parte del coliseo, roto únicamente por los gritos victoriosos de la Clase A-1.",
        "Ken estaba de rodillas en el suelo, con la cabeza gacha, el pecho cubierto de cortes sangrantes y el brazo derecho humeante por el impacto del maná eólico. La sangre goteaba desde su barbilla sobre la arena blanca.",
        "—¡Ken! —exclamó Kaoruko desde nuestro palco, tapándose la boca con las manos y los ojos llenos de angustia.",
        "—¡Levántate, idiota! ¡Tú no puedes caer aquí! —gritó Burnice, aferrándose a la barandilla con tanta fuerza que sus nudillos se pusieron blancos.",
        "Julian enfundó la mitad de su estoque con una sonrisa altanera y se dio media vuelta hacia el árbitro.",
        "—Declara el final, oficial. El salvaje no se levantará más. Este es el límite de los plebeyos.",
        "Sin embargo, el árbitro no levantó la bandera.",
        "Sus ojos estaban clavados en la nube de polvo detrás del noble.",
        "Desde el cráter en el muro de piedra, un sonido gutural comenzó a resonar. No era un lamento, ni un quejido de agonía.",
        "Era una risa.",
        "Una risa baja, áspera y cargada de una emoción desbordante que erizó la piel de todos los presentes en las primeras filas.",
        "—¿De verdad... crees que esto es mi límite? —resonó la voz de Ken.",
        "Julian se giró con el ceño fruncido.",
        "Entre los escombros de piedra, Ken se estaba poniendo de pie.",
        "Sus piernas temblaban por el dolor, pero sus talones se afirmaron en la tierra con la fuerza de un pilar de hierro. Con la mano izquierda, desató su bandana roja ensangrentada, se la ató de nuevo con fuerza sobre la frente apartando sus mechones azulados, y se limpió la sangre del labio con el dorso del brazo.",
        "En su mirada carmesí ya no quedaba rastro del muchacho bromista del aula.",
        "Había una determinación titánica, indómita y pura. Una chispa ancestral que ardía en lo más profundo de su alma, negándose a apagarse ante la tormenta.",
        "—Tus cortes duelen como una caricia comparados con las palizas que me daba mi maestro —dijo Ken, dando un paso al frente mientras el suelo bajo sus pies descalzos comenzaba a resquebrajarse por un calor abrasador.",
        "—¡Muérete de una vez! —rugió Julian, perdiendo por primera vez la compostura aristocrática ante la mirada del muchacho.",
        "El noble cargó con todo su maná, proyectando un tornado masivo de viento cortante alrededor de su espada.",
        "Ken no retrocedió. Cerró los ojos durante una milésima de segundo y conectó con la esencia más pura de su poder marcial.",
        "En su pecho, la silueta ardiente de un fénix celestial batió sus alas de fuego eterno.",
        "* * *",
        "¡KRAAAASH!",
        "Una columna colosal de fuego carmesí y dorado detonó alrededor de Ken, pulverizando la arena en un radio de cinco metros y convirtiendo el suelo en vidrio fundido.",
        "La presión del maná térmico fue tan abrumadora que el tornado de viento de Julian se desintegró al instante, consumido por una temperatura infernal que hizo retroceder a los espectadores de las primeras gradas por el calor sofocante.",
        "El cabello azul de Ken ondeaba hacia el cielo como si estuviera vivo. Sus ojos brillaban como dos soles en miniatura y una corona de llamas doradas danzaba sobre su cabeza, formando el contorno majestuoso de dos alas de fénix desplegadas a sus espaldas.",
        "—Imposible... —balbuceó Julian, pálido como un cadáver mientras el calor le derretía el filo del estoque—. ¡¿Qué demonios es ese monstruo?!",
        "—¡ESTE ES EL VERDADERO FUEGO! —rugió Ken con una voz que estremeció los cimientos del coliseo.",
        "Ken explotó hacia adelante.",
        "Su velocidad fue tan descomunal que dejó una estela de fuego en el aire. Esquivó la última estocada desesperada de Julian inclinando la cabeza un milímetro, entró en su guardia con un paso de rotación impecable y conectó un uppercut devastador directo en el plexo solar del noble:",
        "¡BOOOOOOM!",
        "El impacto sacudió el estadio. Julian salió despedido por los aires como un muñeco de trapo, tosiendo sangre y saliva mientras su barrera mágica estallaba en mil pedazos.",
        "Pero Ken no había terminado.",
        "Flexionó las piernas y se disparó hacia el cielo con un salto monumental, interceptando al noble en el aire a quince metros sobre la arena.",
        "En el cenit del vuelo, Ken concentró todo el fuego del fénix en sus dos puños y piernas, iluminando el coliseo entero con un resplandor dorado más brillante que el propio sol del mediodía.",
        "—¡ARDE HASTA LAS CENIZAS...! —bramó Ken con cada fibra de su ser—. ¡TÉCNICA DEFINITIVA DEL DRAGÓN ÍGNEO: RÁFAGA DE FÉÉÉÉÉÉENIX!",
        "Lo que siguió fue un espectáculo de pura magnificencia marcial.",
        "Ken desató una ráfaga torrencial de cientos de golpes ígneos en una fracción de segundo. Sus puños caían como una lluvia ininterrumpida de meteoros ardientes sobre el noble indefenso, cada golpe resonando como un trueno en el cielo despejado.",
        "¡TATATATATATATATA-BOOM!",
        "Para el golpe final, Ken giró sobre su propio eje en el aire y descargó una patada descendente con el talón envuelto en una llamarada colosal que tomó la forma exacta de un fénix en picado.",
        "¡¡¡KAAA-BOOOOOOOOOOOOOM!!!",
        "El cuerpo de Julian se estrelló contra el centro de la arena con la violencia de un meteorito, levantando una onda expansiva de viento ardiente y polvo que hizo retumbar las paredes de mármol del coliseo.",
        "* * *",
        "Una densa cortina de humo y calor sofocante cubrió la arena.",
        "Los veinte mil espectadores guardaban un silencio sepulcral, paralizados por la magnitud del poder que acababan de presenciar.",
        "Poco a poco, el humo comenzó a disiparse.",
        "En el centro de un cráter humeante de diez metros de diámetro, Julian de Montfort yacía profundamente inconsciente sobre la tierra calcinada, con su estoque rúnico partido en tres pedazos a su lado.",
        "Y de pie sobre el borde del cráter, con la respiración entrecortada pero la cabeza erguida hacia las alturas, estaba Ken.",
        "Su torso ensangrentado exhibía las marcas de la batalla, pero su puño derecho humeaba con chispas doradas que se apagaban lentamente. Con un gesto triunfal, levantó el puño cerrado hacia el cielo del coliseo y soltó un rugido de victoria visceral que desafió a todo el imperio.",
        "El árbitro tragó saliva, con las manos temblando de asombro, antes de bajar la bandera con fuerza:",
        "—¡G-GANADOR POR NOCAUT TOTAL... KEN, DE LA CLASE F-2!",
        "El coliseo entero tardó dos segundos en asimilarlo antes de explotar en el clamor más estruendoso jamás escuchado en la historia del torneo.",
        "En nuestro palco, la euforia fue absoluta: Burnice saltaba sobre los bancos derramando lágrimas de alegría; Lucy gritaba a viva voz con su gorra en alto; Lighter aplaudía con una sonrisa orgullosa de oreja a oreja; y Erza mantenía una sonrisa de profunda satisfacción marcial.",
        "Yo contemplé a Ken desde la barandilla con el corazón latiéndome con fuerza en el pecho.",
        "La Clase F-2 acababa de derribar al campeón de la clase más poderosa de Solaria en el primer asalto.",
        "El mensaje había sido entregado de forma brutal e irrefutable: ya nadie en este reino volvería a subestimar el poder de los marginados."
    ]
}

def compilar_capitulo_23():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 23              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 23
    txt_cap23 = os.path.join(capitulos_dir, "23_Capitulo_23.txt")
    with open(txt_cap23, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_23['header'].upper()}: {CAPITULO_23['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_23['paragraphs']))
    print(f"[OK] Texto plano Capítulo 23 guardado en:\n  {txt_cap23}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 23
    docx_cap23 = os.path.join(capitulos_dir, "23_Capitulo_23.docx")
    create_single_chapter_document(CAPITULO_23, docx_cap23, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 23 guardado en:\n  {docx_cap23}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 23)
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
    compilar_capitulo_23()
