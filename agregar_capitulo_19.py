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

CAPITULO_19 = {
    "header": "Capítulo 19",
    "title": "Calor Inconsciente",
    "paragraphs": [
        "El amanecer del segundo día en la Suite Real de Aguas Termales tiñó las cimas montañosas de un suave tono ambarino.",
        "Una niebla vaporosa y tibia flotaba sobre los estanques de piedra y los jardines de pinos centenarios, envolviendo el complejo en una serenidad reconfortante. Tras el abundante desayuno donde Renji y Rukia se lanzaban miradas fugaces cargadas de secretos compartidos, el resto de la Clase F-2 se dispersó para aprovechar al máximo las últimas veinticuatro horas de descanso antes de regresar al rigor marcial de Solaria.",
        "Sin embargo, si había alguien incapaz de concebir el 'descanso' como sinónimo de quedarse quieto, eran Ken y Burnice White.",
        "Para el mediodía, aquel torbellino de energía ya había encontrado su propio rincón en el patio de piedra anexo a los baños occidentales, donde los respiraderos termales expulsaban un calor seco y constante.",
        "* * *",
        "Yo me encontraba sentado sobre una banqueta de madera baja en el borde del pabellón techado, recuperándome con estiramientos suaves tras una sesión matutina de acondicionamiento físico. Kaoruko estaba sentada a mi lado, preparando una infusión tibia de hojas de loto silvestre en una tetera de barro; su larga y sedosa melena negra azabache estaba recogida en una coleta alta con una horquilla de madera pulida, dejando a la vista la curva delicada de su cuello sonrosado por el vapor ambiental.",
        "A escasos metros de nosotros, en una mesa rústica de piedra repleta de engranajes de bronce, tubos de ventilación y viales con esencias alquímicas de las termas, el 'dúo incinerador' estaba en pleno apogeo.",
        "—¡No, no, Ken-kun! —exclamaba Burnice agitando una llave inglesa plateada frente a su rostro—. Si el regulador de vapor mineral se abre antes de que la cámara alcance los trescientos grados, la chispa no hace '¡FWOOSH!', sino que hace '¡KABOOM!' en mi propia nariz. ¡Y costó demasiado peinar mi flequillo esta mañana como para chamuscarlo!",
        "—¡Pero si le damos más flujo de oxígeno térmico, la explosión primaria te daría el triple de impulso en los giros! —replicaba Ken con entusiasmo desbordante, agitando los brazos en el aire.",
        "Ken estaba sentado con las piernas cruzadas sobre las tablas pulidas del porche, con el torso cubierto únicamente por un yukata corto sin mangas que dejaba al descubierto sus hombros curtidos por el entrenamiento, con su inseparable bandana roja descansando floja sobre la clavícula. Y Burnice, con la más absoluta falta de pudor y de sentido del espacio personal, estaba sentada justo detrás de él, con las rodillas dobladas a ambos lados de los hombros de Ken, utilizando la espalda ancha y musculosa del muchacho como si fuera el respaldo de su propio sillón.",
        "Para cualquier noble estirado de la capital, aquella postura habría sido escandalosamente íntima: los muslos de Burnice rozaban los costados de Ken y su barbilla descansaba sobre la nuca del chico mientras ajustaba una tuerca con precisión de relojero. Pero para ellos dos, era simplemente la forma más natural y cómoda de trabajar.",
        "En un momento dado, Burnice necesitó alcanzar un frasco de reactivo volcánico que había quedado en la esquina opuesta de la mesa. En lugar de bajarse, se inclinó hacia adelante por completo, pasando un brazo por encima del hombro de Ken y pegando su pecho contra la espalda del muchacho. Ken ni pestañeó; con una mano firme sostuvo la cintura de la chica para evitar que se fuera de bruces contra las herramientas, mientras con la otra le acercaba el frasco sin desviar la mirada del plano.",
        "Justo en ese instante, Lucy dobló el sendero de gravilla cargando una pila de pergaminos tácticos.",
        "La capitana de los Hijos de Calydon se detuvo en seco. Sus ojos se abrieron como platos y su gorra militar casi se le resbala hacia atrás, revelando una vena latiéndole en la sien.",
        "—¡¿SE PUEDE SABER QUÉ CLASE DE INDECENCIA ESTÁN HACIENDO AQUÍ AFUERA?! —rugió Lucy, haciendo temblar los papeles entre sus brazos—. ¡Se supone que este es un balneario de descanso real! ¡Ken, estás sirviendo de mueble de descanso y tú, Burnice, pareces una enredadera trepada a un poste!",
        "Burnice parpadeó con inocencia absoluta, subiéndose las gafas rojas hasta la frente.",
        "—¡Pero Lucy! La espalda de Ken-kun es ancha, firme y está calentita como una estufa de piedra —explicó con una sonrisa radiante—. ¡Es el mejor soporte del mundo para calibrar boquillas!",
        "—¡Exacto, capitana enana! —secundó Ken, masticando un panecillo al vapor que había birlado del desayuno—. Además, me tapa el sol en la nuca y no me quemo mientras descansamos. Es un intercambio perfectamente justo.",
        "Lucy se llevó las manos al rostro, ahogando un grito de pura exasperación militar.",
        "—¡Son dos idiotas rematados! ¡Dos cabezas huecas sin una sola pizca de vergüenza! —masculló la rubia antes de darse media vuelta a paso furioso entre refunfuños sobre la decadencia moral de la juventud.",
        "Kaoruko se tapó la boca con ambas manos para disimular su risita tímida, mientras yo soltaba una bocanada de aire divertido.",
        "—Esos dos no tienen remedio —comenté en voz baja, tomando un sorbo de té.",
        "—Son adorables a su manera, Seier-san —respondió Kaoruko con una mirada dulce—. Se tienen una confianza tan pura que ni siquiera piensan en malicia.",
        "* * *",
        "Poco después, la calibración de las boquillas llegó a su fase de prueba práctica.",
        "—Muy bien, Ken-kun —dijo Burnice, colocándose uno de los guanteletes de ignición en el antebrazo derecho y poniéndose de pie de un brinco alegre—. Necesito que concentres una llama pequeña y constante en la palma de tu mano. Solo un hilito de fuego puro para encender la mezcla de combustión lenta.",
        "—Déjamelo a mí —sonrió Ken, frotándose los nudillos antes de extender la palma abierta al frente.",
        "Una llama dorada, limpia y densa, brotó del centro de su mano con un zumbido suave. Burnice acercó la boquilla de su propulsor, ajustando el regulador con un destornillador fino.",
        "Sin embargo, la alquimia experimental y los gases minerales del balneario suelen ser caprichosos.",
        "Una pequeña burbuja de gas comprimido se acumuló en la recámara trasera del dispositivo. Al entrar en contacto con la chispa de Ken, la mezcla detonó con un estallido seco mucho más violento de lo calculado.",
        "¡BANG!",
        "Una llamarada carmesí estalló en retroceso, disparando fragmentos de metal caliente directo hacia el rostro de Burnice.",
        "Ken no lo pensó. Su cuerpo de artista marcial reaccionó antes de que su cerebro pudiera procesar el peligro: cerró la distancia en una fracción de segundo, rodeó a Burnice con su torso y cubrió su cabeza con el brazo izquierdo, recibiendo la ráfaga de calor y la metralla en el antebrazo y en el dorso de la mano.",
        "El destello se apagó en una nube de humo perfumado a canela, queroseno y azufre mineral.",
        "—¡Uf...! —soltó Ken con una mueca de dolor, sacudiéndose el brazo humeante.",
        "Burnice abrió los ojos bajo el refugio del cuerpo de él. Al ver la piel enrojecida del muchacho y las quemaduras superficiales en sus nudillos, toda la alegría despreocupada de su rostro se evaporó al instante.",
        "—¡Ken! —gritó con un tono de angustia que rara vez se le escuchaba—. ¡Idiota, ¿por qué pusiste el brazo?!",
        "—Tranquila, no fue nada... mis brazos están curtidos con rocas —intentó bromear él con una risa forzada, pero una punzada de ardor le hizo apretar la mandíbula.",
        "Burnice no lo escuchó. Arrojó la herramienta al suelo, lo agarró firmemente de la muñeca y lo obligó a sentarse sobre el banco de piedra.",
        "—¡Quédate quieto! —ordenó con una autoridad insólita en ella, con los ojos grandes llenos de genuina preocupación.",
        "Rebuscó rápidamente en su bolsa de cuero hasta encontrar un frasco de ungüento medicinal de aloe y extracto de lirio helado que Kaoruko le había entregado esa misma mañana. Se arrodilló entre las piernas de Ken sobre las tablas, tomó su mano herida entre las suyas y comenzó a aplicar la pasta fría con la yema de sus dedos.",
        "El contacto fue instantáneo y desconcertante.",
        "Las manos de Burnice, que siempre estaban manipulando palancas, tuercas y detonadores, se movían con una delicadeza infinita, casi como si temiera romperlo. Sus dedos suaves y tibios acariciaban cada centímetro de su antebrazo, esparciendo el ungüento sobre los nudillos lastimados.",
        "Luego, inclinó su rostro hacia adelante y comenzó a soplar suavemente sobre las heridas.",
        "Fuuu... fuuu...",
        "Un aliento tibio y dulce, perfumado a canela, caramelo y frutas del balneario, acarició la piel desnuda de Ken.",
        "En ese momento, algo cambió en el aire.",
        "Ken se quedó completamente congelado. El dolor de la quemadura desapareció, eclipsado por una sensación completamente ajena y abrumadora que le recorrió la espina dorsal. Desde esa distancia tan corta, podía ver los mechones rubios desordenados de Burnice rozando sus rodillas, la curva suave de su cuello, la línea rosada de sus labios entreabiertos y el brillo intenso de sus ojos fijos en sus manos.",
        "De pronto, el corazón de Ken dio un vuelco salvaje dentro de su pecho.",
        "Tum-tum. Tum-tum. Tum-tum.",
        "El pulso del muchacho se disparó a una velocidad frenética, resonando en sus oídos como los tambores de una batalla.",
        "Burnice, que sostenía la muñeca de Ken para mantenerla firme, sintió el latido desbocado vibrando bajo sus propios dedos. Detuvo el soplido y levantó la mirada hacia los ojos de Ken, parpadeando con sorpresa genuina.",
        "—Ken-kun... —susurró ella, frunciendo el ceño con una ternura inocente—. Tu corazón está latiendo súper rápido... ¿te duele tanto? ¿está quemándote demasiado?",
        "Ken sintió que el calor subía como un incendio forestal desde su cuello hasta la punta de sus orejas, tiñéndole el rostro de un rojo carmesí furioso.",
        "—¡N-No! ¡Para nada! —balbuceó apartando la mirada con una torpeza mayúscula, rascándose la nuca con la mano sana—. Es... es solo la técnica del Dragón Ígneo... cuando el cuerpo detecta daño térmico, el flujo de maná muscular acelera la circulación de la sangre para enfriar los canales de energía... ya sabes, cosas básicas de artes marciales... ¡técnicas avanzadas de respiración!",
        "Burnice ladeó la cabeza, contemplándolo durante dos segundos con una mezcla de fascinación y duda, antes de que su habitual sonrisa deslumbrante volviera a iluminar su rostro.",
        "—¡Waaah, qué increíble es el entrenamiento marcial! —exclamó dando un pequeño aplauso con las manos manchadas de crema—. ¡Pensé que te estabas muriendo del dolor! Pero qué alivio... me asusté de verdad, tonto.",
        "Sin previo aviso y con la espontaneidad que la caracterizaba, Burnice se estiró hacia arriba, rodeó el cuello de Ken con un brazo y le plantó un beso sonoro y cálido en la mejilla antes de soltarlo.",
        "—¡Gracias por protegerme, héroe! —le dijo con un guiño alegre, antes de ponerse a recoger los tornillos del suelo como si nada hubiera ocurrido.",
        "Ken se quedó sentado como una estatua de piedra en el banco, con la mano izquierda apoyada sobre la mejilla que ella acababa de besar. Su boca estaba entreabierta y sus ojos fijos en la nada, completamente incapaz de procesar por qué su estómago se sentía como si hubiera tragado un enjambre de mariposas en llamas.",
        "* * *",
        "Al caer la tarde, el cielo sobre las aguas termales se vistió de un manto violeta y dorado.",
        "Agotados por horas de reparaciones, baños de vapor y caminatas por los senderos rocosos, el cansancio acumulado finalmente les pasó factura.",
        "Kaoruko y yo regresábamos de pasear por el jardín de cerezos cuando los vimos en el banco largo de roble, ubicado bajo la sombra fresca de un gran sauce llorón junto al riachuelo mineral.",
        "Burnice, rendida por el esfuerzo del día, se había quedado profundamente dormida. Su cuerpo menudo estaba recostado a lo largo del banco, con las piernas recogidas y su cabeza descansando plácidamente sobre el regazo de Ken.",
        "Y Ken, que habitualmente era incapaz de permanecer inmóvil ni por treinta segundos seguidos, estaba sentado con la espalda apoyada contra el tronco, completamente quieto.",
        "Con una lentitud casi devocional, el chico le había quitado las gafas rojas para que no le lastimaran el puente de la nariz, dejándolas con cuidado sobre la madera. Con su mano vendada, le apartaba suavemente los mechones de cabello dorado que caían sobre la frente de la muchacha, acariciándole la coronilla con una ternura instintiva y silenciosa.",
        "Burnice emitió un pequeño murmullo en sueños, sonrió plácidamente y se acurrucó todavía más contra él, rodeándole el muslo con ambos brazos y hundiendo el rostro en el costado del muchacho.",
        "Ken soltó un suspiro profundo, dejó escapar una sonrisa boba y cálida que rara vez mostraba en público, y apoyó su propia cabeza hacia atrás, cerrando los ojos mientras rodeaba el hombro de la chica con su brazo de forma protectora.",
        "Ambos respiraban al unísono, fundidos en un sueño profundo y tranquilo bajo la brisa tibia.",
        "—Míralos... —susurró Kaoruko a mi lado, juntando las manos contra su pecho con los ojos iluminados por una emoción conmovedora—. Parecen dos cachorros cansados después de jugar todo el día. Se ven tan en paz juntos, Seier-san.",
        "—Son un par de despistados —respondí en un murmullo con una media sonrisa—. Ninguno de los dos sabe lo que está haciendo, pero no creo que nadie en este mundo pueda separarlos ahora.",
        "En ese momento, Caesar King y Piper pasaron caminando por el sendero empedrado. La líder de los Hijos de Calydon observó a la pareja dormida de brazos cruzados, con una sonrisa orgullosa y satisfecha dibujándose en su rostro fiero.",
        "—Ese chico Ken es ruidoso y terco, pero cuida de Burnice mejor que cualquiera que haya conocido —dijo Caesar en voz baja, acomodándose la chaqueta sobre los hombros—. En Calydon tenemos una regla simple: cuando encuentras a alguien que no le teme a tu propio fuego y está dispuesto a arder a tu lado... no lo dejas ir jamás.",
        "—Ojalá no despierten pronto... la paz es agradable... —bostezó Piper acurrucándose en su abrigo de lana.",
        "* * *",
        "Cuando las linternas de piedra del balneario se encendieron al anochecer anunciando el banquete de la segunda noche, ambos despertaron estirándose como gatos perezosos.",
        "No hubo gritos de vergüenza, ni mejillas encendidas de culpabilidad, ni dramas aristocráticos. Ken le entregó sus gafas a Burnice con una sonrisa tranquila y ella se las puso de un salto, riéndose por el peinado desastroso que ambos llevaban tras la siesta.",
        "Cuando Lucy se cruzó con ellos en el pasillo exterior y les preguntó con sorna si ya habían terminado su 'cita nupcial de retiro', los dos respondieron a coro, al unísono y con la más absoluta convicción:",
        "—¡¿Cita?! ¡Pero si solo somos el dúo incinerador tomando una siesta de compas!",
        "Sin embargo, cuando caminaron juntos hacia el pabellón principal y sus hombros chocaron por casualidad en el camino, sus manos se rozaron por un instante.",
        "Una pequeña chispa eléctrica, invisible para los ojos de los demás pero ardiente en lo profundo de sus pechos, les hizo acelerar el paso con una sonrisa secreta que ninguno de los dos supo cómo explicar.",
        "Eran lentos, eran torpes y no tenían la menor idea de hacia dónde los llevaba el camino.",
        "Pero la llama ya estaba encendida en las aguas termales. Y en una clase como la nuestra, sabíamos bien que una vez que el fuego comienza a arder bajo las cenizas, no hay fuerza en este mundo capaz de apagarlo."
    ]
}

def compilar_capitulo_19():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 19              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 19
    txt_cap19 = os.path.join(capitulos_dir, "19_Capitulo_19.txt")
    with open(txt_cap19, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_19['header'].upper()}: {CAPITULO_19['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_19['paragraphs']))
    print(f"[OK] Texto plano Capítulo 19 guardado en:\n  {txt_cap19}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 19
    docx_cap19 = os.path.join(capitulos_dir, "19_Capitulo_19.docx")
    create_single_chapter_document(CAPITULO_19, docx_cap19, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 19 guardado en:\n  {docx_cap19}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 19)
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
    compilar_capitulo_19()
