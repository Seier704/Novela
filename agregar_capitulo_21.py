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

CAPITULO_21 = {
    "header": "Capítulo 21",
    "title": "El Ritmo de la Guardia",
    "paragraphs": [
        "El retorno a la Academia Real de Magia de Solaria tras los dos días en las aguas termales se sintió como el despertar de un sueño breve pero reparador.",
        "Los carruajes nos dejaron frente a las imponentes puertas de hierro negro bajo el sol templado del mediodía. Aunque las exigencias del curso y la mirada altiva de los nobles de las Clases A seguían esperando al doblar de cada esquina, la atmósfera dentro de la Clase F-2 había cambiado de forma irreversible. Ya no éramos un puñado de marginados desconfiados conviviendo por obligación; éramos una unidad forjada en el fuego de los calabozos y cohesionada por confidencias nocturnas.",
        "Renji y Rukia se movían con una sintonía tácita, compartiendo tareas sin discutir; Ken y Burnice andaban siempre entre risas y explosiones menores; y Subaru Natsuki parecía haber recuperado la energía para seguir soportando los rigurosos entrenamientos de Yami.",
        "Pero el cambio más sutil, y a la vez más significativo, ocurrió en Subaru Hoshina.",
        "* * *",
        "La tarde del primer día de regreso encontró el campo de entrenamiento oeste bañado por una luz dorada y polvorienta.",
        "El instructor de combate cuerpo a cuerpo nos había ordenado realizar prácticas libres de juego de pies y esquiva de intercepción. Mientras Natsuki intentaba esquivar los golpes de madera de Yami rodando por el fango, y Erza supervisaba la postura de espada de los demás con severidad militar, yo me encontraba descansando tras una intensa ronda de golpes al saco de arena.",
        "A pocos metros de mí, Hoshina practicaba sus desplazamientos rápidos sobre la arena circular.",
        "Su estilo era ágil, furtivo y letal, propio de un combatiente que prioriza las dagas y el sigilo: pasos cortos, cambios de dirección repentinos y saltos hacia atrás con el centro de gravedad bajo. Sin embargo, había una tensión evidente en sus movimientos. Saltaba con los talones suspendidos en el aire, forzando los gemelos y respirando con una rigidez innecesaria, como si cada paso fuera un intento desesperado de evitar una trampa oculta.",
        "Sentado sobre un tronco seco al borde del cuadrilátero, con las piernas abiertas y vendándose las manos con cinta de algodón crudo, Lighter la observaba en silencio.",
        "Sus gafas de sol estaban enganchadas en el cuello de su camiseta gris de tirantes, dejando a la vista sus ojos serenos de color avellana. Tras verla tropezar ligeramente al apoyar mal el pie derecho tras una finta brusca, el boxeador soltó un suspiro tranquilo, dejó el rollo de venda a un lado y se puso de pie.",
        "Caminó hacia la arena arrastrando los pies con esa calma perezosa y segura que lo caracterizaba.",
        "—Tus tobillos van a terminar pidiendo auxilio antes de que caiga el sol, Hoshina —dijo Lighter, deteniéndose a unos dos metros de ella sin invadir su espacio.",
        "Subaru se frenó en seco, con una daga de madera corta en la mano derecha. Por puro reflejo condicionado, sus ojos verdes se afilaron y dio medio paso atrás, pero al reconocer la voz áspera y tranquila de Lighter, la tensión en sus hombros cedió visiblemente. La memoria de la noche anterior en el mirador de bambú seguía fresca en su mente.",
        "—Mi velocidad es suficiente —respondió ella, intentando mantener un tono firme, aunque ya no había veneno ni desprecio en sus palabras.",
        "—Nadie duda de tu velocidad, señorita hielo —replicó Lighter con una media sonrisa—. Pero estás saltando como si pisaras brasas ardientes. En el boxeo, si no plantas la base del pie para pivotar la cadera, el menor impacto de rebote te descoloca y te deja a merced de un contragolpe. Mira mis pies un segundo.",
        "Lighter adelantó el pie izquierdo, flexionó las rodillas con soltura y realizó un giro milimétrico sobre la punta del pie adelantado, rotando el torso noventa grados en un abrir y cerrar de ojos, con un balance tan perfecto que parecía flotar sobre el polvo.",
        "—No saltes hacia atrás —explicó él, volviendo a su posición neutra—. Desliza el pie trasero, gira el talón y usa el impulso del suelo para cambiar el ángulo. La arena absorbe tu fuerza si brincas; si te deslizas, la arena trabaja para ti.",
        "Hoshina contempló el movimiento con atención analítica. Intentó imitarlo: adelantó el pie, intentó pivotar la cadera, pero al girar sobre la arena suelta, su bota resbaló en una pequeña raíz enterrada y su cuerpo se inclinó peligrosamente hacia un costado.",
        "Lighter reaccionó con la velocidad refleja de un campeón de peso medio.",
        "Dio un paso al frente y extendió su brazo firme, atrapándola con seguridad por el codo para impedir la caída. Su agarre fue sólido como el hierro, sosteniéndola durante una fracción de segundo antes de ayudarla a asentar ambos pies en la tierra.",
        "En cuanto Hoshina recuperó el equilibrio, Lighter retiró la mano de inmediato, sin sobrepasarse ni un milímetro.",
        "—Te tengo —dijo él con tranquilidad—. Buen intento. Solo mantén la barbilla recogida para no perder el horizonte.",
        "Subaru se quedó congelada.",
        "Durante ese instante en que la mano grande y callosa de Lighter sostuvo su brazo, no sintió náuseas, ni asco, ni esa alarma estridente que solía dispararse en su pecho cada vez que un hombre se le acercaba a menos de un metro. Al contrario: sintió una firmeza cálida, un soporte protector y seguro que la dejó desconcertada.",
        "Sus mejillas se tiñeron de un rubor tenue y rosado. Tragó saliva disimuladamente y miró hacia el suelo antes de levantar los ojos hacia él.",
        "—G-Gracias... —murmuró en voz baja, con un atisbo de torpeza adorable que rara vez dejaba entrever.",
        "—No hay de qué. Inténtalo otra vez, pero con la mitad de fuerza y el doble de ritmo. Yo marco el compás.",
        "* * *",
        "A un lado del campo, Kaoruko y yo observábamos la escena.",
        "Kaoruko sostenía dos cantimploras de agua fresca contra su pecho, con sus grandes ojos oscuros brillando de pura fascinación y una sonrisa radiante dibujándose en sus labios.",
        "—Seier-kun, ¿estás viendo eso? —susurró con entusiasmo contenido, tirando levemente de la manga de mi chaqueta—. ¡Subaru-chan está dejando que Lighter-san la corrija! ¡Y hasta le dio las gracias!",
        "—Es un gran avance —asentí con una sonrisa tranquila, cruzándome de brazos—. Hoshina es orgullosa, pero sabe reconocer la experiencia real. Y Lighter es de los pocos tipos que no intenta demostrar nada frente a una mujer. Solo es él mismo.",
        "—Me alegra tanto... —suspiró Kaoruko, juntando las manos—. Siempre temía que Subaru-chan se sintiera tan sola por intentar protegerme todo el tiempo. Verla hablar así con alguien me quita un peso enorme del corazón.",
        "En la arena, Hoshina repitió el giro tres veces. En la tercera ocasión, el movimiento fluyó con una elegancia impecable: su cuerpo giró en un suspiro, saliendo de la línea de ataque imaginaria y quedando posicionada perfectamente para un corte lateral.",
        "Lighter soltó un silbido bajo de aprobación y asintió con la cabeza.",
        "—Limpio. Peligroso y limpio —felicitó con sinceridad—. Si algún noble engreído intenta arrinconarte en un pasillo, haz ese giro y ponle la daga en las costillas. Te aseguro que no volverá a molestarte.",
        "Subaru no pudo contener una pequeña risita cristalina.",
        "—Tomaré nota de tu consejo, boxeador.",
        "* * *",
        "Al caer la noche, una brisa fresca comenzó a mecer las copas de los robles alrededor de los dormitorios de estudiantes.",
        "Hoshina se encontraba sentada en el tramo superior de las escaleras de piedra del patio trasero, un rincón apartado y silencioso iluminado por un farol de aceite. Tenía las piernas cruzadas y sostenía un cuaderno de notas tácticas sobre el regazo, repasando los apuntes de la clase de estrategia militar que Erza había impartido esa mañana.",
        "El crujido de pisadas sobre la gravilla del sendero la hizo levantar la vista.",
        "Por el camino venía Lighter, con las manos metidas en los bolsillos de su cazadora de cuero y su paso tranquilo. Venía de regreso del distrito comercial de Solaria, donde solía ir a comprar repuestos para los vehículos de Caesar o suministros para su grupo.",
        "Al verla sentada en la soledad de la escalinata, Lighter frenó el paso al pie de los escalones.",
        "—¿Sigues despierta a estas horas, señorita estratega? —preguntó, alzando una ceja.",
        "—El examen de fin de mes de la profesora Scarlet no va a aprobarse solo —respondió ella, cerrando el cuaderno sobre sus rodillas—. ¿Y tú? Pensé que estarías en la taberna con Caesar y Ken.",
        "—Demasiado ruido para un martes por la noche —dijo él, subiendo un par de peldaños con lentitud—. Además, tenía que hacer un encargo en el mercado central.",
        "Lighter sacó la mano derecha de su bolsillo y extendió un paquete pequeño envuelto en papel encerado y atado con un cordel de cáñamo crudo.",
        "Lo arrojó suavemente hacia ella con un tiro parabólico impecable. Hoshina lo atrapó en el aire con reflejos felinos, mirando el paquete con curiosidad.",
        "—¿Qué es esto? —inquirió, ladeando la cabeza.",
        "—Pastelillos de hojaldre rellenos de mermelada de higo negro y miel silvestre —respondió Lighter con naturalidad, rascándose la nuca—. Pasé por la panadería del barrio bajo donde Silva suele comprar y el olor me recordó que a cierta chica de hielo le brillaban los ojos cada vez que Waguri mencionaba los dulces.",
        "Subaru se quedó sin palabras.",
        "Desató el cordel con dedos temblorosos y abrió el papel. El aroma dulce, mantecoso y tibio de los pasteles recién horneados inundó el aire fresco de la noche. Eran dos hojaldres dorados, crujientes y espolvoreados con azúcar fina.",
        "Un calor profundo y enternecedor le subió desde el pecho hasta la garganta. Nadie, a excepción de Kaoruko en toda su vida, se había tomado la molestia de recordar un detalle tan minúsculo e insignificante sobre ella. Y que fuera precisamente este chico rudo, cubierto de cicatrices y con fama de pendenciero fronterizo, quien lo hiciera, derribó los últimos ladrillos de desconfianza que quedaban en su interior.",
        "Levantó la mirada hacia él. Lighter seguía de pie un par de escalones más abajo, esperando su reacción con una tranquilidad despreocupada.",
        "—Si te quedas parado ahí abajo pareces una torre de asedio —dijo Hoshina en un murmullo suave, apartándose un poco hacia un lado y palmeando el escalón de piedra a su lado—. Siéntate. No pienso comerme los dos sola.",
        "Lighter sonrió de lado, divertido.",
        "—Trato hecho.",
        "Subió los escalones restantes y se sentó a su lado sobre la piedra fresca, dejando una distancia respetuosa pero compartiendo el mismo espacio sin incomodidad.",
        "Hoshina partió uno de los hojaldres por la mitad y le entregó una parte, quedándose con la otra. Al darle el primer bocado, sus ojos se cerraron con un suspiro de deleite puro, mientras el sabor dulce de los higos y la miel se derretía en su boca.",
        "—Está delicioso... —susurró, con una expresión tan genuinamente infantil y feliz que Lighter no pudo evitar mirarla con una ternura cómplice.",
        "—Te lo dije. Los viejos de esa panadería saben lo que hacen.",
        "* * *",
        "Permanecieron allí durante más de una hora, comiendo los dulces y conversando en voz baja bajo la luz plateada de la luna.",
        "Lighter le contó historias de su infancia en las Tierras Yermas de Calydon: las carreras de monstruos mecánicos por el polvo, las tormentas de arena roja y cómo Caesar lo había sacado de un callejón cuando no era más que un chiquillo peleonero que no sabía esquivar un directo. Subaru lo escuchaba con una fascinación sincera, descubriendo detrás de esa fachada ruda a un hombre leal, con un sentido del honor inquebrantable y un corazón noble que cuidaba ferozmente de los suyos.",
        "A su vez, Hoshina se encontró compartiendo recuerdos de su provincia natal: los campos de cerezos fríos, las tardes interminables cosiendo kimonos junto a Kaoruko y el peso que siempre sintió sobre sus hombros por pertenecer a una rama menor de una familia guerrera caída en desgracia.",
        "En un momento de la conversación, Lighter soltó un chiste irónico sobre la obsesión de Lucy por los sombreros militares, y Subaru dejó escapar una carcajada limpia, libre y espontánea.",
        "Ella misma se detuvo un instante, sorprendida por el sonido de su propia risa.",
        "Miró a Lighter a los ojos. El muchacho la contemplaba con una sonrisa serena y cálida, sin una sola pizca de burla, simplemente disfrutando de su compañía.",
        "En ese segundo silencioso, Subaru Hoshina comprendió que algo fundamental había cambiado.",
        "Ya no veía a Lighter como a un peligro potencial, ni como a un simple compañero de clase impuesto por las circunstancias. Lo veía como un amigo verdadero. Como un refugio inesperado donde su orgullo no necesitaba luchar, donde su frialdad podía derretirse sin miedo y donde sus debilidades no eran motivo de vergüenza.",
        "Cuando el reloj de la torre de la academia dio la medianoche y llegó el momento de volver a sus habitaciones, Hoshina se puso de pie, guardó su cuaderno y lo miró fijamente.",
        "—Gracias por los dulces, Lighter —dijo con una voz cargada de una dulzura inédita—. Y por enseñarme ese paso en la arena.",
        "Lighter se levantó, se sacudió los pantalones y le dedicó un guiño amistoso antes de meterse las manos en los bolsillos.",
        "—Nos vemos mañana en el entrenamiento, Hoshina. Trata de no matar a nadie antes del almuerzo.",
        "—No prometo nada —respondió ella con una media sonrisa pícara.",
        "Caminó hacia el ala femenina con el corazón latiendo a un ritmo suave, templado y alegre.",
        "El hielo que la aislaba del mundo se estaba desvaneciendo paso a paso, reemplazado por la promesa silenciosa de una amistad que, sin que ninguno de los dos lo dijera en voz alta, comenzaba a rozar las orillas de algo mucho más profundo."
    ]
}

def compilar_capitulo_21():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 21              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 21
    txt_cap21 = os.path.join(capitulos_dir, "21_Capitulo_21.txt")
    with open(txt_cap21, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_21['header'].upper()}: {CAPITULO_21['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_21['paragraphs']))
    print(f"[OK] Texto plano Capítulo 21 guardado en:\n  {txt_cap21}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 21
    docx_cap21 = os.path.join(capitulos_dir, "21_Capitulo_21.docx")
    create_single_chapter_document(CAPITULO_21, docx_cap21, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 21 guardado en:\n  {docx_cap21}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 21)
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
    compilar_capitulo_21()
