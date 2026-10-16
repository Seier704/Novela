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

CAPITULO_32 = {
    "header": "Capítulo 32",
    "title": "Nueva Sangre, Nuevos Caminos",
    "paragraphs": [
        "Con los cinco días de fiesta y descanso concluidos, la primera etapa de nuestra vida en la Academia Real de Magia de Solaria quedó formalmente cerrada.",
        "Los ecos del torneo de primer año se apagaron en los pasillos de mármol, reemplazados por una tensión mucho más seria y profesional. Ya no éramos simples cadetes recién llegados aprendiendo a canalizar maná en aulas acolchadas ni compitiendo bajo la mirada vigilante de árbitros con banderas de rendición. El primer semestre había sido el tamiz para separar la paja del grano.",
        "Ahora comenzaba la verdadera prueba de fuego: la Segunda Parte de nuestra instrucción imperial.",
        "El inicio de las incursiones a calabozos de nivel medio.",
        "* * *",
        "El Gran Auditorio de la academia congregó a todos los estudiantes de primer año bajo un silencio riguroso.",
        "En el estrado principal, el Mariscal de Campo Vaneck —un veterano de mil batallas con la mitad del rostro marcada por quemaduras de ácido bestial— desenrolló un pergamino de piel de wyvern con el sello de guerra del imperio.",
        "—¡Atención, cadetes! —bramó el mariscal, haciendo retumbar las columnas de granito—. El juego de niños en el coliseo terminó. Los calabozos de rango medio no son arenas controladas con piso de arena limpia. Son abismos subterráneos no explorados en su totalidad, infectados de miasma corrosivo, bestias con inteligencia táctica y trampas antiguas que no perdonan errores.",
        "El oficial paseó su mirada de halcón sobre nosotros antes de lanzar la bomba que descolocó a toda la sala:",
        "—Para esta fase, las clases quedan disueltas operativamente. En la guerra real, los regimientos se mezclan según las necesidades del frente. No pelearán con sus amigos de siempre. El consejo ha estructurado escuadrones mixtos de cinco integrantes combinando cadetes de todas las clases.",
        "Un murmullo de conmoción y quejas ahogadas recorrió el auditorio.",
        "—¡Silencio! —rugió Vaneck—. Las listas oficiales han sido publicadas en la Plaza del Reloj. Tienen cuarenta y ocho horas para conocer a sus nuevos compañeros, coordinar sinergias y preparar suministros. Si no aprenden a confiar en quien tienen al lado, el calabozo se encargará de digerir sus cadáveres.",
        "* * *",
        "En la Plaza Central, la multitud de estudiantes se agolpaba frente a los tableros mágicos flotantes donde los nombres de los nuevos escuadrones brillaban con tinta rúnica.",
        "Por orden del consejo, para mantener un mínimo de estabilidad psicológica, cada cadete de la F-2 había sido emparejado de forma aleatoria con un solo compañero de su clase de origen, mientras que las otras tres plazas correspondían a estudiantes de las Clases A, B y C.",
        "Las reacciones alrededor de las listas fueron un espectáculo aparte:",
        "—¡JA! ¡MIRA ESO, KEN-KUN! —gritó Burnice, jalando a Ken del brazo mientras señalaba el Escuadrón 3—. ¡El destino no puede separarnos! ¡Nos tocó en el mismo grupo con un caballero rúnico de la B-1 y dos magos de asedio de la C!",
        "—¡Eso significa que la cocina y las explosiones siguen bajo control! —festejó Ken con una carcajada sonora, atándose su bandana roja.",
        "A pocos pasos de ellos, frente al Escuadrón 5, Lighter sonreía con una calma relajada mientras Subaru Hoshina observaba la lista con un alivio evidente en sus ojos verdes. Les había tocado compartir filas junto a dos nobles de la Clase A-1 y una exploradora de la B-1.",
        "—Parece que tendré que vigilar que esos dos niños ricos no te miren raro, señorita hielo —dijo Lighter con una sonrisa de medio lado, acomodándose las gafas oscuras.",
        "—Con que vigiles tu propia espalda en el fango será suficiente, boxeador —respondió Hoshina con una sonrisa dulce y cómplice que dejó desarmado al muchacho.",
        "En otra esquina, la suerte había emparejado a **Erza Scarlet y Akame** en el Escuadrón 2, formando una vanguardia de espadas que ya hacía temblar a los otros tres integrantes que les habían asignado; mientras que **Renji y Rukia** celebraban en silencio haber quedado juntos en el Escuadrón 4.",
        "El único drama absoluto lo protagonizó Subaru Natsuki frente al Escuadrón 6:",
        "—¡¿POR QUÉ EL DESTINO ES TAN CRUEL CONMIGO?! —chilló Natsuki cayendo de rodillas con las manos en la cabeza—. ¡¿TODA LA ACADEMIA PARA ELEGIR Y ME VUELVE A TOCAR CON EL CAPITÁN HUMO?!",
        "Detrás de él, Yami Sukehiro soltó una bocanada de humo espeso, apoyó su bota pesada sobre la espalda de Natsuki y sonrió con malicia:",
        "—Agradece a los dioses, mocoso. Eso significa que tendrás que cargar mi tabaco en el calabozo si quieres salir con vida.",
        "* * *",
        "Mientras los demás terminaban de asimilar sus destinos, Kaoruko y yo nos acercamos al panel del **Escuadrón 7**.",
        "La chica azabache se aferró a la manga de mi chaqueta, leyendo los nombres proyectados en luz ámbar:",
        "1. **Seier Silva** (Clase F-2)",
        "2. **Kaoruko Waguri** (Clase F-2)",
        "3. **Elysea de Vane** (Clase A-2)",
        "4. **Dante Vance** (Clase B-1)",
        "5. **Freya Skuld** (Clase C-1)",
        "—Seier-kun... estamos juntos —susurró Kaoruko con una sonrisa llena de alivio, apretando mi brazo contra su pecho—. Pero... ¿Lady Elysea también está en nuestro equipo?",
        "Antes de que pudiera responder, un aroma suave a escarcha y lirios invernales anunció la llegada de la tercera integrante.",
        "Elysea de Vane se abrió paso entre la multitud con su habitual elegancia aristocrática, vestida con su túnica azul marino y su estoque al cinto. Sus ojos celestes se posaron sobre la lista y luego se desviaron hacia nosotros, dejando ver una sonrisa sutil y satisfecha.",
        "—Parece que el destino escuchó nuestro brindis de la otra noche en la terraza, Silva —dijo la elfa con voz melodiosa, inclinando la cabeza con respeto—. Será un verdadero placer luchar en el mismo bando que el hombre que quebró mi *Cero Absoluto*.",
        "Elysea miró entonces a Kaoruko. Lejos de la altivez que muchos nobles solían mostrar, la elfa le dedicó una reverencia impecable.",
        "—Waguri-san. He visto tu magia de restauración durante el torneo. Tu afinidad con el agua pura será vital para purificar el miasma que nos espera.",
        "Kaoruko parpadeó sorprendida y se sonrojó levemente, devolviendo la reverencia con timidez.",
        "—M-Muchas gracias, Lady Elysea... haré mi mayor esfuerzo para mantener a todos a salvo.",
        "* * *",
        "—Vaya, vaya... parece que me tocó la lotería imperial.",
        "Una voz perezosa y ligeramente arrastrada se sumó a la conversación desde nuestro flanco.",
        "Un joven de cabello castaño desordenado y ojos ámbar se acercaba limpiando la lente de un artefacto óptico con un pañuelo. Llevaba una chaqueta de explorador de cuero oscuro repleta de bolsillos, cartucheras y un mosquete rúnico corto cruzado a la espalda.",
        "—Dante Vance, de la Clase B-1 —se presentó el muchacho con una sonrisa socarrona, guardando el pañuelo—. Explorador, trampero y tirador de apoyo. Francamente, cuando vi que me asignaban con el demoledor que derribó la muralla de Eisenhardt y con la mejor espadachina de hielo del reino, casi lloro de felicidad. Eso significa que no tendré que correr tanto en los túneles oscuros.",
        "—No te confíes, tirador —le advertí con una media sonrisa—. En mis grupos nadie se queda atrás esperando que otros hagan el trabajo sucio.",
        "—Entendido, capitán —bromeó Dante guiñando un ojo.",
        "En ese momento, una pisada pesada que hizo crujir las baldosas de piedra anunció a la última integrante.",
        "—¡JAJAJAJA! ¡QUÉ BUEN EQUIPO NOS HA TOCADO!",
        "Una carcajada atronadora retumbó en la plaza. Por el sendero avanzó una chica de complexión imponente, casi de mi misma estatura, con los hombros anchos y una musculatura definida propia de los pueblos nórdicos. Tenía largas trenzas de cabello rubio ceniza, ojos azul glaciar y portaba una lanza pesada con punta dentada y un escudo redondo de roble con refuerzos de hierro rúnico.",
        "La recién llegada se detuvo frente a mí, me miró de arriba abajo con una admiración salvaje y me propinó un manotazo tan amistoso en el omóplato que habría hecho tambalear a cualquier hombre ordinario.",
        "—¡Freya Skuld, de las tribus de la Costa Helada, cadete de la Clase C-1! —bramó la guerrera con una sonrisa enorme que mostraba sus dientes blancos—. ¡Por los dioses del trueno, Silva! ¡Tienes la espalda ancha como un oso de las cavernas! ¡Sé reconocer a un verdadero peleador de choque cuando lo veo enfrente!",
        "Freya miró a Elysea, a Kaoruko y a Dante, asintiendo con vigor.",
        "—Una elfa de hielo elegante, una sanadora de corazón puro, un explorador con olor a pólvora y un titán de puños de hierro. ¡Este escuadrón tiene la bendición del norte!",
        "Dante soltó un suspiro divertido rascándose la nuca, Elysea arqueó una ceja con curiosidad aristocrática y Kaoruko no pudo evitar soltar una risita tímida ante la energía desbordante de la norteña.",
        "* * *",
        "Con los cinco miembros reunidos, nos retiramos a una sala táctica para revisar la orden de misión sellada por el mariscal.",
        "Elysea desplegó el mapa cartográfico sobre la mesa de madera.",
        "—Nuestro objetivo asignado es el **Calabozo de las Cavernas del Miasma Carmesí** —explicó la elfa, señalando una cordillera volcánica en los límites fronterizos del sur—. Rango de peligro: Medio-Alto. Es una red de cavernas subterráneas saturada por un gas tóxico que corroe el maná superficial y engendra bestias subterráneas con caparazones blindados.",
        "—Un terreno peligroso —analizó Dante, señalando los puntos estrechos—. El miasma limitará la visión y el alcance de mis disparos rúnicos a menos de veinte metros.",
        "—Yo me encargaré de mantener una barrera de viento frío y escarcha que condense el gas tóxico alrededor de nuestra formación —propuso Elysea con precisión militar.",
        "—Y yo prepararé infusiones de lirio purificador y mantendré esferas de agua bendita para limpiar nuestras vías respiratorias —añadió Kaoruko con determinación.",
        "—Freya y yo aseguraremos la vanguardia —dictaminé, apoyando los puños cerrados sobre el mapa—. Su lanza y mi cuerpo mantendrán a raya a cualquier criatura que intente romper el perímetro mientras Dante y Elysea barren desde la media distancia.",
        "Freya chocó su lanza contra el escudo con un sonido metálico enérgico.",
        "—¡Me gusta cómo piensas, Silva! ¡Que vengan las bestias!",
        "Miré a los cuatro integrantes de mi nuevo escuadrón: la nobleza calculadora de Elysea, la ternura inquebrantable de Kaoruko, el pragmatismo astuto de Dante y la ferocidad salvaje de Freya.",
        "Era una formación completamente nueva, nacida de la mezcla de sangres y disciplinas de toda la academia.",
        "El torneo de primer año había demostrado que éramos fuertes.",
        "Pero en las profundidades oscuras de las Cavernas Carmesíes, estábamos a punto de demostrar que éramos imparables."
    ]
}

def compilar_capitulo_32():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 32              ")
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
        CAPITULO_30, CAPITULO_31, CAPITULO_32
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 32
    txt_cap32 = os.path.join(capitulos_dir, "32_Capitulo_32.txt")
    with open(txt_cap32, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_32['header'].upper()}: {CAPITULO_32['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_32['paragraphs']))
    print(f"[OK] Texto plano Capítulo 32 guardado en:\n  {txt_cap32}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 32
    docx_cap32 = os.path.join(capitulos_dir, "32_Capitulo_32.docx")
    create_single_chapter_document(CAPITULO_32, docx_cap32, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 32 guardado en:\n  {docx_cap32}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 32)
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
    compilar_capitulo_32()
