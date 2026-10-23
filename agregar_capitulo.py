# -*- coding: utf-8 -*-
"""
===============================================================================
GESTOR UNIFICADO DE CAPÍTULOS - "EL CAMINO DEL REENCARNADO"
===============================================================================
Este archivo es la plantilla GENERAL y ÚNICA para agregar o sobrescribir capítulos.
Para agregar un nuevo capítulo, simplemente edite el diccionario `CAPITULO` con
el número, título y párrafos correspondientes y ejecute este script:
    python agregar_capitulo.py

El script automáticamente:
1. Genera el documento Word individual: capitulos/XX_Capitulo_XX.docx
2. Recompila el libro maestro completo: compilado/El_Camino_del_Reencarnado.docx
===============================================================================
"""

import os
import sys
import glob
import re
import docx

# Definición de rutas base
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"c:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

for d in [CURRENT_DIR, NOVELA_DIR, PROYECTO_DIR]:
    if os.path.exists(d) and d not in sys.path:
        sys.path.insert(0, d)

from generar_libro import create_novel_document, create_single_chapter_document

# ===============================================================================
# DEFINICIÓN DEL CAPÍTULO A AGREGAR / ACTUALIZAR
# ===============================================================================
CAPITULO = {
    "num": 86,
    "header": "Capítulo 86",
    "title": "Las Galerías de Ónix: El Laberinto de Miasma y la Primera Cámara Nodriza",
    "paragraphs": [
        "A nuestras espaldas, la superficie de la Garganta del Cuervo bramaba como el vientre de un volcán en erupción. Las seis baterías de cañones de vapor del General Alden mantenían un fuego graneado ensordecedor: el silbido de los proyectiles de fósforo, el estampido sordo de la pólvora rúnica y el estrépito de la metralla sacudían la corteza terrestre, atrayendo hacia las empalizadas del Campamento de los Cien Leones a millares de bestias y tropas de asalto Skaven. El ejército cumplía con fiereza su papel de señuelo masivo.",
        "Pero al cruzar el umbral de la fisura primaria, el fragor de la guerra exterior comenzó a desvanecerse, sofocado por el peso milenario de la roca.",
        "Nos adentramos en las entrañas de las **Ruinas Negras del Sur**. La luz de las estrellas y los fogonazos de campaña se extinguieron por completo, sustituidos por una penumbra abismal y opresiva, rasgada únicamente por los destellos violetas y enfermizos del miasma que manaba de las grietas. Las galerías no eran meras cavernas naturales: eran inmensas calzadas subterráneas talladas en basalto volcánico y vetas de ónix negro, restos de una civilización imperial sepultada siglos atrás.",
        "A cada paso, la atmósfera se volvía más espesa y cáustica. El vapor púrpura serpenteaba por el suelo como una marea viva, lamiendo las botas y las grebas de acero. Sin embargo, los **respiradores rúnicos de loto de invierno** suministrados por Lortelle Kecheln funcionaban con una precisión milagrosa. A través de la doble membrana de latón y seda mágica, cada bocanada de aire entraba a mis pulmones limpia, fría y aromatizada con un sutil matiz de menta silvestre que despejaba los sentidos.",
        "En mi pecho, oculto bajo el peto de cuero tachonado, el broche de plata con la flor de escarcha que Lady Elysea me había ceñido palpitaba con un pulso rítmico y gélido. Su aura amortiguaba el bochorno sulfuroso del abismo, creando un microclima que blindaba mis terminaciones nerviosas contra la letargia ponzoñosa del miasma. Y en el bolsillo interior, el amuleto de jazmín sagrado de Kaoruko perfumaba mi piel con una calidez protectora.",
        "—Por todos los cielos... este lugar parece el décimo círculo de los infiernos del maestro Sakaki —susurró Kenichi Shirahama a mi derecha, ajustándose con manos temblorosas las correas de su petate mientras sus ojos se movían en todas direcciones tras el visor de cristal—. Si el Anciano Hayato estuviera aquí, juraría que me arrojaría por este precipicio con barriles de plomo amarrados a los brazos diciendo: «¡Es una excelente práctica para fortalecer los tendones de los tobillos, muchacho!»...",
        "Natsuki Subaru soltó una carcajada contenida, palmeándole el hombro con afecto fraternal:",
        "—Tranquilo, Kenichi. Esta vez no estás solo contra monstruos de artes marciales que rompen montañas con un estornudo. Tienes a toda la División Beta y a los Cuatro Estandartes cubriéndote la espalda. Si algo salta de esa niebla, Erza lo partirá en ocho pedazos antes de que toque el suelo.",
        "—Y si le queda una pizca de aliento, mi rifle de mithril le volará la sesera a cincuenta metros —añadió Dante Vance con una sonrisa burlona mientras amartillaba su cerrojo.",
        "Erza Scarlet caminaba en la vanguardia junto a mí, con la mirada de halcón fija en las sombras. Cuatro espadas de acero templado flotaban ingrávidas a su alrededor, trazando un halo defensivo en rotación pasiva:",
        "—No bajen la guardia. La acústica de estas galerías distorsiona el eco. Nos movemos en territorio donde el enemigo tiene ventaja de visión y número.",
        "El avance se ejecutó con la disciplina milimétrica de un ejército veterano.",
        "En la retaguardia de la fisura, el **Capitán Kaelen Thorne** y su **División Bastión de Hierro** clavaron tres inmensas empalizadas de acero enano y losas de basalto reforzadas con runas de tierra, sellando el cuello de botella del desfiladero. Thorne asignó una escuadra de choque para custodiar la barricada:",
        "—La ruta de extracción está anclada, Silva —afirmó el capitán de tercer año con un golpe seco de su puño contra el pavés—. Nadie nos cortará la retirada por la espalda.",
        "Por encima de nosotros, a unos quince metros de altura en las cornisas y estalactitas de basalto, los duelistas de la **División Colmillo Plateado** del **Capitán Bawden** avanzaban como espectros de plata. Con sus estoques encantados con amortiguación sonora y botas de cuero de reptil, purgaban en silencio a los centinelas Skaven que acechaban en las alturas antes de que pudieran dar la voz de alarma.",
        "En el flanco derecho, **Lighter** y **Subaru Hoshina** lideraban su división manteniendo los rifles a vapor y guanteletes térmicos en guardia baja, listos para desatar fuego de contención.",
        "* * *",
        "A medida que descendíamos hacia una cota más profunda, el suelo de ónix comenzó a inclinarse y el sonido de aguas subterráneas turbulentas llegó a nuestros oídos.",
        "**Agnes Tachyon** se detuvo un instante junto a una pared rezumante. Con unas pinzas de alquimista extrajo una muestra de la condensación viscosa que goteaba de la piedra y la vertió en un tubo de ensayo de cuarzo:",
        "—Fascinante y deplorable... —murmuró Agnes, agitando el reactivo mientras el líquido se teñía de un verde esmeralda fosforescente—. El ph es extremadamente ácido y la concentración de esporas fúngicas se ha triplicado en los últimos trescientos metros. Todo el miasma converge hacia el antiguo acueducto subterráneo de Valdemar. Si la corriente de agua freática se contamina por completo, el veneno llegará a las napas de riego de las tres provincias en menos de una semana.",
        "—Entonces no hay tiempo que perder —afirmó Ken, apretando los puños envueltos en vendas y ajustando su bandana roja con una mirada encendida—. Hay que reventar ese nido antes de que el veneno corra.",
        "Doblamos una gigantesca arcada de piedra labrada y desembocamos en la **Gran Falla de Ónix**.",
        "El espectáculo era sobrecogedor. Una sima ciclópea de más de cien metros de profundidad cortaba la montaña en dos. Sobre el abismo se alzaba un colosal puente de basalto negro, de seis metros de ancho y cincuenta de largo, que conectaba la galería de entrada con una bóveda cavernosa al otro lado. Pero la estructura estaba infestada: gruesas capas de biomasa correosa, capullos palpitantes de esporas y totems de huesos y chatarra cubrían los pilares del puente, mientras el aire vibraba con una pulsación rítmica y viscosa.",
        "—¡Alto! —susurró Erza levantando una mano enguantada.",
        "Desde las grietas del techo y los laterales del puente, un coro de chillidos guturales desgarró la penumbra.",
        "Decenas de ojos enrojecidos brillaron en la oscuridad. Una horda de **Alimañas de Choque Skaven** —ratas guerreras con cotas de malla oxidadas y alabardas envenenadas— surgió de las sombras en formación de cuña, mientras desde las estalactitas se descolgaban varios **Acechadores Nocturnos** armados con garras de garfio y dardos untados en bilis cáustica.",
        "—¡Emboscada en el puente! —gritó Dante, abriendo fuego de inmediato con su rifle de cerrojo. El proyectil de mithril atravesó el cráneo del primer líder de escuadra Skaven con un chasquido seco.",
        "La División Colmillo Plateado descendió desde las cornisas superiores con una precisión letal: Bawden atravesó a dos acechadores en pleno vuelo con estocadas sónicas que reventaron sus órganos internos sin darles tiempo a tocar el suelo. A la derecha, Lighter activó la cámara de combustión de sus guanteletes, lanzando una llamarada concentrada de vapor a doscientos grados que barrió a los piqueros enemigos del flanco.",
        "Sin embargo, dos Acechadores Nocturnos especialmente ágiles burlaron la primera línea, deslizándose por el pretil inferior del puente para caer a espaldas de la formación, directamente sobre Agnes Tachyon, quien aún sostenía sus tubos de ensayo.",
        "—¡Cuidado, señorita Tachyon! —gritó Kenichi Shirahama.",
        "Lo que ocurrió en ese microsegundo desafió la comprensión visual de los soldados convencionales.",
        "Kenichi no dudó. El terror cómico que solía paralizarlo se evaporó en el instante exacto en que una vida inocente estuvo en peligro. Su cuerpo, cincelado por miles de horas de tortura física bajo los maestros de Ryōzanpaku, reaccionó antes de que su cerebro formulara un pensamiento.",
        "Dio un paso deslizante hacia adelante y su presencia pareció dilatarse en una esfera invisible de tres metros: el ***Seikūken***.",
        "El primer acechador lanzó un zarpazo venenoso buscando la yugular de Agnes; Kenichi no bloqueó con fuerza bruta, sino que interceptó la muñeca de la criatura con la base de su palma izquierda (*Tenchi In'yō no Kamae*), desviando la trayectoria del filo por milímetros. Con un movimiento circular fluido de Jujitsu, se deslizó bajo el centro de gravedad del engendro, tomó su hombro y su cadera y ejecutó un magistral ***Tomoe Nage***: una proyección circular con apoyo de pierna que catapultó al Skaven sobre su cabeza, lanzándolo directamente al vacío del abismo sin fondo.",
        "El segundo acechador intentó clavarle una daga por el costado; Kenichi pivotó sobre su talón derecho, atrapó el brazo armado con una palanca articular a la velocidad del rayo y barrió las piernas del monstruo con un ***Uchimata*** devastador. El impacto contra el borde del puente quebró los huesos de la criatura, que cayó rodando hacia la sima oscura.",
        "—¡L-Lo siento mucho! ¡Pero no debían acercarse a la señorita Tachyon! —exclamó Kenichi, recuperando el aliento con los ojos muy abiertos, casi disculpándose con el aire mientras Agnes lo miraba boquiabierta, ajustándose las gafas con una fascinación científica absoluta.",
        "—¡Impresionante biomecánica neuromuscular...! Cero disipación de energía cinética, ángulo perfecto de torsión... —anotó Agnes en voz alta con euforia pura.",
        "—¡Buen trabajo, Kenichi! —bramé, encabezando la carga por el centro del puente—. ¡Astraea, abran la puerta de la caverna!",
        "* * *",
        "Arrollamos a los últimos defensores del puente y cruzamos el umbral hacia la gran cúpula natural que se abría al otro lado: la **Primera Cámara Nodriza**.",
        "El hedor aquí era una bofetada viscosa que hizo zumbar las membranas de nuestros respiradores. En el fondo de la caverna, donde las compuertas de piedra del antiguo acueducto vertían cascadas de agua envenenada, se alzaba una monstruosidad aberrante: un gigantesco **Saco Embrionario Nodriza**, una masa palpitante de carne corrupta de más de ocho metros de diámetro, cruzada por venas negras que bombeaban miasma violáceo directo a las corrientes subterráneas.",
        "Y frente al saco, custodiando el corazón del nido, aguardaba la pesadilla biológica de los clanes subterráneos.",
        "Un **Ogro Rata Acorazado**. Una mole de más de tres metros y medio de músculo hipertrofiado y pelaje purulento. Planchas de hierro ennegrecido habían sido clavadas directamente a su carne viva con remaches oxidados; en su brazo derecho llevaba injertada una colosal cuchilla de sierra impulsada por un reactor rudimentario de piedra de disformidad, cuyos vapores verdosos silbaban con furia.",
        "El monstruo rugió con una onda sónica que desprendió esquirlas de piedra del techo y cargó contra nosotros a una velocidad aterradora para su tonelaje.",
        "—¡Permítanme abrir el camino! —exclamó Erza Scarlet.",
        "Un destello mágico carmesí envolvió su figura en una fracción de segundo: ***Re-equip: Armadura de la Emperatriz del Fuego***. La coraza plateada de Erza fue sustituida por escamas de rubí carmesí, hombreras en forma de alas de fénix y dos cimitarras gemelas envueltas en un fuego ardiente que desintegraba el miasma al contacto. Erza se deslizó con gracia sobrehumana cortando en abanico a la escolta de guerreros de clan que flanqueaban a la bestia, neutralizando las toxinas ácidas con el calor de sus hojas.",
        "El Ogro Rata descargó su cuchilla mecánica contra el suelo en un tajo descendente que habría partido en dos a un caballo de guerra.",
        "Kaelen Thorne apareció en el punto exacto de impacto: clavó su escudo de torre en el basalto y concentró todo su maná telúrico en una barrera de granito reforzado. El choque fue brutal: el metal chilló y el suelo bajo las botas de Thorne se agrietó, pero el capitán de tercer año no cedió un solo milímetro.",
        "—¡Ahora, flancos! —rugió Thorne con los dientes apretados.",
        "Bawden se deslizó como un rayo de plata por debajo del brazo mecánico del coloso, ejecutando dos estocadas cruzadas que cercenaron los tendones de Aquiles de la bestia. El Ogro Rata aulló de dolor y cayó sobre una rodilla, perdiendo su base de equilibrio.",
        "Ese era mi momento.",
        "Activé la compresión neuromuscular en mis piernas con el juego de pies explosivo del Jeet Kune Do. Cerré la distancia en un pestañeo, canalizando la energía telúrica a través de los **brazaletes de Morvath**: los grabados de cobre y hierro negro resplandecieron con una luz ámbar profunda, comprimiendo el maná en una onda de choque densa.",
        "Salté hacia adelante, flexioné mi pierna derecha y descargué un ***Khao Dui*** —el rodillazo en lanza de Muay Thai— con todo el peso de mi cuerpo impulsado como un ariete directamente contra el esternón del monstruo.",
        "El impacto fue seco y demoledor. La placa de hierro clavada a su pecho se hundió en una concavidad perfecta; las costillas de la bestia se partieron con un crujido sordo y la tremenda masa del Ogro Rata fue levantada del suelo y proyectada hacia atrás cuatro metros, estrellándose contra la base del saco embrionario.",
        "—¡Ken, remata el núcleo! ¡Fuego solar total! —grité en el aterrizaje.",
        "Ken ya estaba en el aire.",
        "Su silueta se recortó contra la bóveda de la caverna, impulsada por un salto acrobático prodigioso. Su bandana roja ondeaba como una llamarada y sus dos puños ardían con una energía dorada incandescente, tan pura y brillante que disipó de golpe las tinieblas del abismo.",
        "—¡*Puño Solar del Fénix: Ignición Purificadora*! —bramó Ken con toda la fuerza de sus diecisiete años y el honor marcial de su maestra Hwa Young.",
        "Descendió en picado como un meteoro. Su puño envuelto en fuego solar se incrustó en el centro del pecho de la bestia y penetró hasta el núcleo de disformidad. La deflagración fue monumental: no fue una explosión caótica, sino una onda expansiva de calor estelar que calcinó al Ogro Rata de adentro hacia afuera, convirtiendo sus órganos corruptos y su carne mutada en cenizas doradas en menos de tres segundos.",
        "Sin detener su impulso, Erza y Ken sincronizaron sus auras: Erza desató una tormenta de llamas mágicas con sus cimitarras mientras Ken extendía su fuego solar en un barrido circular hacia el gigantesco Saco Embrionario Nodriza.",
        "El fuego sagrado devoró la biomasa purpúrea. Los quistes de miasma y las venas de esporas no estallaron ni liberaron gases ponzoñosos: la pureza molecular del fuego solar los desintegró en una combustión perfecta. El inmenso saco se colapsó sobre sí mismo, reduciéndose a un montículo de ceniza blanca inofensiva que fue arrastrado por la brisa.",
        "* * *",
        "Un silencio sobrecogedor descendió sobre la caverna.",
        "El aire, antes denso y asfixiante, se aclaró con una rapidez asombrosa. Las aguas del acueducto de Valdemar, liberadas de la ponzoña del nido, comenzaron a brotar cristalinas y transparentes sobre el lecho de piedra.",
        "Agnes Tachyon se quitó momentáneamente la boquilla del respirador, oliendo el ambiente con sus ojos desorbitados de emoción:",
        "—¡Increíble! La concentración de miasma en este sector cayó un cuarenta y cinco por ciento. ¡El fuego solar literalmente quemó los enlaces atómicos de la toxina sin generar subproductos venenosos!",
        "Kaelen Thorne respiró hondo, bajó su enorme escudo y soltó una carcajada ronca, mirando a Seier, Ken y Erza con un respeto militar reverencial:",
        "—Llevo tres años en la academia y nunca vi una purga de nido ejecutada con semejante limpieza quirúrgica. Muchachos... son unos auténticos monstruos de combate.",
        "Bawden envainó su estoque con un floreo elegante y sonrió:",
        "—La F-2 ciertamente hace honor a su reputación legendaria. El frente exterior de Alden debe estar sintiendo el alivio en este mismo momento.",
        "Lighter se acercó trotando, secándose el sudor de la frente mientras chocaba los puños con Ken y conmigo:",
        "—¡Eso fue hermoso! Aunque no se confíen... miren lo que había detrás de esa masa de carne podrida.",
        "Nos acercamos al fondo de la caverna purificada.",
        "Donde antes palpitaba el saco embrionario, el colapso de la biomasa y la fractura de la roca por la combustión habían dejado al descubierto una estructura ciclópea: un inmenso portón de piedra negra grabado con relieves de la dinastía fundadora del Reino de Solaria y sellado con gigantescas cerraduras rúnicas.",
        "Más allá de ese portal, las corrientes de aire susurraban desde una profundidad aún mayor. El mapa de inteligencia que Lortelle me había entregado era inequívoco: tras ese umbral yacían las galerías inferiores, la **Segunda Cámara Nodriza** y el corazón mismo de las Ruinas Negras.",
        "Palpé el broche de escarcha de Lady Elysea y el jazmín de Kaoruko en mi pecho, sintiendo sus presencias dándome fuerza a la distancia. Miré a Erza, a Ken, a Kenichi y a los comandantes de tercer año. Todos estaban intactos, con la moral por las nubes y el acero afilado.",
        "—Aseguren el perímetro y beban agua —dije con voz serena y firme—. La primera nodriza cayó. Ahora vamos a descender a las profundidades."
    ]
}




# ===============================================================================
# FUNCIONES DE GESTIÓN Y COMPILACIÓN
# ===============================================================================

def guardar_capitulo_individual(cap_dict, base_dir=NOVELA_DIR):
    """Guarda el archivo .docx individual del capítulo."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    num = cap_dict.get("num", 0)
    if num == 0 or "PROLOGO" in cap_dict.get("header", "").upper() or "PRÓLOGO" in cap_dict.get("header", "").upper():
        base_name = "00_Prologo"
    else:
        base_name = f"{num:02d}_Capitulo_{num}"
        
    docx_path = os.path.join(capitulos_dir, f"{base_name}.docx")
    create_single_chapter_document(
        cap_dict,
        docx_path,
        novel_title="El camino del reencarnado",
        author="Seier Silva"
    )
    print(f"[OK] Documento Word individual guardado:\n  {docx_path}")
    return docx_path


def cargar_todos_los_capitulos(base_dir=NOVELA_DIR):
    """Lee todos los capítulos directamente desde la carpeta capitulos/*.docx."""
    capitulos_dir = os.path.join(base_dir, "capitulos")
    docx_files = sorted(glob.glob(os.path.join(capitulos_dir, "*.docx")))
    
    # Intentar cargar mapa de ilustraciones
    ilustraciones = {}
    try:
        from ilustraciones_capitulos import ILUSTRACIONES_POR_CAPITULO
        ilustraciones = ILUSTRACIONES_POR_CAPITULO
    except Exception:
        pass
        
    lista_capitulos = []
    
    for fpath in docx_files:
        filename = os.path.basename(fpath)
        num = None
        m_num = re.match(r"^(\d+)_", filename)
        if m_num:
            num = int(m_num.group(1))
            
        doc = docx.Document(fpath)
        if not doc.paragraphs:
            continue
            
        title_p = doc.paragraphs[0].text.strip()
        parts = title_p.split("\n")
        header = parts[0].strip() if len(parts) > 0 else "Capítulo"
        title = parts[1].strip() if len(parts) > 1 else ""
        
        paragraphs = []
        for p in doc.paragraphs[1:]:
            t = p.text.strip()
            if not t:
                continue
            if t in ["— ✦ —", "— ◆ —", "✦ ✦ ✦", "— ✧ —"]:
                continue
            if t.startswith("Ilustración:") or t.startswith("Ilustracion:"):
                continue
            paragraphs.append(t)
            
        chap_dict = {
            "header": header,
            "title": title,
            "paragraphs": paragraphs
        }
        
        # Asociar ilustraciones
        if num in ilustraciones:
            illus_info = ilustraciones[num]
            if "illustrations" in illus_info:
                chap_dict["illustrations"] = illus_info["illustrations"]
            elif "illustration" in illus_info:
                chap_dict["illustration"] = illus_info["illustration"]
                chap_dict["illustration_caption"] = illus_info.get("illustration_caption", "")
                
        lista_capitulos.append(chap_dict)
        
    return lista_capitulos


def compilar_libro_completo(base_dir=NOVELA_DIR):
    """Recompila el libro maestro completo a partir de todos los .docx de capitulos/."""
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    
    capitulos = cargar_todos_los_capitulos(base_dir)
    print(f"[*] Compilando Libro Maestro con {len(capitulos)} capítulos (desde .docx)...")
    
    compilado_dir = os.path.join(base_dir, "compilado")
    os.makedirs(compilado_dir, exist_ok=True)
    master_docx = os.path.join(compilado_dir, "El_Camino_del_Reencarnado.docx")
    
    try:
        create_novel_document(titulo, subtitulo, autor, capitulos, master_docx)
        print(f"[OK] Libro Maestro actualizado exitosamente en:\n  {master_docx}")
    except PermissionError:
        print(f"[AVISO] '{master_docx}' está abierto en Microsoft Word.")
        backup_master = os.path.join(compilado_dir, "El_Camino_del_Reencarnado_temp.docx")
        create_novel_document(titulo, subtitulo, autor, capitulos, backup_master)
        print(f"[OK] Guardado en copia auxiliar:\n  {backup_master}")


def ejecutar():
    print("===============================================================")
    print(f" AGREGAR / ACTUALIZAR: {CAPITULO['header']} - {CAPITULO['title']}")
    print("===============================================================")
    for base in [CURRENT_DIR, NOVELA_DIR, PROYECTO_DIR]:
        if os.path.exists(base):
            print(f"[*] Procesando en: {base}")
            guardar_capitulo_individual(CAPITULO, base)
            compilar_libro_completo(base)
    print("===============================================================")
    print(" ¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("===============================================================")

if __name__ == "__main__":
    ejecutar()
