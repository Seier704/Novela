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

CAPITULO_44 = {
    "header": "Capítulo 44",
    "title": "Danza de Sangre y Escarcha",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "akame.png"),
            "caption": "Ilustración: Akame (La Hoja Silenciosa de la F-2)"
        },
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "rukia_renji.jpg"),
            "caption": "Ilustración: Rukia y Renji (Dúo de Hielo y Furia)"
        }
    ],
    "paragraphs": [
        "El aire en el patio interior era un veneno espeso que quemaba la garganta como ácido hirviente.",
        "Cuando Kaoruko y yo atravesamos a toda prisa el corredor de piedra que conectaba la muralla con el ala médica, la escena que nos recibió fue una bofetada despiadada de la cruda realidad de la guerra. No era la gloria de los torneos ni el combate limpio de la academia; era una masacre silenciosa y repugnante.",
        "El hospital de campaña estaba envuelto en una densa niebla verdosa.",
        "En el umbral de las tiendas de lona, el Sargento Marcus —un instructor veterano de la guarnición que nos había enseñado guardias de escudo apenas dos días atrás— yacía sobre los adoquines. Su armadura de acero pesado estaba derretida por el vómito corrosivo de una bomba de plaga; con su propio cuerpo había cubierto a dos cadetes heridos para evitar que el veneno los consumiera. Cuando pasé a su lado, sus ojos vidriosos miraban al cielo sin parpadear, con la mano crispada sobre su escudo partido.",
        "A pocos metros, dos estudiantes de logística de la Clase D, Thomas y Mia, de tan solo dieciséis años, estaban tendidos en el suelo. Habían intentado evacuar las camillas de pociones cuando una jauría de Acechadores de la Muerte los emboscó por la espalda. Sus manos inertes todavía se aferraban a un rollo ensangrentado de vendas.",
        "Un nudo ardiente de furia y dolor me desgarró el estómago.",
        "—¡Malditos bastardos...! —la voz de Kaoruko tembló en un sollozo ahogado, con los ojos empañados en lágrimas mientras sus nudillos se ponían blancos sobre la empuñadura de su katana.",
        "Seis Monjes de la Plaga, con túnicas hechas jirones y rostros hinchados por pústulas purulentas, avanzaban hacia la carpa donde Freya y Elysea intentaban mantener a salvo a los heridos más graves. Las criaturas hacían girar pesados mayales de hierro llenos de brasas tóxicas, preparándose para demoler la enfermería.",
        "Pero no alcanzaron a dar un paso más.",
        "* * *",
        "—¡¡¡RUGIDO DE SANGRE, ZABIMARU!!! —un bramido salvaje y atronador desgarró la niebla venenosa.",
        "Una silueta alta, con una cabellera roja como el fuego y tatuajes tribales que ardían con maná carmesí, cayó en picada desde el tejado de la capilla.",
        "Era Renji.",
        "Su pesada espada dentada no cortó como un filo ordinario; la hoja se fragmentó en decenas de segmentos de acero afilado unidos por un cordón de maná reforzado, extendiéndose en un látigo de colmillos gigantes que barrió el patio en un arco demoledor. El impacto de la hoja articulada impactó de lleno en los seis monjes, quebrando sus mayales, triturando sus costillas y arrojando sus cuerpos mutilados contra los muros de piedra como muñecos de trapo.",
        "Renji aterrizó sobre los adoquines con un crujido sordo, recogiendo los segmentos de su arma con un tirón seco de su brazo tatuado. La sangre negra de las ratas salpicaba su rostro, pero en sus ojos brillaba una ferocidad inquebrantable.",
        "—¡Nadie toca a los heridos mientras yo siga respirando! —rugió Renji, colocándose de espaldas a la tienda médica con la guardia alta.",
        "Sin embargo, la verdadera amenaza no eran solo las alimañas cuerpo a cuerpo. La nube de gas tóxico de piedra bruja continuaba expandiéndose desde el socavón central, reptando hacia el interior de la carpa y amenazando con asfixiar a los supervivientes en cuestión de segundos.",
        "Fue entonces cuando una presencia etérea y gélida descendió sobre el campo de batalla.",
        "* * *",
        "Rukia Kuchiki avanzó con pasos silenciosos y elegantes, como si flotara sobre el fango ensangrentado.",
        "Su rostro sereno y aristocrático no mostraba miedo, sino una concentración pura y casi sagrada. Sostenía su espada delgada con un agarre invertido, cuya cinta blanca ondulaba en el viento como un pétalo en la tormenta.",
        "—*Danza Primera: Luna Blanca* —susurró Rukia con una voz cristalina que calmó el caos del aire.",
        "Clavó la punta de su hoja en el centro del patio.",
        "Un círculo de luz blanca e inmaculada brotó desde el suelo, expandiéndose a una velocidad vertiginosa. Una temperatura de cero absoluto barrió el recinto en una onda expansiva silenciosa. En un instante milagroso, la niebla verde y purulenta del veneno no se dispersó; simplemente se congeló en el aire, transformándose en diminutos copos de nieve cristalizada de color esmeralda que cayeron sobre los adoquines y se hicieron polvo inofensivo bajo nuestras botas.",
        "El aire volvió a ser puro y respirable.",
        "Ocho Acechadores de la Muerte intentaron saltar sobre Rukia desde las vigas del campanario para interrumpir su magia. La joven espadachina ni siquiera levantó la mirada.",
        "—*Danza Siguiente: Onda Blanca.*",
        "Rukia trazó cuatro cortes en el aire. Ráfagas de escarcha pura surgieron de la punta de su katana, congelando a las criaturas en pleno vuelo. Los ocho asesinos cayeron al suelo convertidos en estatuas de hielo macizo que se despedazaron en mil pedazos al impactar contra el suelo.",
        "Freya y Elysea asomaron la cabeza desde la tienda, con lágrimas en los ojos pero a salvo gracias al santuario gélido de la chica.",
        "* * *",
        "Pero la raíz del mal seguía viva.",
        "Sobre el montículo de escombros del socavón, el Sacerdote de la Plaga Skaven levantó su báculo de carnero, comenzando a recitar una maldición prohibida para detonar un artefacto de piedra bruja que haría colapsar toda la montaña sobre nosotros.",
        "El brujo creía estar a salvo en la retaguardia, protegido por el humo y las sombras.",
        "No sabía que la verdadera muerte no hace ruido.",
        "Un destello carmesí surcó la oscuridad detrás de él.",
        "Akame apareció de la nada, materializándose en su espalda como un fantasma implacable. Su larga cabellera negra ondeaba al viento y sus ojos rojos centelleaban con una frialdad absoluta que helaba la sangre mucho más que la magia de Rukia.",
        "No hubo discursos ni advertencias.",
        "La katana de Akame cortó el aire en una fracción de segundo, trazando una línea limpia y perfecta a través del cuello y el corazón del brujo Skaven.",
        "—*Eliminar* —pronunció Akame en un susurro gélido.",
        "La maldición de su filo se expandió de inmediato por el torrente sanguíneo del sacerdote. Runas negras se extendieron como grietas de ceniza por la piel del monstruo, deteniendo su corazón al instante y consumiendo su cuerpo en una pira de polvo antes de que sus garras pudieran tocar el detonador.",
        "El báculo maldito cayó al suelo, partido en dos pedazos inertes.",
        "* * *",
        "El silencio regresó al patio interior.",
        "Solo se escuchaba el murmullo del viento helado y el llanto bajito de los cadetes heridos que habían sobrevivido al horror.",
        "Akame sacudió la sangre de su katana con un movimiento seco de muñeca y la envainó con un clic metálico. Se acercó despacio hacia nosotros, contemplando los cuerpos sin vida del Sargento Marcus y de los dos chicos de la Clase D. Una sombra de dolor profundo cruzó sus ojos carmesí.",
        "—Llegamos a tiempo para la mayoría... —murmuró Akame, con la voz cargada de un pesar contenido—. Pero no pudimos salvar a todos.",
        "Me arrodillé junto al cuerpo del sargento y de los dos estudiantes. Con mi mano grande y ensangrentada, les cerré con cuidado los párpados.",
        "Eran compañeros. Chicos de nuestra misma edad que tenían sueños, familias y un futuro por delante que les acababa de ser arrebatado en un pozo de podredumbre.",
        "Apreté los puños con tanta fuerza que mis nudillos crujieron como piedras.",
        "Kaoruko se arrodilló a mi lado, apoyando una mano temblorosa sobre mi hombro. Renji clavó su espada en el suelo inclinando la cabeza con respeto marcial, mientras Rukia juntaba las palmas de sus manos en una breve y silenciosa oración por las almas de los caídos.",
        "—Sus muertes no van a ser en vano —dije, y mi voz salió tan grave y cargada de una furia telúrica que hizo estremecer el suelo alrededor—. Esos engendros van a pagar por cada gota de sangre que derramaron en esta fortaleza.",
        "Me puse de pie, mirando hacia el portón principal donde el fragor de la batalla exterior seguía rugiendo bajo el mando de Vlad, Yami y Gerald.",
        "La retaguardia estaba segura.",
        "Ahora era el momento de aplastar el asedio de raíz."
    ]
}

def compilar_capitulo_44():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 44              ")
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
        CAPITULO_40, CAPITULO_41, CAPITULO_42, CAPITULO_43, CAPITULO_44
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 44
    txt_cap44 = os.path.join(capitulos_dir, "44_Capitulo_44.txt")
    with open(txt_cap44, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_44['header'].upper()}: {CAPITULO_44['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_44['paragraphs']))
    print(f"[OK] Texto plano Capítulo 44 guardado en:\n  {txt_cap44}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 44
    docx_cap44 = os.path.join(capitulos_dir, "44_Capitulo_44.docx")
    create_single_chapter_document(CAPITULO_44, docx_cap44, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 44 guardado en:\n  {docx_cap44}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 44)
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
    compilar_capitulo_44()
