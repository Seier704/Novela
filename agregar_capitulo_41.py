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

CAPITULO_41 = {
    "header": "Capítulo 41",
    "title": "La Víspera del Asedio",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "vlad.png"),
            "caption": "Ilustración: Vlad (Caballero Aprendiz del Ducado)"
        },
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "agnes_tachyon.png"),
            "caption": "Ilustración: Agnes Tachyon (Maga Experimental de la Academia)"
        }
    ],
    "paragraphs": [
        "El retorno a la superficie no trajo consigo el alivio que todos esperábamos.",
        "Cuando el Escuadrón 7 y el contingente de la Clase C cruzamos el imponente portón levadizo de la Fortaleza de la Roca Carmesí —el bastión militar avanzado que servía como cuartel general para la incursión de primer año—, el ambiente no era el de una base en descanso, sino el de un polvorín a punto de estallar.",
        "Campanas de bronce tañían con un repique frenético y continuo desde lo alto de los torreones.",
        "Soldados regulares y centinelas de la guarnición corrían de un lado a otro cargando cajas de flechas, lanzas pesadas y calderos de brea hacia las almenas. En el patio de armas, varios sanadores atendían a los heridos que habían logrado escapar de los túneles inferiores, mientras un murmullo denso de incertidumbre y preocupación flotaba en el aire gélido de la cordillera.",
        "—¡Atención a todos los cadetes y escuadrones recién llegados! —bramó la voz amplificada por maná del Instructor Principal Gerald desde las escaleras del cuartel táctico—. ¡Reúnanse de inmediato en la explanada central!",
        "Kaoruko ajustó la empuñadura de su espada y me miró de reojo con el ceño fruncido.",
        "—Esto no es un simulacro, Seier... Algo muy grave acaba de suceder.",
        "—Los Skavens que enfrentamos abajo no eran simples rezagados —respondí en voz baja, sintiendo una vibración sorda e inquietante bajo la suela de mis botas—. Se estaban organizando.",
        "* * *",
        "En la explanada, los instructores veteranos desplegaron un mapa topográfico encantado sobre una mesa de campaña.",
        "El rostro curtido de Gerald parecía tallado en granito.",
        "—Hace dos horas, nuestras patrullas de reconocimiento aéreo y los sabuesos de rastreo confirmaron el peor de los escenarios —anunció el comandante ante la multitud de estudiantes—. A menos de cuatro leguas al oeste de este bastión, oculto en una falla geológica que no figuraba en los registros cartográficos, se descubrió una enorme madriguera subterránea Skaven. Y no se trata de una colonia pequeña: es un nido de guerra en plena ebullición.",
        "Un escalofrío colectivo recorrió las filas de los cadetes de dieciséis años.",
        "—Están cavando túneles de asedio hacia nuestras murallas —continuó Gerald con gravedad punzante—. Acumulan armas químicas de piedra bruja, alimañas de asalto y miles de esclavos de choque. Su objetivo es barrer esta fortaleza en un golpe sorpresa antes de que el reino pueda reaccionar.",
        "—¿Y los refuerzos, señor? —preguntó un cadete de la Clase B con voz temblorosa—. Ya habrán enviado los mensajeros al Reino y a la sede central de la academia, ¿verdad?",
        "—Enviamos señales de emergencia astrales y tres mensajeros alados de inmediato —respondió el instructor con amargura contenida—. Pero la capital real está a más de cinco días de marcha forzada a través de los pasos montañosos. Los refuerzos tardarán en llegar. Nadie vendrá a salvarnos antes de que empiece la batalla.",
        "El pánico amenazó con quebrar la disciplina del patio. Abandonar la fortaleza no era una opción; con cientos de estudiantes heridos y el terreno montañoso infestado de saboteadores que cavaban bajo tierra, una retirada abierta se convertiría en una masacre segura.",
        "Los instructores veteranos intercambiaron miradas sombrías. Sabían cuál era su deber.",
        "—Defenderemos la fortaleza —sentenció Gerald, desenvainando su pesada hoja de oficial—. No cederemos un solo palmo de piedra. Pero necesitamos cada brazo, cada hechizo y cada espada disponible.",
        "A pesar del miedo palpable, la determinación de los estudiantes de la Clase F-2 y de varios escuadrones destacados no tardó en encenderse. Ken dio un paso al frente con el puño en llamas; Yami apretó la vaina de su katana oscura con una sonrisa salvaje; Erza y Akame se colocaron hombro a hombro listas para el combate. Nadie iba a huir.",
        "Fue en ese instante de debate táctico cuando dos figuras inesperadas rompieron el círculo.",
        "* * *",
        "La primera en adelantarse fue una joven de dieciséis años de porte peculiar.",
        "Llevaba una bata de alquimista salpicada de reactivos fosforescentes sobre su uniforme de combate, y sobre su alborotada melena castaña asomaban unas orejas que se movían con curiosidad hiperactiva. Sus ojos carmesí-ámbar brillaban con una mezcla inconfundible de intelecto brillante y pura locura científica. Se trataba de **Agnes Tachyon**, una maga investigadora famosa en la academia por sus experimentos poco ortodoxos y su fascinación por desafiar los límites de la teoría mágica.",
        "—¡Pido la palabra, distinguidos señores de la guerra! —exclamó Agnes con una sonrisa radiante y descarada, arrojando sobre el mapa táctico tres pergaminos repletos de fórmulas geométricas y varios frascos que burbujeaban con un líquido dorado inestable.",
        "Gerald arqueó una ceja, visiblemente irritado por la interrupción.",
        "—Señorita Tachyon, este no es momento para sus experimentos de laboratorio...",
        "—Al contrario, mi estimado instructor, es el momento *perfecto* para la ciencia aplicada —interrumpió ella, ajustándose con gracia el cuello de su bata—. Si esperamos pasivamente en las murallas a que esas ratas repugnantes escalen las piedras o perforen los cimientos, la masa numérica nos aplastará sin remedio. Pero yo tengo una propuesta mucho más... explosiva y divertida.",
        "Agnes señaló los desfiladeros y las colinas que rodeaban el perímetro exterior de la fortaleza.",
        "—He perfeccionado una serie de runas de ignición por compresión cinética —explicó, con los ojos centelleando de emoción—. Son sellos mágicos camuflables que permanecen latentes hasta que una masa física o un flujo de peso específico pisa su superficie. En cuanto los primeros batallones de Skavens carguen en estampida por los cuellos de botella o intenten perforar el suelo desde abajo... ¡BUM! Una reacción en cadena convertirá todo el valle exterior en un horno de metralla mágica antes de que alcancen a tocar nuestras puertas.",
        "Los presentes guardaron silencio, procesando la audacia del plan.",
        "—El único dilema —añadió Agnes, posando su mirada inquisitiva sobre mí—, es que necesito colocarlas a gran profundidad en la roca sólida sin alertar a sus excavadores subterráneos. Requiere alguien con un control telúrico exquisito y la fuerza para moldear el lecho de piedra en tiempo récord...",
        "—Yo me encargo de eso —di un paso al frente, respondiendo con firmeza—. Con mi maná de tierra puedo abrir fisuras subterráneas, sellar las runas bajo las capas de roca y cerrar el terreno sin dejar rastro en la superficie.",
        "Agnes me dedicó una sonrisa de satisfacción demencial, frotándose las manos.",
        "—¡Magnífico! Un sujeto de pruebas... quiero decir, ¡un colaborador con visión de futuro! Te aseguro, grandulón, que esas ratas van a bailar al ritmo de mis fuegos artificiales.",
        "* * *",
        "Antes de que los instructores pudieran objetar los detalles técnicos, una segunda voz, grave, mesurada y cargada de una autoridad innata, se alzó desde el flanco izquierdo.",
        "Un joven de porte impecable dio un paso adelante.",
        "Tenía dieciséis años, como todos nosotros, pero su semblante desprendía la serenidad implacable de un veterano de frontera. Su cabello rubio dorado caía sobre una mirada de ojos azules fríos y calculadores como témpanos de hielo. Llevaba una capa de viaje oscura abrochada al hombro por una fíbula de hierro y una coraza ligera ajustada con maestría sobre su pecho. En su cadera descansaba una espada bastarda de acero templado. Era **Vlad**, un caballero aprendiz enviado directamente desde las tierras del Ducado para foguearse en combate real y perfeccionar su camino hacia el título de caballero.",
        "Vlad cruzó su puño derecho sobre el pecho en un saludo militar perfecto hacia Gerald y los instructores.",
        "—Instructor Gerald, señores oficiales —habló Vlad con una entonación calmada y firme que impuso orden en toda la explanada—. El plan de la señorita Tachyon es brillante para quebrar el primer embate, pero una trampa solo compra tiempo. La verdadera prueba vendrá cuando las líneas de choque colisionen en las brechas.",
        "Vlad se acercó a la mesa, señalando con un dedo enguantado los puntos ciegos de las murallas secundarias.",
        "—El mayor peligro para nosotros es la falta de coordinación entre los cadetes. Si el pánico se apodera de los estudiantes inexpertos, las murallas caerán por puro caos interno. Por eso propongo una división estratégica de mando: yo asumiré la organización y el liderazgo de las defensas estudiantiles en las almenas y las barricadas exteriores.",
        "Los oficiales lo miraron con sorpresa, evaluando la audacia de su propuesta.",
        "—Fui instruido en doctrina de asedios y formaciones de infantería pesada por los mariscales del Ducado —continuó Vlad con total seguridad—. Si los cadetes nos organizamos en falanges cerradas y rotaciones de guardia bajo mis órdenes y el apoyo de los líderes de escuadrón, podremos contener a la infantería Skaven en las murallas secundarias. Y eso, señores instructores, les otorgará a ustedes total libertad táctica para actuar como una fuerza de choque pesada y móvil, cazando a los líderes del nido, los sacerdotes de plaga y las abominaciones gigantes antes de que rompan las puertas principales.",
        "Las palabras de Vlad cayeron con el peso demoledor de la lógica militar más pura.",
        "Gerald contempló al joven noble durante varios segundos, notando el brillo indomable y disciplinado en sus ojos azules. Finalmente, una sonrisa de respeto asomó en la barba del curtido instructor.",
        "—Tienes agallas, muchacho. Y un cerebro bien puesto sobre los hombros —asintió Gerald con fuerza—. Aprobado. El cadete Vlad coordinará las líneas defensivas de los estudiantes en los bastiones este y norte. Los instructores asumiremos la vanguardia pesada en la puerta principal. Y tú, Seier Silva, acompaña a la señorita Tachyon al perímetro exterior; que no quede un solo túnel sin una de esas malditas sorpresas explosivas.",
        "* * *",
        "Las horas siguientes fueron una carrera contrarreloj contra la oscuridad.",
        "Bajo la dirección táctica de Vlad, los cadetes de todas las clases comenzaron a trabajar como un verdadero ejército. Ken y los suyos transportaban empalizadas y sacos de arena; Erza organizaba las reservas de armas; Lighter y Hoshina coordinaban los arqueros y francotiradores mágicos en los puestos de tiro elevados. La fría eficiencia de Vlad infundió una calma contagiosa en los corazones más asustados, transformando la incertidumbre en una disciplina de hierro.",
        "Mientras tanto, en las pendientes y cañones exteriores, Agnes Tachyon y yo trabajábamos sin descanso.",
        "Cada vez que ella preparaba una de sus intrincadas runas de presión, yo colocaba mi palma sobre la roca virgen, canalizando maná de tierra para abrir un receptáculo hueco en el estrato mineral, introduciendo el dispositivo y sellándolo con una capa milimétrica de basalto que no alteraba la pisada pero transmitiría la onda cinética directo al núcleo detonador.",
        "—¡Arte puro, Silva! ¡Esto es verdadera sinergia alquímica! —exclamaba Agnes con risitas maniáticas cada vez que una nueva sección quedaba activa.",
        "Cuando el sol finalmente se ocultó tras los picos escarpados de la cordillera, el silencio cayó sobre la Fortaleza de la Roca Carmesí.",
        "Regresé a la muralla principal, limpiándome el sudor de la frente, donde Kaoruko me esperaba con una cantimplora de agua y su espada desenvainada, reflejando las primeras estrellas del cielo nocturno. Vlad nos observó desde su puesto de mando, asintiendo en silencio hacia nosotros con una mirada de camaradería guerrera.",
        "De pronto, el viento del oeste cambió de dirección.",
        "Un hedor pestilente a azufre, podredumbre y orina de rata impregnó la noche. Y desde las profundidades del cañón rocoso, el repiqueteo metálico de millones de garras y el tamborileo infernal de los tambores Skaven hicieron temblar las piedras de la fortaleza.",
        "La invasión había comenzado.",
        "Pero sobre las almenas, hombro a hombro entre el fuego de las runas y el acero de los camaradas, estábamos listos para la guerra."
    ]
}

def compilar_capitulo_41():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 41              ")
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
        CAPITULO_40, CAPITULO_41
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 41
    txt_cap41 = os.path.join(capitulos_dir, "41_Capitulo_41.txt")
    with open(txt_cap41, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_41['header'].upper()}: {CAPITULO_41['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_41['paragraphs']))
    print(f"[OK] Texto plano Capítulo 41 guardado en:\n  {txt_cap41}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 41
    docx_cap41 = os.path.join(capitulos_dir, "41_Capitulo_41.docx")
    create_single_chapter_document(CAPITULO_41, docx_cap41, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 41 guardado en:\n  {docx_cap41}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 41)
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
    compilar_capitulo_41()
