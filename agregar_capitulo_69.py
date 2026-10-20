# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_69 = {
    "header": "Capítulo 69",
    "title": "Emboscada en Sauceda: Sangre, Fuego y el Juicio del Acero",
    "paragraphs": [
        "El atardecer cayó sobre el valle de Sauceda con la pesadez de una mortaja teñida en tonos cobrizos y violetas.",
        "Avanzando a paso forzado por los senderos ocultos que bordeaban el río de los Sauces, la Sección Uno y Julian alcanzaron el perímetro de la aldea dos horas antes del crepúsculo. Sauceda era un conjunto desolado de medio centenar de chozas de adobe y techumbres de paja podrida, sumidas en un silencio aterrador donde ni siquiera los perros se atrevían a ladrar.",
        "En el cuartel bajo, una tosca empalizada de troncos junto al molino harinero, encontramos lo que quedaba de la verdadera milicia cívica del condado.",
        "Eran treinta hombres demacrados, con corazas de cuero agrietadas, lanzas herrumbrosas y ojos consumidos por el insomnio y la desesperación. Su líder, el **Sargento Roderick**, un veterano de barba gris y una cuenca vacía cruzada por una vieja cuchillada, nos encañonó con su ballesta apenas cruzamos la trinchera.",
        "—¡Atrás! —rugió Roderick con la voz rota—. ¡No nos queda más grano ni ganado para sus requisas de mierda!",
        "Julian dio un paso al frente y se despojó de la capucha de explorador. Extendió su mano derecha hacia la luz del farol, dejando al descubierto el anillo de plata maciza con el halcón bicéfalo grabado.",
        "—Roderick. Mírame a los ojos —ordenó Julian con voz firme pero cargada de una emoción reprimida—. No vengo como el hijo del conde que los ha desangrado. Vengo como el segundo heredero a darles la orden que llevan tres meses esperando: defiendan a sus familias.",
        "El veterano soltó la ballesta en el barro. Miró el sello de plata, miró el rostro exhausto de Julian y cayó de rodillas, con las lágrimas surcándole las arrugas mugrientas del rostro:",
        "—¡Por los dioses... mi señor Julian! Creímos que la Casa Morvath nos había condenado a todos al matadero...",
        "—La justicia tardó, pero está aquí —intervine, dando un paso adelante con el peso macizo de mi porte militar—. Soy Seier Silva, Caballero Defensor de la Corona. Esta noche, los hombres de Valerius y los cultistas de Khorne vienen a llevarse a sesenta de sus mujeres y niños para degollarlos en las minas. Pero esta noche, la presa va a morder.",
        "* * *",
        "El plan táctico se desplegó con la precisión despiadada de un reloj de guerra.",
        "Vlad y Roderick organizaron la evacuación inmediata de todos los civiles vulnerables hacia el sótano abovedado del viejo granero de piedra comunal, atrancando los accesos con vigas de roble y apostando a diez lanceros en el interior con órdenes de resistir hasta el último aliento.",
        "En las cabañas exteriores, dejamos velas encendidas en las ventanas y muñecos de trapo sobre los lechos para simular que la aldea dormía indefensa.",
        "Las posiciones de emboscada se sellaron en absoluto sigilo:",
        "Faelar se apostó en la torre del campanario del molino con su arco largo de tejo; Kaelan se agazapó sobre los tejados de paja con las cuchillas retráctiles listas; Borgar y Vlad tomaron el callejón central junto a veinte milicianos parapetados tras carros de heno volcados; Akame se fundió con las sombras de la acequia este; y yo me situé en el centro de la plaza de barro junto a Julian y Vance.",
        "La medianoche envolvió a Sauceda en una neblina densa y helada. El único sonido era el murmullo del río.",
        "Hasta que el aire se pudrió de golpe.",
        "* * *",
        "Un aullido gutural, bestial y enloquecido rasgó el silencio desde el linde del bosque.",
        "Por el camino principal irrumpieron treinta fanáticos del Culto a Khorne, corriendo con zancadas desquiciadas sobre el fango. Iban con los torsos desnudos cubiertos de runas de sangre coagulada, bramando letanías al Dios de la Sangre y blandiendo enormes hachas de carnicero y mayales de púas. Detrás de ellos cabalgaban veinte mercenarios de la guardia personal de Valerius, escoltando tres carromatos con jaulas de hierro vacías.",
        "Al frente de los mercenarios venía el **Capitán Malrik**, el perro tuerto y vicioso de Valerius, que reía a carcajadas ajustando las riendas con el cofre de hierro atado a la grupa de su montura:",
        "—¡Rompan las puertas y aten a las perras y a los cachorros! —bramó Malrik escupiendo al barro—. ¡El señor Valerius y el conde tienen prisa por llenar las jaulas antes del alba!",
        "Los cultistas se abalanzaron hacia las primeras cabañas rugiendo sedientos de matanza.",
        "Fue en ese instante exacto cuando silbé.",
        "* * *",
        "El infierno estalló sobre los invasores.",
        "Desde el campanario del molino, un proyectil rúnico imbuido en maná de viento silbó en la noche. La flecha de Faelar atravesó el ojo del cultista que encabezaba la carga, clavándolo contra el poste de madera con tal impacto que su cráneo estalló hacia atrás.",
        "—¡¡FUEGO Y HIERRO!! —rugió Borgar desde el callejón.",
        "La barricada de carros de heno se abrió de golpe. Veinte milicianos arremetieron con picas largas y Vlad cargó al frente con su espada bastarda destellando en un arco letal que decapitó a dos mercenarios antes de que pudieran desmontar.",
        "Kaelan cayó desde el alero como un felino de sombras, clavando sus dos cuchillas retráctiles en la nuca de un jinete y degollando al caballo en una maniobra acrobática que desató el pánico en la retaguardia.",
        "Pero la verdadera pesadilla para los herejes fue Akame.",
        "Emergió de la niebla como una ráfaga negra y silenciosa. Su katana *Murasame* no trazaba brillos; trazaba muerte absoluta. Cada estocada era un murmullo helado: un corte en la carótida, una estocada al corazón, una decapitación limpia sin salpicar una sola gota sobre su uniforme.",
        "—*Eliminar... eliminar... eliminar...* —susurraba Akame mientras tres cultistas de Khorne caían fulminados al lodo antes de que sus hachas tocaran el suelo.",
        "* * *",
        "Tres berserkers colosales de Khorne, con el rostro deformado por cicatrices rituales y los ojos inyectados en sangre por los narcóticos de batalla, me vieron en el centro de la plaza y cargaron rugiendo al unísono con sus hachas dobles en alto.",
        "Bajé mi centro de gravedad y canalicé maná telúrico puro hacia mis cuádriceps y antebrazos.",
        "El primer berserker descargó un hachazo diagonal con fuerza para partir un roble. No retrocedí; di un paso angular al frente, esquivando la hoja por escasos centímetros, y entré de lleno en la guardia corta. Mi codo derecho, reforzado con compresión de roca y deflagración de fuego interno, impactó directamente contra su mandíbula en un uppercut cortante de Muay Thai.",
        "El crujido óseo retumbó por encima del estruendo. La cabeza del cultista giró en un ángulo imposible y su cuerpo de cien kilos cayó desplomado como un saco de plomo.",
        "El segundo intentó embestirme con el hombro. Lo recibí con un bloqueo de antebrazo, trabé su cuello con ambas manos en un clinch de hierro y le descargué un rodillazo ascendente demoledor en mitad del esternón.",
        "El impacto de *Fa Jin* comprimió la caja torácica, destrozando costillas y perforando los pulmones del gigante, que escupió una bocanada de sangre negra antes de caer de rodillas.",
        "El tercero dudó medio segundo, aterrorizado. Di una patada circular baja con la tibia blindada que le segó la rodilla como una hoz, y rematé con un golpe de palma abierta al rostro que le hundió el tabique nasal en el cerebro.",
        "Tres gigantes de Khorne eliminados en menos de cinco segundos.",
        "* * *",
        "A mi derecha, Vance von Holst luchaba con una furia desesperada pero técnica. Un mercenario intentó clavarle una lanza corta; Vance desvió la punta con su estoque dorado en una parada de segunda perfecta y contragolpeó con una estocada certera que atravesó la garganta del atacante.",
        "—¡Por Solaria y por la verdadera nobleza! —gritó Vance con la respiración jadeante, mientras dos milicianos cubrían sus flancos con sonrisas de respeto marcial.",
        "Al ver a sus cultistas masacrados y a sus mercenarios atrapados en el cuello de botella de la empalizada, el Capitán Malrik entró en pánico absoluto.",
        "—¡Maldita sea! ¡Nos tendieron una trampa! —bramó, espoleando a su caballo hacia el camino del bosque mientras sujetaba el cofre de hierro bajo el brazo.",
        "No llegó lejos.",
        "Cargué con un sprint explosivo que levantó una estela de barro y salté sobre el lomo del corcel. Mi brazo atrapó a Malrik por el cuello de la cota de malla y lo arranqué de la silla en el aire, azotándolo brutalmente contra el suelo empapado.",
        "El cofre de hierro rodó por el fango, abriéndose con un chasquido metálico.",
        "Malrik intentó arrastrarse hacia su espada corta, pero mi bota se plantó pesadamente sobre su muñeca derecha, partiendo el radio con un crujido seco que le arrancó un aullido desgarrador.",
        "—Se acabó, capitán —dije, apoyando la punta de mi puño calcificado contra su sien.",
        "* * *",
        "Julian y Vlad llegaron corriendo a la carrera, rodeados por los milicianos que vitoreaban con lágrimas de victoria.",
        "Julian se arrodilló junto al cofre desparramado y recogió un pliego de pergamino grueso, empapado por los bordes pero intacto en su contenido.",
        "La luz de una antorcha reveló el texto despiadado.",
        "Era un decreto formal de requisa y entrega de sesenta súbditos de Sauceda para 'fines de pacificación rúnica en las minas', firmado de puño y letra por el **Conde Raymond de Morvath** y refrendado por **Valerius**, sellado con la inconfundible cera negra de la casa noble.",
        "—Aquí está... —susurró Julian con la voz temblorosa de indignación pura—. La firma del conde. El sello de cera negra. La orden explícita de entrega a la secta de sangre.",
        "Vlad tomó el pergamino con manos enguantadas, guardándolo en un cilindro de latón sellado herméticamente:",
        "—Esta es la sentencia de muerte de la Casa Morvath. Con esto en manos del Rey, no hay senado ni corte que pueda salvar a tu padre ni a tu hermano del cadalso.",
        "Miré alrededor de la plaza de Sauceda.",
        "Los cultistas de Khorne yacían muertos en el barro; los mercenarios supervivientes estaban atados y desarmados por la milicia cívica; y los campesinos comenzaban a salir del granero, llorando y abrazando a los soldados que habían salvado a sus hijos de una muerte atroz.",
        "Julian se puso de pie, mirando hacia las luces distantes del castillo en la cima del desfiladero con una resolución inquebrantable.",
        "—Sauceda está libre —dijo el joven noble—. Ahora solo queda marchar hacia el castillo... y purgar el trono podrido de mi padre.",
        "Apreté los nudillos manchados de barro y sangre.",
        "La Sección Uno tenía la prueba en sus manos.",
        "Y el amanecer que se aproximaba sobre Morvath traería el juicio final del acero."
    ]
}

def compilar_capitulo_69():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 69              ")
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
    from agregar_capitulo_47 import CAPITULO_47
    from agregar_capitulo_48 import CAPITULO_48
    from agregar_capitulo_49 import CAPITULO_49
    from agregar_capitulo_50 import CAPITULO_50
    from agregar_capitulo_51 import CAPITULO_51
    from agregar_capitulo_52 import CAPITULO_52
    from agregar_capitulo_53 import CAPITULO_53
    from agregar_capitulo_54 import CAPITULO_54
    from agregar_capitulo_55 import CAPITULO_55
    from agregar_capitulo_56 import CAPITULO_56
    from agregar_capitulo_57 import CAPITULO_57
    from agregar_capitulo_58 import CAPITULO_58
    from agregar_capitulo_59 import CAPITULO_59
    from agregar_capitulo_60 import CAPITULO_60
    from agregar_capitulo_61 import CAPITULO_61
    from agregar_capitulo_62 import CAPITULO_62
    from agregar_capitulo_63 import CAPITULO_63
    from agregar_capitulo_64 import CAPITULO_64
    from agregar_capitulo_65 import CAPITULO_65
    from agregar_capitulo_66 import CAPITULO_66
    from agregar_capitulo_67 import CAPITULO_67
    from agregar_capitulo_68 import CAPITULO_68
    from agregar_capitulo_69 import CAPITULO_69
    
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
        CAPITULO_45, CAPITULO_46, CAPITULO_47, CAPITULO_48, CAPITULO_49,
        CAPITULO_50, CAPITULO_51, CAPITULO_52, CAPITULO_53, CAPITULO_54,
        CAPITULO_55, CAPITULO_56, CAPITULO_57, CAPITULO_58, CAPITULO_59,
        CAPITULO_60, CAPITULO_61, CAPITULO_62, CAPITULO_63, CAPITULO_64,
        CAPITULO_65, CAPITULO_66, CAPITULO_67, CAPITULO_68, CAPITULO_69
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 69
    txt_cap69 = os.path.join(capitulos_dir, "69_Capitulo_69.txt")
    with open(txt_cap69, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_69['header'].upper()}: {CAPITULO_69['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_69['paragraphs']))
    print(f"[OK] Texto plano Capítulo 69 guardado en:\n  {txt_cap69}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 69
    docx_cap69 = os.path.join(capitulos_dir, "69_Capitulo_69.docx")
    create_single_chapter_document(CAPITULO_69, docx_cap69, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 69 guardado en:\n  {docx_cap69}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 69)
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
    compilar_capitulo_69()
