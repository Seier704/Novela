# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_66 = {
    "header": "Capítulo 66",
    "title": "Caminos Grises y Sombras de Morvath: El Feudo del Olvido",
    "paragraphs": [
        "El mapa táctico que nos entregó el Capitán Vaneck marcaba una ruta de más de diez jornadas hacia el extremo oriental del Imperio de Solaria.",
        "A medida que la caravana militar de la Sección Uno se alejaba de la capital, la pulcritud de los caminos adoquinados y los campos fértiles fue cediendo el paso a un paisaje desolado y frío: ciénagas cenagosas, bosques de robles marchitos cubiertos de hongos negruzcos y puestos de guardia imperiales abandonados con sus torres carcomidas por el fuego y el tiempo.",
        "El destino oficial de nuestra primera misión de larga duración era el **Condado de Morvath**, un feudo fronterizo aislado del que hacía meses apenas llegaban tributos ni informes regulares a la corte.",
        "Sentado sobre el pescante del carromato de provisiones, vigilaba el camino con la vista clavada en la niebla que reptaba sobre el fango.",
        "—Huele a podrido desde hace tres leguas —soltó Kaelan, agazapado en el techo de madera del convoy con sus orejas de lince apuntando hacia los matorrales—. Y no hablo de animales muertos. Es humo de turba, cuero empapado y sangre rancia.",
        "—Tu olfato no se equivoca, chico gato —masculló Borgar, golpeando el mango de su martillo contra una rueda de hierro—. Miren ese puente de ahí adelante. Las dovelas están cuarteadas y las vigas podridas. El señor de estas tierras no ha gastado una sola moneda de cobre en reparar sus caminos en diez años. ¡Es una vergüenza para cualquiera que sepa tallar piedra!",
        "Faelar bajó de su montura con un salto elástico, examinando las cortezas de los árboles junto al sendero:",
        "—No solo es negligencia, enano. Miren las ramas altas. Hay marcas de corte ritual y trampas de alambre herrumbroso. Los cazadores de esta zona no buscan venados; buscan fugitivos.",
        "—¿F-fugitivos? —tartamudeó Vance von Holst, cabalgando muy cerca de Vlad con el rostro pálido y el capote carmesí salpicado de barro—. P-pero el instructor dijo que esto era solo una patrulla de inspección de rutina... ¿Por qué el camino parece un cementerio sin tumbas?",
        "Vlad tiró de las riendas de su corcel negro, mirándolo con frialdad marcial:",
        "—Bienvenido a la realidad del imperio, Holst. Lejos de los bailes de salón de Solaria, las leyes las dictan el hambre y el hierro.",
        "* * *",
        "Al cruzar un recodo del camino flanqueado por sauces deshojados, la advertencia de Vlad se materializó con crudeza.",
        "De una rama gruesa colgaban tres cadáveres esqueléticos balanceados por el viento. Sus ropas eran harapos campesinos y alrededor de sus cuellos pendían tablillas de madera carcomida con una sola palabra grabada a fuego: *'MOROSIDAD'*.",
        "Akame detuvo el paso, levantando la vista hacia los cuerpos con sus ojos rojos carmesí inmóviles:",
        "—Llevan muertos al menos un mes. Murieron por asfixia lenta. Ninguna bestia los tocó porque están cubiertos de cal viva para espantar carroñeros.",
        "—Una ejecución pública feudal —dije con la mandíbula apretada, sintiendo la rabia hervir en mis nudillos—. El conde ahorcó a sus propios labradores para que los demás paguen el diezmo sin rechistar.",
        "Vance desvió la mirada y vomitó sobre los helechos al borde del fango. Kaelan siseó enseñando los colmillos y Borgar escupió al suelo con repugnancia.",
        "—Sigamos avanzando —ordené, aflojando los hombros y canalizando un pulso sutil de maná telúrico a través de las suelas de mis botas para palpar las vibraciones del terreno—. En este lugar nadie va a recibirnos con flores.",
        "* * *",
        "Al caer la tarde, las murallas decrépitas de **Morvath** surgieron entre la bruma densa.",
        "La ciudad feudal era una mole de piedra oscura encajada en un desfiladero rocoso, flanqueada por dos asentamientos satélites miserables: al oeste, **Sauceda**, un pueblo de chozas de adobe donde los campesinos se amontonaban con rostros famélicos; y al este, **Pedregal Alto**, una aldea minera en las faldas del monte de donde salían lentas columnas de humo grisáceo.",
        "Frente al rastrillo levadizo de Morvath, un grupo de diez mercenarios con cotas de malla oxidadas y alabardas desgastadas nos cerró el paso.",
        "—¡Alto ahí, forasteros! —gritó el capitán de la guardia local, un hombre tuerto con aliento a alcohol barato—. El paso a la ciudad cuesta cinco monedas de plata por cabeza. Impuesto de peaje del Conde Raymond.",
        "Vlad avanzó a caballo un paso por delante de la formación. Desenvainó un tercio de su pesada espada bastarda y alzó el medallón con el sello de la Orden del León y el emblema de la Academia Imperial:",
        "—Somos cadetes de la Gran Academia de Astraea bajo mandato directo de la Corona. Si intentas cobrarle tributo a un destacamento militar imperial, te atravesaré la garganta antes de que toques tu campana, guardia.",
        "El tuerto vio el blasón ducal de Vlad, miró los noventa kilos de músculo compacto de Seier y la mano de Akame sobre *Murasame*, y tragó saliva retrocediendo con el rostro lívido:",
        "—¡A-abran el portón! ¡Son los soldados de la capital!",
        "* * *",
        "Cruzamos las calles de Morvath entre charcos de aguas negras, niños descalzos con vientres hinchados por la desnutrición y miradas de terror absoluto de los pobladores, que cerraban sus postigos apenas nos veían pasar.",
        "En la cima de la colina, contrastando grotescamente con la miseria circundante, se alzaba el castillo condal: una fortaleza de mármol decadente, iluminada con cientos de antorchas y protegida por verjas de hierro macizo.",
        "Fuimos conducidos al gran salón de recepciones.",
        "Allí, repantingado sobre un trono de terciopelo manchado de licor, nos esperaba el **Conde Raymond de Morvath**.",
        "Era un hombre gordo y rubicundo, de mejillas fofas y ojos acuosos cargados de indolencia, rodeado de copas de plata y sirvientas de miradas aterrorizadas.",
        "—Vaya, vaya... —soltó el conde con una risa pastosa, tamborileando sus dedos enjoyados sobre el brazo del trono—. El Senado de Astraea me prometió una división de veteranos de guerra y me envía a un puñado de mocosos de diecisiete años, a un enano y a una bestia lince. Supongo que la capital tiene cosas más importantes que atender.",
        "A su diestra, de pie con una espada de empuñadura de oro y una capa de armiño, estaba su primogénito, **Valerius de Morvath**.",
        "El joven noble, de mirada altiva, labios finos y una sonrisa llena de desprecio, nos barrió de arriba abajo:",
        "—¿Estos son los héroes de la Roca Carmesí? Parecen perros callejeros recién salidos de una zanja. Espero que al menos sepan empuñar una lanza contra los bandidos de Sauceda, plebeyos.",
        "Borgar apretó el martillo hasta que sus nudillos crujieron y Kaelan tensó los hombros, listo para saltar.",
        "Di un paso al frente antes de que la sangre llegara al río. Mi voz sonó grave, densa y calmada en medio de la sala:",
        "—Nosotros acabamos con un Señor de la Guerra Skaven de cuatro metros y con cuatrocientas ratas acorazadas, Lord Valerius. Si tus bandidos de bosque son más duros que eso, con gusto te dejaremos a ti ir primero al frente.",
        "Valerius abrió los ojos de par en par, ofendido hasta la médula, y dio un paso hacia su espada. Pero antes de que pudiera desenvainar, una voz firme y pausada intervino desde las sombras de una columna lateral:",
        "—Basta, Valerius. No avergüences más a la Casa Morvath.",
        "* * *",
        "El que avanzó hacia la luz era el segundo hijo del conde: **Julian de Morvath**.",
        "A diferencia de su hermano mayor, Julian no vestía terciopelo ni joyas; llevaba una cota de cuero reforzada, botas de jinete embarradas y ojeras profundas en un rostro joven pero templado por la vigilia y la sensatez. A su espalda, sosteniendo tímidamente su mano, caminaba su hermano menor, **Leo**, un niño de no más de ocho años que miraba a Seier y a Akame con ojos asustadizos.",
        "—Disculpen a mi padre y a mi hermano —dijo Julian con una reverencia respetuosa hacia Vlad y hacia mí—. Soy Julian. Y a diferencia de ellos, yo sé exactamente por qué están aquí... y el infierno en el que acaban de meterse.",
        "El conde Raymond agitó la mano con desdén:",
        "—Bah, tonterías de muchacho paranoico...",
        "Julian ignoró a su padre y clavó sus ojos oscuros directamente en los míos:",
        "—En los dos pueblos, Sauceda y Pedregal, no solo hay bandidos, Seier Silva. Hay cosechas envenenadas, guardias desaparecidos cada luna llena y una secta oscura operando en las minas abandonadas que mi familia se niega a reconocer. Si vinieron buscando una misión de entrenamiento fácil, los engañaron.",
        "Akame soltó la saya de su espada y Vlad asintió con gravedad.",
        "Miré a Julian, luego al niño que temblaba tras sus piernas, y finalmente al conde decadente y a su primogénito.",
        "La misión de la Sección Uno acababa de comenzar. Y la podredumbre de Morvath iba a requerir mucho más que órdenes de academia para ser purgada."
    ]
}

def compilar_capitulo_66():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 66              ")
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
        CAPITULO_65, CAPITULO_66
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 66
    txt_cap66 = os.path.join(capitulos_dir, "66_Capitulo_66.txt")
    with open(txt_cap66, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_66['header'].upper()}: {CAPITULO_66['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_66['paragraphs']))
    print(f"[OK] Texto plano Capítulo 66 guardado en:\n  {txt_cap66}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 66
    docx_cap66 = os.path.join(capitulos_dir, "66_Capitulo_66.docx")
    create_single_chapter_document(CAPITULO_66, docx_cap66, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 66 guardado en:\n  {docx_cap66}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 66)
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
    compilar_capitulo_66()
