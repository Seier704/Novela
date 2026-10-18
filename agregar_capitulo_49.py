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

CAPITULO_49 = {
    "header": "Capítulo 49",
    "title": "Olas Doradas, Rubor y Nuevos Lazos",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "zemina.png"),
            "caption": "Ilustración: Zemina (La Invitada Especial de Vlad)"
        }
    ],
    "paragraphs": [
        "El inicio formal del receso de verano desató una euforia incontenible en los pabellones de la Academia.",
        "Atrás quedaban las marchas forzadas con cotas de malla, las clases teóricas de Vaneck y el recuerdo amargo del miasma Skaven. Sin embargo, antes de que cada quien tomara rumbos dispersos por el imperio, **Lortelle Kecheln** tomó las riendas de la situación con la astucia que la caracterizaba.",
        "—Escuchen todos —declaró Lortelle aquella misma mañana en la plaza de los dormitorios, apoyando una mano en su cintura con una sonrisa cargada de seguridad y picardía—. Sé que muchos tenían planes familiares o querían encerrarse a entrenar. Pero la **Compañía Mercantil Elte** posee un resort privado en la **Costa Esmeralda del Sur**: arena dorada, aguas cristalinas, pabellones de descanso frente al mar y los mejores chefs de mariscos del continente. He reservado el complejo entero para nosotros.",
        "La invitación no era solo para Seier: Lortelle extendió el convite a **toda la Clase F-2 en pleno** —Kaoruko, Ken, Yami, Akame, Erza, Hoshina, Natsuki, **Renji Abarai**, **Rukia Kuchiki** y los seis miembros de los Hijos de Calydon: Caesar King, Lighter, Piper Wheel, Lucy, Burnice White y Pulchra— y a **Lady Elysea de Vane**, quien aceptó con una dignidad aristocrática que apenas disimulaba sus celos y su determinación de no ceder terreno.",
        "Pero la comitiva no se detuvo ahí.",
        "Aprovechando su presencia en los jardines de la Academia tras la condecoración, Lortelle se dirigió hacia el recién ascendido Capitán Táctico de Campaña.",
        "—Capitán Vlad —dijo la joven mercader con elegancia—. Tu liderazgo en las murallas salvó las vidas de muchos de los nuestros. Como socia comercial y compañera de armas, sería un honor que nos acompañaras a la costa.",
        "Vlad parpadeó sorprendido. El joven noble de la Clase B, habituado a la rigidez castrense del Ducado, contempló la algarabía de la F-2 con una mezcla de pudor y sincera gratitud.",
        "—Agradezco profundamente la invitación, señorita Kecheln... y a todos ustedes —respondió Vlad con una leve inclinación de cabeza—. Me encantaría aceptar. Sin embargo... me gustaría preguntar si la invitación permite un acompañante adicional.",
        "El silencio curioso se hizo en el círculo. Ken arqueó las cejas con una sonrisa pícara, Renji soltó una carcajada burlona dándole un codazo a Ken, y Lighter soltó un silbido bajo.",
        "—¿Un acompañante, Vlad? —preguntó Caesar King cruzándose de brazos con curiosidad fraternal—. ¿De quién se trata?",
        "Vlad se aclaró la garganta, sintiendo cómo un ligero tinte carmesí asomaba por encima de su cuello almidonado:",
        "—Es mi mejor amiga de la infancia... **Zemina**. Ha estado preocupada por mí durante todo el asedio y las semanas de campaña, esperándome en la capital. Me gustaría mucho que pudiera descansar con nosotros.",
        "Me adelanté medio paso, apoyando una mano firme sobre el hombro de Vlad.",
        "—Cualquiera que sea un hermano o hermana leal para ti, Vlad, tiene un lugar en nuestra mesa —le aseguré con mi tono grave habitual—. Dile a Zemina que prepare su equipaje.",
        "Vlad sonrió con un alivio genuino en la mirada, asintiendo con fuerza.",
        "* * *",
        "La presentación oficial de **Zemina** tuvo lugar a la mañana siguiente frente a los carruajes de lujo de la Compañía Elte.",
        "Era una muchacha de una belleza singular y magnética: su abundante melena rojiza y cobriza caía en ondas suaves sobre sus hombros, enmarcando un rostro sereno con ojos cálidos y un salpicado adorable de pecas sobre el puente de su nariz. Vestía una blusa blanca de lino con mangas abullonadas y un corsé de cuero marrón entallado que resaltaba su cintura esbelta sobre una falda escarlata.",
        "Al mirar a Vlad, un brillo de ternura evidente y un afecto cómplice iluminaron sus pupilas.",
        "—Buenos días a todos... Mi nombre es Zemina —se presentó con una reverencia suave y una sonrisa tímida pero desenvuelta—. Vlad no ha parado de hablarme de la valentía de la Clase F-2 y de la destreza del señor Silva. Es un verdadero placer.",
        "Kaoruko fue la primera en recibirla con un abrazo cálido y un trozo de pastel recién horneado, rompiendo el hielo al instante. En cuestión de minutos, Zemina ya estaba integrada en la comitiva, compartiendo risas y anécdotas del viaje junto a Hoshina, Burnice y Rukia, mientras Renji y Natsuki ayudaban a cargar los baúles más pesados en el pescante trasero.",
        "* * *",
        "El trayecto de dos días hacia el sur en los carruajes acolchados de la Compañía Elte fue un festival de camaradería.",
        "Cuando el camino de la cordillera desembocó finalmente en la línea costera, un murmullo de asombro colectivo llenó las ventanillas.",
        "La Costa Esmeralda hacía honor a su nombre: una franja interminable de arena blanca y fina bordeada por aguas turquesas de una transparencia hipnótica, bajo un cielo azul sin una sola nube. Las olas rompían con una mansedumbre relajante contra la orilla y, sobre una colina baja cubierta de palmeras tropicales, se alzaba la imponente villa privada de Lortelle, construida con maderas nobles, terrazas abiertas y balcones con vista directa al mar.",
        "—¡¡¡EL MAAAAAAR!!! —bramaron Burnice y Lucy al unísono, saltando de los estribos antes de que los caballos se detuvieran por completo.",
        "—¡A cambiarse todo el mundo! —ordenó Caesar con una carcajada—. ¡El último en tocar el agua paga la cena!",
        "* * *",
        "El momento de salir a la playa con los trajes de baño fue un acontecimiento que paralizó el aire de la costa.",
        "Los muchachos salimos primero hacia la arena dorada.",
        "A pesar de que el solsticio invitaba al relajo, semanas enteras de condicionamiento espartano, supervivencia en calabozos y combates a muerte habían dejado su huella en nuestros cuerpos.",
        "Con mis noventa kilos de masa muscular compacta en bañador oscuro, mi torso era un bloque macizo de granito: dorsales anchos, hombros redondeados como rocas de río y el abdomen tallado por el acondicionamiento neuromuscular del Muay Thai y el Jeet Kune Do. Las pequeñas cicatrices de combate en mis costillas y brazos no afeaban mi figura; le daban el porte temible y magnético de un guerrero que vive en la vanguardia.",
        "A mi lado, **Ken** lucía un físico fibroso y atlético, con los pectorales y deltoides esculpidos por el entrenamiento acrobático de Hwa Young y su piel bronceada por el maná solar.",
        "**Lighter** exhibía la silueta clásica y estética de un boxeador de campeonato: cintura estrecha, oblicuos marcados y brazos fornidos listos para percutir combinaciones de velocidad.",
        "**Renji Abarai** caminaba con arrogante soltura, luciendo una imponente musculatura de espadachín de asalto: hombros anchos, pectorales densos y los exóticos tatuajes tribales que serpenteaban por su cuello y clavículas hacia su espalda definida, con su larga melena carmesí atada en una coleta alta.",
        "**Yami**, con sus dieciséis años y su descaro habitual, caminaba descalzo con bermudas oscuras mostrando una complexión ancha, pesada y sorprendentemente musculosa, mientras **Vlad** ostentaba la figura esbelta, elegante y definida de un espadachín ducal, y **Natsuki** sorprendía a todos con una espalda notablemente más firme y tonificada que el día que pisó la Academia.",
        "* * *",
        "Y entonces... salieron las chicas de los vestuarios de la villa.",
        "Un silencio repentino cayó sobre la orilla, acompañado por el chasquido suave de las olas.",
        "La belleza femenina desplegada sobre la arena dorada cortaba la respiración.",
        "Al frente avanzaba **Kaoruko Waguri**, luciendo un traje de baño de dos piezas de tono lavanda claro. Su figura de reloj de arena, con curvas generosas, una cintura diminuta y caderas suaves, contrastaba de forma embriagadora con la blancura de su piel. Al notar mi mirada fija sobre ella, sus mejillas se encendieron en un rubor escarlata delicioso; intentó taparse con timidez antes de sonreírme con una devoción dulce y deslumbrante.",
        "A su diestra caminaba **Lady Elysea de Vane**.",
        "La elfa aristocrática vestía un bikini azul zafiro con ribetes plateados que acentuaba sus piernas largas y estilizadas, su abdomen plano y la elegancia glacial de su escote. Intentó mantener su porte regio y distante, pero el sonrojo en sus orejas puntiagudas delataba lo consciente que estaba de los ojos que la contemplaban.",
        "**Lortelle Kecheln**, en cambio, no mostró timidez alguna.",
        "Llevaba un traje de baño esmeralda con cintas doradas que dejaba al descubierto su espalda curvilínea y la redondez perfecta de sus caderas. Caminaba sobre la arena con un bamboleo sensual y deliberado, guiñándome un ojo mientras sostenía un cóctel de frutas tropicales en la mano:",
        "—¿Te gusta lo que ves, Silva? —susurró al pasar a mi lado, rozando deliberadamente su hombro tibio contra mi brazo.",
        "Detrás de ellas, **Caesar King** deslumbraba con un bañador deportivo que resaltaba su altura de amazona y la fuerza magnética de sus muslos torneados; **Lucy y Piper** lucían conjuntos alegres y coquetos; **Burnice** saltaba en la arena con un bikini naranja brillante; **Pulchra** se desperezaba bajo el sol con la agilidad grácil de una felina; **Erza Scarlet** imponía su presencia escultural en un bañador carmesí digno de una reina; y **Akame** devoraba una rodaja de sandía fresca con un traje negro simple pero sumamente ceñido.",
        "**Rukia Kuchiki** descendía hacia la arena vistiendo un hermoso traje de baño de dos piezas en tono celeste pastel y blanco nieve. Aunque su figura era menuda y delicada en comparación con la voluptuosidad de las demás, poseía una gracia aristocrática y pura: cintura esbelta, piel inmaculada de porcelana y unas piernas gráciles que atraparon instantáneamente la atención de la orilla.",
        "**Hoshina**, con un bañador verde agua discreto pero muy favorecedor, se ajustaba una cinta en el cabello mientras sus ojos se cruzaban con los de Lighter, haciendo que ambos apartaran la mirada sonrojados casi al unísono.",
        "Y junto al grupo, **Zemina** era una visión encantadora.",
        "Su traje de baño en tonos terracota y blanco resaltaba sus pecas, su cintura firme y su cabello cobrizo ondeando al viento salino. Al encontrarse con la mirada de Vlad, la chica bajó la cabeza con una sonrisa dulce y coqueta que dejó al joven capitán de la Clase B completamente petrificado y con el rostro rojo como un tomate maduro.",
        "* * *",
        "El rubor entre ambos bandos fue generalizado.",
        "Los latidos de los corazones se aceleraron al ver los cuerpos descubiertos bajo la luz del mediodía.",
        "Sin embargo, lo que definió aquel instante no fue la vulgaridad, sino el respeto inquebrantable que unía a nuestra hermandad.",
        "Los muchachos éramos guerreros forjados en la lealtad y el honor. Apreciamos la belleza indiscutible de nuestras camaradas sin miradas lascivas ni comentarios soeces; cada uno mantuvo sus ojos anclados con respeto y adoración únicamente en quienes hacían latir su pecho.",
        "Ken clavó su sonrisa descarada en Burnice, quien le arrojó una toalla mojada riendo a carcajadas. Lighter se acercó a Hoshina para ofrecerle una sombrilla, tartamudeando levemente como nunca antes en el cuadrilátero. Vlad caminó con torpeza entrañable hacia Zemina para ofrecerle su mano sobre las piedras de la orilla, mientras los dos compartían susurros que hablaban de un cariño que pronto dejaría de ser solo amistad.",
        "A pocos metros de allí, **Renji Abarai**, con el rostro encendido en un rubor tan ardiente como su propio cabello rojo, se atragantó con su propia saliva al contemplar la figura delicada de **Rukia**; apresuradamente apartó la vista frotándose la nuca con torpeza y fingiendo examinar las nubes para no ser indiscreto. Rukia, percatándose del sonrojo del fogoso espadachín, dejó escapar una risita suave y se acercó a ofrecerle una fruta que acababa de escarchar con una brisa de su maná gélido, tiñendo sus propias mejillas de un rosa tenue.",
        "Y en el centro de la playa, me vi rodeado por tres de las mujeres más codiciadas y poderosas del imperio.",
        "Kaoruko tomó mi mano izquierda entrelazando sus dedos con los míos, Elysea se situó a mi derecha cruzando miradas desafiantes con la comerciante, y Lortelle apoyó su codo sobre mi hombro con una risita triunfal.",
        "—Parece que el agua está tibia, Seier —dijo Kaoruko con un brillo travieso en sus ojos azabache.",
        "No respondí con palabras.",
        "Pasé mi brazo por la cintura de Kaoruko y la levanté en vilo sobre mi hombro como una pluma, corriendo hacia la espuma de las olas mientras ella soltaba un chillido de risa y sorpresa.",
        "—¡¡SEIER, IDIOTA, ESPERA!! —gritó Lortelle corriendo detrás de nosotros con los pies descalzos levantando arena, seguida de cerca por una Elysea que por primera vez en meses soltó una carcajada abierta y juvenil.",
        "El agua fresca nos envolvió en un estallido de espuma blanca y sol brillante.",
        "Las vacaciones de verano habían comenzado.",
        "Y el verano prometía arder más que el fuego de mil batallas."
    ]
}

def compilar_capitulo_49():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 49              ")
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
        CAPITULO_45, CAPITULO_46, CAPITULO_47, CAPITULO_48, CAPITULO_49
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 49
    txt_cap49 = os.path.join(capitulos_dir, "49_Capitulo_49.txt")
    with open(txt_cap49, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_49['header'].upper()}: {CAPITULO_49['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_49['paragraphs']))
    print(f"[OK] Texto plano Capítulo 49 guardado en:\n  {txt_cap49}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 49
    docx_cap49 = os.path.join(capitulos_dir, "49_Capitulo_49.docx")
    create_single_chapter_document(CAPITULO_49, docx_cap49, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 49 guardado en:\n  {docx_cap49}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 49)
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
    compilar_capitulo_49()
