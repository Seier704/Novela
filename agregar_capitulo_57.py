# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document
from agregar_capitulo_56 import CAPITULO_56

CAPITULO_57 = {
    "header": "Capítulo 57",
    "title": "Mareas de Verano: Sandías, Sonrojos y Corazones Entrelazados",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "natsuki_lucy_playa.png"),
            "caption": "Ilustración: Natsuki y Lucy (Complicidad, Sandía y Sonrojos en la Orilla)"
        }
    ],
    "paragraphs": [
        "El tercer día de vacaciones en la Costa Esmeralda amaneció con un cielo de azul purísimo y una brisa marina que mitigaba con gracia el calor estival.",
        "Tras la mágica noche compartida en la suite de **Lady Elysea de Vane**, me levanté temprano con el cuerpo lleno de vigor y el alma en una calma inquebrantable.",
        "Cuando bajé hacia la playa privada de la villa, el espectáculo que encontré era una estampa viva de vitalidad, descanso y camaradería que borraba cualquier rastro del asedio Skaven y las penurias de los calabozos.",
        "Elysea bajó poco después luciendo un elegante vestido veraniego de lino color cielo, con el cabello plateado trenzado a un costado; sus mejillas aún conservaban un rubor delicioso que delataba la intensidad de nuestras horas previas, pero su porte aristocrático se había suavizado con una calidez luminosa. Al cruzar miradas conmigo, me dedicó una sonrisa tímida y cómplice que jamás le habría visto en los pasillos de la Academia.",
        "—Buenos días, mi caballero —me saludó con gracia al pasar a mi lado, rozando discretamente sus dedos con los míos.",
        "Bajo la gran carpa central, **Kaoruko** me recibió con un cuenco colmado de frutas frescas y rodajas de piña escarchada, depositando un beso cariñoso en mi mejilla sin el menor atisbo de recelo, mientras **Lortelle**, reclinada en una reposera de mimbre con una copa de vino dorado en mano y gafas de sol ahumadas, nos observaba con su habitual sonrisa pícara y calculadora.",
        "—Vaya, Silva... pareces más descansado que nunca —bromeó la heredera de la Compañía Elte con su tono burlón y seductor—. Aunque me atrevería a decir que las sábanas de seda del ala norte tuvieron una noche bastante ajetreada.",
        "Elysea desvió la mirada con un sonrojo instantáneo que le llegó hasta las puntas de sus orejas élficas, mientras Kaoruko soltaba una risita dulce y cómplice. La armonía entre mis tres compañeras no solo era real, sino que se había convertido en un pacto tácito de felicidad compartida.",
        "* * *",
        "A lo largo de la costa, la Clase F-2 y nuestros invitados disfrutaban de una jornada rebosante de anécdotas cotidianas y romanticismo desinhibido.",
        "En la zona norte de la playa, **Erza Scarlet** había decidido que la construcción de castillos de arena era un ejercicio táctico de ingeniería militar. Con una seriedad marcial que contrastaba de forma hilarante con su traje de baño escarlata, daba órdenes estrictas a **Renji Abarai** y a **Caesar King** sobre el grosor que debían tener las murallas de contención frente a la marea.",
        "—¡Abarai, ese foso tiene una pendiente deficiente! ¡El oleaje romperá el baluarte este en tres minutos si no compactas la arena con maná! —exclamaba Erza con los brazos en jarras.",
        "—¡Oye, Erza, que esto se supone que son vacaciones, no una prueba de asedio del Mariscal Gerald! —se quejaba Renji con la pala de madera en la mano, cubierto de arena hasta los hombros, desatando las carcajadas estruendosas de Caesar, quien moldeaba con orgullo un colosal jabalí de arena gigante en la cabecera del fuerte.",
        "Un poco más allá, **Rukia Kuchiki** observaba la escena sentada bajo una sombrilla de paja con una caracola en el regazo. Renji, decidido a impresionarla tras su regaño militar, recogió un coco gigante de cáscara durísima de una palmera y se acercó a ella con el pecho inflado:",
        "—¡Mira esto, Rukia! Los hombres de verdad no necesitan cuchillos para abrir un coco —fanfarroneó, golpeando la fruta con el canto de su mano.",
        "El coco ni se inmutó, y Renji contuvo una mueca de dolor que le hizo apretar los dientes. Rukia suspiró con una sonrisa divertida y, rozando con delicadeza la superficie con la punta de su dedo imbuido en el frío de *Sode no Shirayuki*, congeló una pequeña fisura en la cáscara. Renji asestó entonces un golpe seco partiéndolo limpiamente en dos mitades rebosantes de agua fresca.",
        "—Buen trabajo, fortachón —le dijo Rukia con una mirada dulce, tomando una mitad. Al ofrecerle la otra, sus dedos se rozaron durante varios segundos; ambos apartaron la vista sonrojados hasta la coronilla, bebiendo al unísono con torpe ternura.",
        "A la orilla del agua, la marea acariciaba los pies descalzos de **Vlad** y **Zemina**. Caminaban tomados de la mano, recolectando caracoles de nácar y pequeños cristales marinos. Vlad, despojado de cualquier rigidez noble del Ducado, rodeaba los hombros de Zemina con su brazo musculoso y depositaba besos espontáneos en su frente cada vez que una ola rompía cerca de sus rodillas, haciéndola reír con una felicidad cristalina.",
        "En las aguas cristalinas de los arrecifes, **Burnice** chapoteaba alegremente salpicando a **Ken**, quien desplegaba pequeñas plumas de fuego solar sobre la superficie del agua haciendo que el vapor silbara entre burbujas; luego la alzaba en vilo sobre sus hombros mientras la chica rubia se sujetaba de su cabello riendo a viva voz.",
        "Bajo una carpa de lona blanca más retirada, **Lighter** le acomodaba con infinita paciencia la visera a **Hoshina**, acercándole una copa de nieve escarchada con jarabe de bayas mientras conversaban en voz baja, con Hoshina apoyando suavemente su cabeza en el hombro del boxeador.",
        "Y cerca de las rocas, **Akame** y **Pulchra** emergían del mar cargando dos gigantescos cangrejos reales y varios pargos plateados recién capturados con sus dagas, listas para que **Yami Sukehiro** —quien fumaba su habitual cigarrillo recostado contra una palmera— las felicitara con un gruñido perezoso de aprobación.",
        "* * *",
        "Sin embargo, la escena más entrañable de la tarde se desarrollaba bajo una sombrilla de palmeras a unos metros del agua.",
        "Allí estaba **Lucy**.",
        "La pequeña artillera de los Hijos de Calydon estaba sentada con las piernas cruzadas sobre una toalla de rayas, luciendo un llamativo conjunto deportivo de playa de dos piezas en tonos rojo y negro. Tenía sus características coletas rubias sujetas a ambos lados y los brazos cruzados bajo su pecho, inflando los mofletes en un mohín de enfado infantil mientras miraba con desdén hacia el océano.",
        "—¡Hmph! ¡El mar es una estupidez! ¡Hay demasiada sal, demasiada arena pegajosa y hace un calor insoportable! ¡No entiendo por qué todos actúan como si fuera el paraíso! —bufaba entre dientes, pateando un puñado de arena.",
        "A su lado, con una calma infinita y una sonrisa paciente grabada en el rostro, se encontraba **Subaru Natsuki**.",
        "Llevaba una camisa hawaiana abierta sobre una camiseta blanca, bermudas oscuras y el cabello peinado hacia atrás. Lejos de fastidiarse por los reproches constantes de la chica, Natsuki colocó sobre la toalla un vaso alto de limonada recién exprimida con hielo picado y una rodaja de limón.",
        "Luego, con una reverencia cómica y una calidez que desarmaba cualquier guardia, le extendió una gruesa tajada de sandía roja, fresca y jugosa.",
        "—Para una experta del páramo que come polvo de motor todos los días, un poco de sol no debería ser tanto drama, señorita artillera —bromeó Natsuki con voz suave—.",
        "Lucy se sobresaltó, girando la cabeza hacia él como un felino acorralado.",
        "—¡¿Qué crees que haces, idiota?! —espetó con las mejillas encendiéndose en un rojo furioso—. ¡Yo no te pedí nada! ¡Puedo soportar el calor perfectamente sola! ¡No necesito que un cadete novato venga a consentirme!",
        "—Ya lo sé, Lucy —respondió Natsuki sin borrar su sonrisa amable, sentándose a su lado con las piernas cruzadas sobre la toalla—. Sé que eres fuerte y que derribas monstruos con tu bate de fuego sin pestañear. Pero incluso la mejor artillera de Calydon necesita refrescarse antes de que se le recaliente el motor. Vamos, tómala. La corté yo mismo con las provisiones que preparó Kaoruko.",
        "Lucy miró la tajada de sandía, luego el vaso de limonada escarchada, y finalmente los ojos sinceros de Natsuki.",
        "Sus labios temblaron ligeramente; la máscara de niña mimada y matona del páramo se resquebrajó por completo. Con un bufido derrotado y un puchero irresistiblemente tierno, arrebató la fruta con brusquedad:",
        "—¡Solo la acepto porque tengo sed, no porque me lo hayas pedido tú! —rezongó.",
        "Le dio un gran mordisco a la fruta. Sus ojos dorados se abrieron de par en par al sentir el dulzor frío y refrescante recorriendo su garganta, dejando escapar un suspiro involuntario de satisfacción:",
        "—¡Mmh...! —Lucy se cubrió la boca de inmediato, dándose cuenta de su reacción, y miró a Natsuki con el ceño fruncido pero con la cara totalmente encendida en carmín—. ¡Está... pasable! ¡Solo pasable!",
        "Natsuki soltó una carcajada limpia y alegre.",
        "—Me alegra que te guste.",
        "—Oye... Natsuki —dijo Lucy de pronto, bajando la voz y jugando nerviosamente con la pulsera de cuero de su muñeca—. ¿Por qué eres así conmigo? En los campamentos del páramo, la gente siempre me tenía miedo o se largaba cuando empezaba a gritar. Me decían que era insoportable y consentida.",
        "Natsuki miró hacia el mar con un brillo nostálgico y maduro en la mirada.",
        "—En mi ciudad natal... conocí a alguien muy parecida a ti —confesó en voz baja—. Alguien bajita, orgullosa y con un temperamento explosivo que siempre gritaba y exigía cosas imposibles para que nadie notara lo sola que se sentía. Aprendí que cuando alguien grita tan fuerte, a menudo es porque teme que, si se calla, nadie se dé cuenta de que está allí. Pero tú no estás sola, Lucy. Tienes a Caesar, a Piper, a los muchachos... y ahora también me tienes a mí. Detrás de tus berrinches, sé que tienes un corazón de oro y una lealtad inmensa.",
        "Las palabras de Natsuki cayeron sobre Lucy como un bálsamo cálido e inesperado.",
        "La pequeña artillera abrió los ojos conmovida, sintiendo un nudo en la garganta. Sus defensas se desplomaron por completo; bajó la mirada hacia la arena, balbuceando con timidez:",
        "—Eres... un idiota sentimental, Natsuki... —susurró, pero en su tono ya no había pizca de enojo, sino un agradecimiento profundo y una ternura naciente.",
        "Natsuki vio una pequeña gota de jugo de sandía en la comisura de los labios de Lucy. Con delicadeza natural, tomó una servilleta y la limpió suavemente.",
        "Al sentir el roce de sus dedos, Lucy se estremeció de pies a cabeza, volviéndose tan roja como la pulpa de la sandía. Sin saber qué hacer con sus manos, tomó el vaso de limonada y se lo extendió a Natsuki con timidez:",
        "—Toma... b-bebe tú también un poco... no quiero que te desmayes por insolación y digan que fue mi culpa...",
        "Natsuki sonrió de oreja a oreja, tomó un sorbo del vaso que ella le ofrecía y asintió:",
        "—Deliciosa. Gracias, Lucy.",
        "A unos metros de distancia, detrás de una sombrilla, **Piper Wheel** y **Caesar King** observaban la escena con una amplia sonrisa de complicidad fraternal. Piper le dio un codazo suave a la líder de los Hijos de Calydon:",
        "—¿Viste eso, jefa? —dijo la mecánica con tono de abuela orgullosa—. Parece que ese muchacho de la F-2 sabe exactamente cómo aflojarle las tuercas a nuestra pequeña cascarrabias.",
        "—Ese chico tiene agallas y un corazón enorme —asintió Caesar, cruzándose de brazos con orgullo—. Me alegra que Lucy haya encontrado a alguien que la entienda así.",
        "* * *",
        "Observé toda la escena desde mi tumbona, con un vaso de néctar frío en la mano.",
        "Kaoruko descansaba su cabeza en mi hombro, Elysea me servía un poco más de bebida con una mirada cómplice y Lortelle sonreía satisfecha contemplando la bahía.",
        "El sol comenzaba a teñir las olas de un dorado deslumbrante.",
        "La Clase F-2 ya no era únicamente una unidad militar invencible forjada en la sangre y el acero del asedio; en estas playas de la Costa Esmeralda, entre risas compartidas, castillos de arena, rodajas de sandía y amores consolidados, nos habíamos convertido en una verdadera familia.",
        "Y supe, con una certeza absoluta que resonaba en cada músculo de mi cuerpo, que protegería la paz y la felicidad de cada uno de ellos hasta mi último aliento."
    ]
}

def compilar_capitulo_57():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 57              ")
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
        CAPITULO_55, CAPITULO_56, CAPITULO_57
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 57
    txt_cap57 = os.path.join(capitulos_dir, "57_Capitulo_57.txt")
    with open(txt_cap57, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_57['header'].upper()}: {CAPITULO_57['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_57['paragraphs']))
    print(f"[OK] Texto plano Capítulo 57 guardado en:\n  {txt_cap57}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 57
    docx_cap57 = os.path.join(capitulos_dir, "57_Capitulo_57.docx")
    create_single_chapter_document(CAPITULO_57, docx_cap57, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 57 guardado en:\n  {docx_cap57}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 57)
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
    compilar_capitulo_57()
