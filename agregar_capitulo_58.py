# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_58 = {
    "header": "Capítulo 58",
    "title": "Chispas al Atardecer: El Secreto de la Pequeña Fiera",
    "paragraphs": [
        "El sol de la tarde comenzaba a inclinarse sobre las aguas turquesas de la Costa Esmeralda, tiñendo el horizonte de una mezcla deslumbrante de tonos ámbar, violeta y carmesí.",
        "En la zona principal de la playa, la euforia del mediodía había dado paso a una calma reconfortante. Ken y Burnice descansaban recostados en las tumbonas tras horas de nadar en los arrecifes; Caesar King y Piper Wheel revisaban un juego de cartas náuticas junto a la mesa del porche; y a lo lejos, el murmullo de las olas marcaba un compás hipnótico que arrullaba el campamento.",
        "Sin embargo, bajo la sombrilla del ala este, la atmósfera seguía cargada de una vibración eléctrica.",
        "Lucy, con las mejillas todavía ardiendo en un tono escarlata brillante tras haber compartido su vaso de limonada con Natsuki, no soportó un solo segundo más las miradas burlonas que Piper y Caesar le lanzaban a la distancia.",
        "La pequeña artillera de la Clase F-2 se puso de pie de golpe, sacudiéndose la arena del bañador con movimientos bruscos y bufando con una indignación que no engañaba a nadie:",
        "—¡Uf, qué calor tan sofocante hace aquí! ¡Es inaguantable! —declaró con los brazos en jarras y la barbilla alzada, fingiendo que la brisa marina no era la más fresca de la temporada—. Voy a caminar hacia los arrecifes de la caleta este... ¡y tú vienes conmigo, Natsuki!",
        "Natsuki parpadeó, sorprendido por la orden repentina pero divertido por el puchero adorable que coronaba los labios de la chica.",
        "—¿Yo? Creí que estabas descansando a la sombra.",
        "—¡No discutas! —espetó Lucy, señalándolo con un dedo acusador mientras sus orejitas se encendían de rubor—. ¡Alguien tiene que vigilar que las rocas no tengan erizos venenosos y cargar mi sombrilla si me canso! ¿O es que vas a dejar que una dama camine sola por terreno rocoso, so vago?",
        "Natsuki soltó una risita suave y se levantó sin la menor queja, sacudiéndose el pantalón corto.",
        "—A sus órdenes, señorita Lucy. Seré su escolta personal.",
        "—¡Y nada de 'señorita'! ¡Suenas como los mayordomos estirados de la capital! —gruñó ella dándose media vuelta con paso enérgico, aunque una media sonrisa de triunfo se dibujaba en su rostro mientras sus tres pequeños jabalíes mágicos de fuego salían disparados detrás de ella, trotando felices sobre la arena húmeda.",
        "* * *",
        "Ambos se alejaron del bullicio de la villa, dejando atrás las sombrillas y las tumbonas hasta alcanzar la ensenada oriental.",
        "Allí, la línea de arena se transformaba en una hermosa caleta de piedra arenisca y arrecifes de coral. Las mareas bajas habían formado pequeñas pozas naturales de agua cristalina donde nadaban pececillos de colores y caracolas nacaradas que brillaban bajo el resplandor del crepúsculo.",
        "El rumor de la fiesta comunitaria quedó amortiguado por la distancia, dejando únicamente el susurro envolvente de la brisa marina.",
        "Los pequeños jabalíes de Lucy chapoteaban en los charcos, resoplando minúsculas chispas de vapor caliente al olisquear las conchas, mientras Natsuki y Lucy caminaban uno al lado del otro en un silencio que poco a poco perdió toda incomodidad.",
        "Se detuvieron junto a una saliente de roca lisa y pulida por el oleaje, lo bastante ancha para que dos personas se sentaran a contemplar la inmensidad del mar.",
        "Lucy se sentó primero, balanceando sus piernas desnudas sobre el borde. Con la mirada fija en el agua dorada, abrazó sus rodillas contra el pecho, perdiendo por completo la postura agresiva que siempre exhibía en público.",
        "Natsuki se acomodó a su lado, dejando un espacio respetuoso entre ambos.",
        "—Oye... —murmuró Lucy en un tono tan bajo y suave que apenas superó el murmullo de la marea—. Lo que dijiste antes... sobre esa persona en tu mundo anterior... ¿era verdad?",
        "Natsuki contempló el horizonte, donde las gaviotas volaban en círculos buscando su refugio nocturno.",
        "—Totalmente verdad —respondió él con serenidad—. Era una chica que siempre fruncía el ceño, exigía cosas imposibles y trataba a todos con frialdad. Al principio me molestaba, pero con el tiempo comprendí que detrás de esa coraza había alguien aterrorizada de mostrarse vulnerable. Creía que si dejaba de ser perfecta y demandante, nadie se preocuparía por ella.",
        "Lucy tragó saliva. Sus pequeños dedos se aferraron con fuerza a la tela de su pareo.",
        "—Yo... yo no siempre fui parte de los Hijos de Calydon —confesó ella, y por primera vez, su voz de dieciséis años tembló con una herida antigua—. En la capital, mi familia era una de las ramas mercantiles más adineradas del distrito este. Mi padre quería que fuera una dama sumisa, que me casara con un noble viejo por interés comercial y que jamás levantara la voz.",
        "La chica soltó un bufido amargo, apretando los dientes.",
        "—Pero yo nací con un maná de fuego inestable y un bate de béisbol en las manos. La primera vez que accidentalmente quemé el carruaje de un pretendiente porque me faltó al respeto, mi propio padre me llamó monstruo y me desterró al Páramo Exterior como si fuera basura inservible.",
        "Natsuki la miró en silencio, conmovido por la crudeza de su relato.",
        "—En el páramo aprendí que si eres blanda, te devoran —continuó Lucy, con los ojos empañados en un brillo cristalino que se esforzaba por contener—. Si no gritaba más fuerte que nadie, si no era la más explosiva, la más ruidosa y la más temida, volverían a romperme. Por eso cuando Caesar y Piper me acogieron... juré que jamás volvería a ser débil. Pero a veces... a veces tengo miedo de ser solo una carga chillona que cansa a todos.",
        "* * *",
        "Natsuki no dudó un solo segundo.",
        "Deslizó su mano derecha sobre la roca tibia y cubrió con delicadeza la pequeña mano temblorosa de Lucy.",
        "La chica se sobresaltó por el contacto, pero lejos de apartarse, sus dedos se crisparon sobre la palma de Natsuki buscando un ancla en mitad de la tormenta de sus recuerdos.",
        "—Mírame, Lucy —le pidió él con suavidad.",
        "Ella levantó despacio la mirada. Sus ojos ámbar estaban húmedos y sus labios entreabiertos temblaban con una fragilidad hermosa y desgarradora.",
        "—En el asedio a la fortaleza, cuando los Skavens tenían a los rehenes y todo parecía perdido, tú y los muchachos cayeron del cielo como un vendaval —le dijo Natsuki con una firmeza absoluta—. Te vi blandir ese bate envuelta en llamas, quemando el fango para proteger a los indefensos sin importarte el peligro. No vi a una niña consentida ni a una carga... vi a una guerrera feroz, con un corazón gigantesco y una lealtad que no tiene precio.",
        "Natsuki entrelazó sus dedos con los de ella, dejando que un tenue y reconfortante hilo de su **Aura de Batalla** plateada envolviera sus manos, armonizando en un baile cálido con las diminutas chispas rojas que brotaban de la piel de Lucy.",
        "—Tu fuego no es un monstruo, Lucy. Es el fuego que protege a tu familia. Y tú no necesitas gritar ni fingir fiereza para que yo te preste atención. Yo te veo exactamente como eres... y me encantas así.",
        "El aliento de Lucy se quebró.",
        "Las palabras de Natsuki derribaron el último muro de su fortaleza interior. Dos lágrimas rodaron por sus mejillas coloradas, pero no eran de dolor; eran del alivio más dulce y abrumador que jamás había experimentado.",
        "En ese instante, una ola más alta de la marea ascendió por la roca, salpicando agua fría sobre sus tobillos. Lucy soltó un grito ahogado de sorpresa y perdió el equilibrio hacia atrás.",
        "Natsuki reaccionó con los reflejos sobrehumanos templados en las batallas del calabozo: pasó su brazo libre por la cintura de Lucy y la atrajo con firmeza hacia su torso para evitar que cayera a la poza profunda.",
        "Quedaron atrapados en un abrazo íntimo, a escasos centímetros de distancia.",
        "El pecho menudo de Lucy subía y bajaba con respiraciones entrecortadas contra el torso de Natsuki. El calor de sus cuerpos se fundió bajo el manto dorado del atardecer. Los ojos ámbar de la chica miraron fijamente los labios de Natsuki y, por primera vez en su vida, sus puños no se cerraron para golpear, sino que se aferraron con fuerza a la camisa del chico.",
        "—Natsuki... idiota... —susurró ella con un hilo de voz temblorosa.",
        "—Dime, Lucy.",
        "—Bésame ya o te juro que te incendio...",
        "Natsuki sonrió con una ternura infinita, acortó la distancia que los separaba y unió sus labios con los de ella.",
        "* * *",
        "Fue un beso dulce, torpe al inicio, pero cargado de una pasión pura y chispeante que les aceleró el corazón como un tambor de guerra.",
        "Los labios de Lucy sabían a azúcar de sandía y salitre fresco, cálidos como una brasa de hogar. Lucy soltó un suspiro ahogado contra su boca, rodeando el cuello de Natsuki con ambos brazos y profundizando el contacto con una entrega total que borró de golpe años de soledad y desconfianza.",
        "A su alrededor, los tres pequeños jabalíes mágicos saltaban y correteaban en círculos sobre el agua, emitiendo chillidos alegres y disparando chispitas pirotécnicas hacia el cielo del crepúsculo.",
        "Cuando finalmente se separaron para tomar aire, la frente de Lucy quedó apoyada contra la de Natsuki.",
        "La chica estaba roja hasta las puntas de las orejas, pero una sonrisa radiante, la más hermosa y desarmada que jamás había mostrado, iluminaba su rostro pecoso.",
        "—Escúchame bien, Natsuki Subaru... —murmuró ella con un puchero adorable mientras escondía la cara en su pecho—. Si te atreves a contarle esto a Lighter o a las chicas... te juro que te hago picadillo.",
        "—Mi boca está sellada, señorita fiera —bromeó Natsuki, acariciándole la coronilla con suavidad.",
        "—Y ni se te ocurra soltarme la mano en el camino de vuelta —añadió Lucy, apretando sus pequeños dedos entre los de él con posesividad tierna—. Ahora eres mío.",
        "—No pensaba soltarte jamás.",
        "* * *",
        "Caminaron de regreso hacia la villa bajo las primeras estrellas que comenzaban a salpicar el firmamento nocturno.",
        "Iban de la mano, con los jabalíes trotando mansamente detrás de sus talones, compartiendo miradas de complicidad y sonrisas cómplices que sellaban el inicio de una historia hermosa.",
        "Desde la barandilla de la terraza de la villa, **Caesar King** y **Piper Wheel** los vieron aparecer por el sendero de la playa. Caesar soltó un largo suspiro de satisfacción, sonriendo con orgullo maternal mientras brindaba al aire con su copa de vino.",
        "—Te lo dije, Piper —comentó la líder de Calydon—. Ese chico de la F-2 era exactamente lo que nuestra pequeña necesitaba.",
        "A su lado, Seier Silva, con Kaoruko apoyada en su hombro y Lortelle sonriendo a su costado, contempló a la nueva pareja con la satisfacción del líder que ve a sus hermanos de armas encontrar la paz y la felicidad que tanto les costó ganar.",
        "La noche caía sobre la Costa Esmeralda.",
        "Y entre el fuego de los amores compartidos y la lealtad inquebrantable de la F-2, el verano se convertía en el refugio eterno de nuestras almas."
    ]
}

def compilar_capitulo_58():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 58              ")
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
        CAPITULO_55, CAPITULO_56, CAPITULO_57, CAPITULO_58
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 58
    txt_cap58 = os.path.join(capitulos_dir, "58_Capitulo_58.txt")
    with open(txt_cap58, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_58['header'].upper()}: {CAPITULO_58['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_58['paragraphs']))
    print(f"[OK] Texto plano Capítulo 58 guardado en:\n  {txt_cap58}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 58
    docx_cap58 = os.path.join(capitulos_dir, "58_Capitulo_58.docx")
    create_single_chapter_document(CAPITULO_58, docx_cap58, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 58 guardado en:\n  {docx_cap58}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 58)
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
    compilar_capitulo_58()
