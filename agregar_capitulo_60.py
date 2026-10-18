# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_60 = {
    "header": "Capítulo 60",
    "title": "El Compás de Dos Latidos: Natsuki y Lucy",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "natsuki_besa_lucy.jpg"),
            "caption": "Ilustración: Natsuki y Lucy (Un Beso Dulce y Sincero Bajo el Ocaso)"
        }
    ],
    "paragraphs": [
        "El bullicio del almuerzo comunitario comenzó a dispersarse cuando el sol de la tarde inició su pausado declive hacia el mar.",
        "Mientras la mayoría de los cadetes de la F-2 se acomodaban en los salones de la villa para descansar de la agotadora jornada de voleibol o se dispersaban en pequeñas tertulias bajo los porches sombreados, **Natsuki** y **Lucy** se alejaron de las mesas comunales con pasos tranquilos.",
        "No hubo gritos histéricos ni órdenes teatrales esta vez.",
        "Lucy, caminando descalza sobre la madera tibia de la pasarela que conducía hacia el muelle de observación, lucía un suéter ligero de hilo color crema sobre su atuendo de playa y sus coletas castaño rojizas se mecían al compás de la brisa marina. Aunque conservaba ese porte orgulloso y decidido que la caracterizaba, la rigidez hosca de los primeros días se había suavizado notablemente.",
        "A su lado, Natsuki caminaba con las manos en los bolsillos de su sudadera oscura, sintiendo una agradable sensación de paz que contrastaba con el eterno estrés de los campos de batalla.",
        "—El mar se ve muy tranquilo desde aquí... —comentó Lucy en voz baja, deteniéndose junto a la barandilla de madera del mirador frente a las olas.",
        "—Sí. Es difícil creer que hace menos de una semana estábamos cubiertos de fango y piedra bruja defendiendo la fortaleza —respondió Natsuki con una media sonrisa reflexiva.",
        "Lucy ladeó la cabeza, mirándolo de reojo. Sus ojos dorados, que habitualmente lanzaban chispas de fuego defensivo, tenían ahora un brillo suave y curioso.",
        "—Oye, Natsuki... —murmuró ella, jugueteando con el extremo de una de sus coletas—. Durante el partido de voleibol... estuviste animando a todo el mundo como un loco. Te vi correr por toda la arena para llevarle agua a Seier y a Ken.",
        "—Bueno, no soy el más fuerte del aula, así que intento apoyar en lo que puedo —dijo él, rascándose la nuca con timidez.",
        "—Eres un tonto —repuso Lucy, pero en su tono no había aspereza, sino una calidez genuina que hizo latir el corazón de Natsuki—. No tienes que minimizarte siempre. En la trinchera del asedio te vi pelear. Tu aura plateada es increíble, y la forma en que cuidaste a los heridos... nadie en la F-2 piensa que eres débil. Yo ciertamente no lo creo.",
        "Natsuki se quedó mirándola, gratamente sorprendido. Las palabras de Lucy no venían envueltas en sarcasmo ni en reclamos; eran una muestra limpia y honesta de aprecio que le arrancó una sonrisa sincera.",
        "—Gracias, Lucy. Que me lo digas tú significa mucho para mí.",
        "La chica sintió un leve cosquilleo en las mejillas y desvió la mirada hacia el agua con un puchero suave, aunque sus labios delataban una sonrisa contenida.",
        "—N-no te acostumbres a los halagos. Solo digo las cosas como son.",
        "* * *",
        "Se sentaron al borde del muelle, con las piernas colgando sobre la marea mansa que rompía con un susurro blanco contra los pilares de piedra.",
        "Lucy sacó de su bolsillo un pequeño frasco de vidrio que contenía una chispa mágica de ignición. La destapó y, concentrando un hilo fino de maná en la punta de sus dedos índices, comenzó a moldear la pequeña llama para que formara la silueta de un diminuto jabalí de fuego que correteaba sobre la palma de su mano sin quemarla.",
        "—Cuando era pequeña en la capital —explicó Lucy, observando la figura incandescente con nostalgia serena—, los instructores de magia me decían que el fuego solo sirve para destruir y que mi temperamento descontrolaría mis reservas. Por eso me obsesioné con hacerlo explotar todo. Pensaba que si era destructiva, al menos nadie podría ignorarme.",
        "—Pero ahora controlas el fuego con muchísima delicadeza —observó Natsuki, fascinado por la destreza de la chica.",
        "—Porque en los Hijos de Calydon aprendí que el calor también cocina para los amigos y mantiene tibio el campamento en las noches heladas del páramo —respondió Lucy, cerrando el puño para apagar la figura y girándose hacia él—. Y... últimamente, cuando estoy contigo, mi maná no se siente caótico. Se siente... en calma.",
        "Natsuki extendió su mano derecha sobre la madera y, lentamente, canalizó una pizca de su **Aura de Batalla**.",
        "Un fulgor plateado y etéreo, suave como la luz de la luna, envolvió sus dedos. Con un gesto pausado, acercó su mano a la de Lucy. La chica no se apartó; al contrario, extendió sus dedos y dejó que ambas energías se rozaran.",
        "Las chispas cálidas de Lucy y la neblina plateada de Natsuki no chocaron en una explosión; se entrelazaron en una danza armoniosa, desprendiendo un calor reconfortante que les subió por los brazos hasta el pecho.",
        "Lucy soltó un suspiro suave, contemplando el entrelazado de sus manos con una emoción profunda.",
        "—Esto se siente muy bien... —susurró ella, alzando la mirada hacia él.",
        "—A mí también me hace sentir en paz, Lucy. Me gusta mucho estar a tu lado.",
        "Ninguno de los dos mencionó palabras grandilocuentes ni promesas apresuradas. Su historia apenas estaba dando sus primeros pasos tras el fragor de la guerra; no necesitaban juramentos eternos de amor para saber que lo que estaba naciendo entre ellos era algo verdadero, dulce y especial que ambos querían cuidar.",
        "* * *",
        "El cielo comenzó a teñirse de un naranja encendido y tonos púrpuras cuando el sol rozó la línea del horizonte marítimo.",
        "El viento fresco del atardecer hizo que Lucy temblara levemente por el contraste de temperatura. Natsuki, notando el gesto, no dudó: deslizó su brazo por detrás de sus hombros y la atrajo con delicadeza contra su costado para resguardarla de la brisa.",
        "Lucy apoyó su cabeza en el hombro de Natsuki con una naturalidad que días atrás habría parecido imposible. Sus manos se buscaron en el regazo, entrelazando sus dedos con una firmeza cariñosa.",
        "—Natsuki...",
        "—¿Sí?",
        "—A veces me cuesta decir las cosas sin ponerme a la defensiva —admitió ella con una sinceridad conmovedora, mirándolo a los ojos con sus mejillas teñidas de carmín bajo la luz dorada del ocaso—. Pero quiero que sepas que me alegra mucho haberte conocido en esta academia. Y... me gusta que seas tú quien me acompañe.",
        "Natsuki sintió que el pecho se le ensanchaba de dicha. Con su mano libre, acarició suavemente la mejilla de Lucy, apartándole un mechón rebelde de cabello.",
        "—A mí también me encanta acompañarte, Lucy. Cada día que pasa me gustas más.",
        "Lucy cerró los ojos por un instante, disfrutando de la caricia cálida de sus dedos. Cuando volvió a abrirlos, no había miedo ni distancia entre ellos.",
        "Natsuki se inclinó despacio hacia adelante. Lucy alzó la barbilla, sosteniéndole la mirada con los labios entreabiertos en un suspiro entrecortado. Apoyó ambas manos con timidez sobre el pecho de Natsuki, sintiendo el latido firme y acelerado de su corazón.",
        "Bajo el resplandor majestuoso del atardecer sobre la Costa Esmeralda, sus labios se encontraron en un beso tierno, pausado y profundamente dulce.",
        "No fue un choque impulsivo de batalla, sino un contacto sincero y demorado donde el calor de sus almas se reconoció mutuamente. Lucy se estremeció de pies a cabeza, cerrando los ojos con fuerza y profundizando el beso con una entrega inocente que le aceleró el pulso a mil por hora.",
        "* * *",
        "Cuando finalmente se separaron para tomar aire, sus frentes quedaron unidas con suavidad.",
        "Lucy respiraba con agitación, con las mejillas intensamente sonrojadas y una sonrisa radiante, hermosa y desarmada que iluminaba su rostro entero.",
        "—Besas bastante bien para ser un novato... —susurró ella con una mirada pícara y afectuosa, dándole un golpecito suave con la frente.",
        "—Tengo una maestra muy linda que me motiva —respondió Natsuki con una sonrisa cómplice.",
        "—Tonto... —murmuró Lucy escondiendo el rostro contra el cuello de Natsuki, aspirando su aroma mientras rodeaba su torso con ambos brazos en un abrazo apretado y dulce.",
        "Permanecieron así durante largos minutos, mientras el sol desaparecía por completo y las primeras estrellas de la noche comenzaban a titilar sobre el manto índigo del océano.",
        "No había prisa. El verano continuaba su curso en la Costa Esmeralda, y paso a paso, latido a latido, dos corazones que habían conocido la soledad descubrían juntos la magia de empezar a enamorarse."
    ]
}

def compilar_capitulo_60():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 60              ")
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
        CAPITULO_60
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 60
    txt_cap60 = os.path.join(capitulos_dir, "60_Capitulo_60.txt")
    with open(txt_cap60, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_60['header'].upper()}: {CAPITULO_60['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_60['paragraphs']))
    print(f"[OK] Texto plano Capítulo 60 guardado en:\n  {txt_cap60}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 60
    docx_cap60 = os.path.join(capitulos_dir, "60_Capitulo_60.docx")
    create_single_chapter_document(CAPITULO_60, docx_cap60, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 60 guardado en:\n  {docx_cap60}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 60)
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
    compilar_capitulo_60()
