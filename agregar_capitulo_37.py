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

CAPITULO_37 = {
    "header": "Capítulo 37",
    "title": "Caída en el Abismo",
    "illustration": os.path.join(NOVELA_DIR, "imagenes", "lortelle_kecheln_herida.png"),
    "illustration_caption": "Ilustración: Lortelle Kecheln (Caída en el Abismo)",
    "paragraphs": [
        "El colapso de la galería subterránea parecía haber quedado atrás, pero las entrañas de la tierra rara vez perdonan la violencia de una explosión.",
        "Apenas habíamos comenzado a reorganizar las filas de los cadetes heridos en el corredor superior, cuando un crujido sordo, profundo y telúrico sacudió los cimientos de la montaña bajo nuestras botas. No fue un temblor ordinario de maná; la onda expansiva del puente detonado por mi impacto de tierra y las galerías carcomidas por los miles de túneles de los Skavens habían provocado una falla estructural masiva.",
        "El suelo de basalto se desgarró en dos en un abrir y cerrar de ojos.",
        "—¡CUIDADO! ¡LA ROCA SE ESTÁ HUNDIENDO! —gritó Erza a través del polvo.",
        "Una grieta gigantesca, una fisura negra y sin fondo, se abrió en mitad del pasadizo como las fauces de un leviatán hambriento.",
        "La mayoría de los guerreros de la F-2 reaccionaron por puro instinto de supervivencia dando un salto hacia las paredes firmes. Sin embargo, entre el caos y la densa polvareda, una cadete de la Clase C que había quedado rezagada perdió el equilibrio. Las losas de piedra bajo sus pies se desmoronaron al vacío y la muchacha resbaló hacia la oscuridad soltando un grito ahogado de terror.",
        "No lo pensé.",
        "Mi cuerpo se movió antes de que mi cerebro pudiera evaluar el peligro. Exploté hacia el borde del abismo con un salto de Muay Thai, lanzándome de cabeza hacia la negrura de la fosa para alcanzarla.",
        "—¡¡¡SEIER-KUN!!! —el grito desgarrador de Kaoruko resonó en las alturas de la caverna, perdiéndose en el rugido de la roca quebrada.",
        "* * *",
        "La caída libre fue un descenso vertiginoso en la oscuridad más absoluta.",
        "El viento subterráneo aullaba en mis oídos mientras fragmentos afilados de basalto y estalactitas caían junto a nosotros como una lluvia de metralla. En mitad del vacío, alcancé la silueta de la chica, rodeé su cintura con mi brazo derecho y la atraje con fuerza contra mi pecho.",
        "Era menuda y ligera. Sentí su corazón latiendo desbocado como un pájaro atrapado entre mis brazos. Sin perder una décima de segundo, encorvé mi espalda, crucé mis antebrazos reforzados alrededor de su cabeza y envolví su cuerpo menudito con mis noventa kilos de músculo, convirtiéndome en un escudo humano.",
        "Canalicé todo mi maná de **Tierra** hacia mi piel y mi columna vertebral para densificar mi masa corporal.",
        "Golpeamos contra un saliente de roca en pendiente: *¡KRAAAK!* El impacto me arrancó un gemido sordo cuando varias piedras me laceraron la espalda y los hombros, pero apreté los dientes y no aflojé el agarre ni un milímetro. Rebotamos a través de una densa cortina de raíces subterráneas, cascadas de agua gélida y taludes de fango resbaladizo, hasta que finalmente nos estrellamos contra el lecho pedregoso de un río subterráneo.",
        "El golpe final contra la grava mojada fue brutal.",
        "El aire fue expulsado con violencia de mis pulmones, la cabeza me dio vueltas en una marea de luces rojas y la inconsciencia me tragó por completo.",
        "* * *",
        "El sonido constante de gotas de agua cayendo sobre la piedra fue lo primero que me devolvió a la vigilia.",
        "Plip... plip... plip...",
        "Abrí los ojos con lentitud, parpadeando para despejar la vista. Un dolor sordo me recorría los omóplatos y las costillas, pero mi respiración era estable y mis huesos no estaban fracturados gracias al blindaje de maná que había sostenido durante la caída.",
        "El aire a nuestro alrededor era frío, espeso y húmedo, saturado de olor a musgo y agua mineral profunda. Nos encontrábamos en una caverna gigantesca y silenciosa, iluminada únicamente por vetas de cristales fosforescentes azulados que crecían en las paredes rocosas.",
        "Estábamos muy, muy por debajo de los estratos cartografiados por la academia.",
        "A mi lado, tendida sobre la arena húmeda de la orilla, yacía la chica.",
        "Tenía dieciséis años, como todos nosotros. Una cabellera desordenada de un tono castaño rojizo caía sobre sus hombros empapados, enmarcando un rostro fino y aristocrático manchado de hollín y sangre seca. Su uniforme de la Clase C estaba rasgado en las mangas y en la falda, y su respiración era rápida y entrecortada.",
        "Me incorporé apretando los dientes por las punzadas en la espalda y me arrodillé junto a ella.",
        "—Oye... despierta. ¿Puedes oírme? —le hablé con mi voz grave, apoyando con delicadeza dos dedos en su cuello para comprobar su pulso.",
        "La muchacha emitió un pequeño quejido doloroso y abrió los ojos.",
        "Eran dos gemas de un verde esmeralda hipnótico y penetrante. A pesar de los moretones en su pómulo y del labio partido, no había llanto histérico ni pánico ciego en su mirada; al contrario, una chispa de astucia y fascinación calculadora se encendió en sus pupilas en cuanto me reconoció.",
        "—Vaya... parece que sobreviví a un viaje directo al infierno —susurró con una voz melodiosa y ligeramente ronca, dejando escapar una sonrisa traviesa y fascinante que desafiaba por completo su estado herido.",
        "Intentó incorporarse sobre los codos, pero un espasmo agudo le hizo soltar un quejido entre dientes y volver a caer sobre la arena, sujetándose el tobillo derecho.",
        "—Quédate quieta —le ordené con calma firme pero cuidadosa—. Tienes el tobillo torcido y probablemente una fisura en las costillas por la caída.",
        "La chica parpadeó, contemplando mi torso desnudo cubierto de raspones y los vendajes de hierro en mis brazos, antes de mirar las piedras del techo por donde habíamos caído.",
        "—Tú... te arrojaste al vacío para atraparme —dijo ella, ladeando la cabeza con esa media sonrisa irresistible y desconcertante—. Sentí cómo tu cuerpo absorbía todos los golpes contra la roca. Te usaste a ti mismo como amortiguador para que una extraña no se hiciera pedazos contra el suelo... ¿Es que acaso eres el héroe más idiota o el más temerario de toda Solaria?",
        "—No suelo dejar morir a la gente si puedo evitarlo —respondí con una media sonrisa tranquila, sacando una tira limpia de tela de mi cartuchera para inmovilizarle el tobillo—. Soy Seier Silva, de la Clase F-2.",
        "La chica entrecerró sus ojos esmeralda con una chispa de deleite.",
        "—Sé perfectamente quién eres, Silva. El plebeyo que noquea nobles de un puñetazo y lideró a los marginados a la gloria del torneo —la pelirroja extendió su mano delgada, estrechando la mía con una firmeza sorprendente para una joven herida—. Mi nombre es **Lortelle Kecheln**, de la prestigiosa Compañía Mercantil Elte y cadete de la Clase C.",
        "Lortelle soltó una pequeña risita cristalina, a pesar del dolor en sus costillas.",
        "—En el mundo del comercio, Silva, todo favor tiene un precio y toda deuda se salda con intereses. Pero me parece que salvarme la vida de una caída de cien metros es una deuda que difícilmente podré pagarte con monedas de oro.",
        "—No te preocupes por deudas ahora, Lortelle —le dije, ajustando la venda sobre su tobillo con firmeza precisa—. Guárdate las monedas para cuando salgamos de este agujero.",
        "* * *",
        "Examiné nuestro entorno.",
        "El río subterráneo fluía hacia un cañón oscuro y angosto donde el agua se perdía en las profundidades. No había rastro de cuerdas, ni de gritos, ni de señales de rescate. El derrumbe que Dante y yo habíamos provocado arriba probablemente había sellado la entrada del estrato, y los Skavens seguían infestando las galerías superiores.",
        "A unos metros de la orilla, encontré una pequeña gruta seca bajo un saliente de granito, protegida de la humedad y del viento frío.",
        "Regresé hacia Lortelle, me agaché a su lado y pasé mis brazos con cuidado por debajo de sus rodillas y su espalda.",
        "—Sujétate a mi cuello. Te voy a mover a un lugar seco.",
        "Lortelle abrió un poco los ojos por la sorpresa, pero no protestó. Enredó sus brazos suaves alrededor de mi cuello con una naturalidad felina, apoyando su rostro tibio contra mi hombro mientras la levantaba en vilo con facilidad. A pesar de los golpes y la fatiga, el cuerpo de la chica desprendía un aroma dulce y sutil a canela y especias finas.",
        "La deposité con delicadeza sobre un lecho de musgo seco dentro de la pequeña gruta. Me quité mi chaqueta de combate impermeable y la cubrí con ella para resguardarla del frío.",
        "Lortelle se acurrucó bajo la prenda, inhalando el calor residual de mi cuerpo con una mirada que combinaba agradecimiento y una intriga intensa.",
        "—Estamos en territorio no cartografiado, ¿verdad? —preguntó ella en voz baja, mirando la oscuridad del río.",
        "—Muy por debajo del Segundo Estrato —asentí, sentándome en el borde de la gruta para vigilar el túnel—. Pero no te preocupes.",
        "Me giré hacia ella y le sostuve la mirada con una calma inquebrantable.",
        "—Te aseguro que te sacaré de aquí con vida, Lortelle. Aunque tenga que abrirme paso a puñetazos a través de la montaña entera, te protegeré.",
        "Lortelle se quedó congelada durante dos segundos eternos.",
        "La máscara calculadora de comerciante que siempre llevaba puesta se derritió por completo ante mis palabras. Una calidez profunda le encendió las mejillas y una sonrisa hermosa, tímida y cómplice se dibujó en sus labios mientras me miraba con sus ojos esmeralda brillando en la penumbra.",
        "—Entonces confiaré en ti, mi protector —susurró con una dulzura inédita en su voz—. No me decepciones, Seier.",
        "En lo profundo del túnel subterráneo, un eco lejano anunció la presencia de sombras en movimiento.",
        "La verdadera lucha por la supervivencia en las entrañas del abismo acababa de comenzar."
    ]
}

def compilar_capitulo_37():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 37              ")
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
        CAPITULO_35, CAPITULO_36, CAPITULO_37
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 37
    txt_cap37 = os.path.join(capitulos_dir, "37_Capitulo_37.txt")
    with open(txt_cap37, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_37['header'].upper()}: {CAPITULO_37['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_37['paragraphs']))
    print(f"[OK] Texto plano Capítulo 37 guardado en:\n  {txt_cap37}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 37
    docx_cap37 = os.path.join(capitulos_dir, "37_Capitulo_37.docx")
    create_single_chapter_document(CAPITULO_37, docx_cap37, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 37 guardado en:\n  {docx_cap37}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 37)
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
    compilar_capitulo_37()
