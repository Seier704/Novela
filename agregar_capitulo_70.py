# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_70 = {
    "header": "Capítulo 70",
    "title": "Ecos de Herejía y el Vuelo del Espíritu: La Guerra Silenciosa de Morvath",
    "paragraphs": [
        "El humo acre de las antorchas derribadas y el vapor caliente de la sangre espesaban el aire de la plaza de Sauceda.",
        "Los vítores de los aldeanos comenzaban a apagar el estruendo del combate. El Capitán Malrik yacía encadenado en el lodo con la muñeca destrozada, mientras dos milicianos custodiaban al único cabecilla cultista que había sobrevivido a la arremetida: un fanático de torso marcado con quemaduras rituales y una máscara de bronce brujo rota por la mitad.",
        "Vlad y yo nos acercamos al prisionero. El noble de la Orden del León apoyó la punta de su espada bastarda contra la garganta del hereje:",
        "—Habla ahora si quieres que tu muerte sea rápida —le exigió Vlad con voz de trueno—. ¿Cuántos son en las minas y dónde está el portal?",
        "El cultista no mostró terror. Al contrario, una risa gorgoteante y desquiciada comenzó a brotar de sus labios ensangrentados, deformando sus facciones en una mueca de fanatismo demente.",
        "—¡Creen que ganaron una batalla, cachorros del emperador...! —graznó el hombre, mientras sus ojos se teñían de un negro azabache absoluto y sus venas se hinchaban como gusanos oscuros bajo la piel—. ¡Sauceda era solo un señuelo...! ¡Aún faltan cientos de nuestros hermanos cultistas en las sombras...! ¡Están dentro de sus murallas, en sus sótanos y bajo sus calles...! ¡Morvath arderá en sangre para Khorne y nadie quedará vivo para contarlo!",
        "—¡Cuidado! —gritó Akame con su instinto de asesina encendiéndose en una milésima de segundo—. ¡Va a detonar!",
        "Reaccioné por puro reflejo: agarré a Julian del cuello del jubón y derribé a Vance hacia el fango detrás de una barricada de piedra caliza, mientras Vlad interponía su escudo de asedio blindado con maná.",
        "El cuerpo del fanático se convulsionó como un fuelle sobrecargado. Un brillo carmesí oscuro desgarró su piel desde adentro y estalló en una deflagración violenta de fuego negro, miasma hirviente y metralla ósea.",
        "La onda expansiva sacudió las paredes de adobe y arrojó cenizas ardientes sobre la plaza. Cuando el humo se disipó, en el suelo solo quedaba un cráter calcinado y una runa de ocho puntas humeante grabada en el barro.",
        "* * *",
        "Nos pusimos de pie, sacudiéndonos la ceniza hedionda de las armaduras.",
        "El ambiente triunfal de Sauceda se enfrió de golpe, reemplazado por una pesadez sombría.",
        "—Maldición demoníaca de autodestrucción —masculló Borgar, escupiendo al suelo mientras revisaba las muescas de su martillo—. Estos bastardos prefieren volar en pedazos antes que soltar información. Y por las barbas de mis ancestros, lo que dijo no era un farol.",
        "—Tiene razón el enano —añadió Kaelan, olfateando el aire con las orejas gachas—. Siento más rastros de ese hedor a azufre. No solo vienen de las minas altas; hay tufos similares dispersos por los bosques y cerca de las murallas de la ciudad.",
        "Faelar bajó del campanario del molino y se reunió con nosotros, con el rostro pálido y la mirada analítica:",
        "—Si tienen células durmientes dentro de la propia Morvath, esto no es una simple purga de bandidos. Es el inicio de una guerra prolongada de estrategia, supervivencia y resiliencia.",
        "Vlad miró el cilindro de latón donde descansaba el decreto sellado con cera negra del Conde Raymond:",
        "—No podemos enfrentarnos a un ejército subterráneo del Caos y a la guarnición completa del conde solo con siete cadetes y treinta milicianos. Necesitamos la intervención directa de la Corona.",
        "* * *",
        "Nos refugiamos en el interior del molino para trazar el curso de acción de inmediato.",
        "Sobre una mesa enharinada, Vlad y yo redactamos un informe de emergencia con terminología militar estricta: detallamos el hallazgo del Culto a Khorne, los sacrificios de niños y mujeres, la emboscada repelida en Sauceda y adjuntamos una copia transcrita de la orden de entrega firmada por el Conde Raymond y Valerius, avalada con el sello de plata de Julian.",
        "Al pie del pergamino, Vlad grabó su sello ducal junto a mi estampa de Caballero Defensor:",
        "—Solicitamos el envío urgente de al menos **cien caballeros nobles** de la Orden Real —dijo Vlad, lacrando el mensaje con cera carmesí—. Una fuerza montada pesada de esa magnitud es lo que se requiere para tomar el control de la ciudad, desarmar a los mercenarios y sentenciar al conde bajo la ley marcial del Rey.",
        "Pero la capital imperial estaba a más de diez jornadas de distancia a caballo. Un jinete común jamás llegaría a tiempo.",
        "—Déjamelo a mí —intervino Faelar, dando un paso al frente.",
        "El elfo silvano extrajo de su bolsa de cuero una pequeña talla de madera de sauce sagrado y una pluma de halcón plateada. Se arrodilló en el suelo, cerró los ojos y comenzó a entonar una salmodia suave en la lengua de los bosques antiguos, liberando hilos de maná de viento que hicieron girar las aspas del molino.",
        "De la madera tallada brotó un destello de luz esmeralda y plateada que tomó forma corpórea:",
        "Un **Espíritu de Ave Mensajera** —un halcón etéreo de plumaje translúcido y ojos de jade brillante— se posó sobre el antebrazo de Faelar, batiendo sus alas con un susurro que no perturbaba el polvo.",
        "Vlad ató el cilindro de latón a la pata del espíritu:",
        "—Vuela sin descanso hasta el palacio real de Solaria. Entrega este mensaje en las manos del Mariscal o del propio Monarca.",
        "El halcón espiritual soltó un graznido cristalino, se elevó hacia el cielo nocturno y, con una aceleración sónica que desafiaba las leyes físicas, surcó las nubes rumbo al oeste dejando una estela de polvo de estrellas que se desvaneció en el horizonte.",
        "* * *",
        "Julian observó el cielo con una mezcla de esperanza y angustia contenida:",
        "—El ave tardará al menos un día en llegar, y los cien caballeros tardarán cinco o seis jornadas más en cabalgar a marcha forzada hasta Morvath. ¿Qué hacemos mientras tanto? Si nos quedamos aquí sentados, mi hermano Valerius notará que Malrik no regresó, matará a mi hermano Leo y quemará la ciudad antes de escapar con el oro.",
        "Di un paso al frente, cruzando miradas con cada uno de mis compañeros: Vlad, Akame, Borgar, Faelar, Kaelan y Vance.",
        "—Julian tiene razón. La espera pasiva es una sentencia de muerte —declaré con voz firme y resolutiva—. La Corona enviará los refuerzos para la batalla final, pero nosotros no vamos a escondernos en este granero.",
        "—¿Cuál es la orden, Seier? —preguntó Akame, apoyando la mano en la saya de *Murasame*.",
        "—Vamos a limpiar el feudo desde adentro hacia afuera —dictaminé, señalando el mapa táctico—. Primero: aseguramos Sauceda como nuestra base logística fortificada, dejando al Sargento Roderick y a diez milicianos custodiando a las familias. Segundo: nos infiltramos de regreso en Morvath antes de que amanezca, tomamos la guarnición baja por sorpresa y desarmamos a los mercenarios de las puertas. Y tercero: barremos los alrededores y sellamos las entradas secundarias de las minas de Khorne.",
        "Vance von Holst alzó su estoque con un brillo de determinación en los ojos:",
        "—Cortar los tentáculos del monstruo antes de decapitarlo. Cuenten conmigo.",
        "—¡Aplastaremos a cada hereje que asome la nariz! —rugió Borgar chocando el puño con su martillo.",
        "—Cazaré a los centinelas en las sombras —añadió Kaelan mostrando sus garras.",
        "Vlad asintió con una sonrisa severa y marcial:",
        "—Un plan agresivo, pero el único viable. Moverse antes de que el enemigo entienda que ya ha perdido.",
        "Tomé mi chaqueta blindada y la ajusté a mis hombros.",
        "El sol comenzaba a teñir las colinas orientales con una luz pálida y fría.",
        "La batalla por el Condado de Morvath apenas estaba comenzando. Y la Sección Uno estaba dispuesta a librarla palmo a palmo hasta la victoria total."
    ]
}

def compilar_capitulo_70():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 70              ")
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
    from agregar_capitulo_70 import CAPITULO_70
    
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
        CAPITULO_65, CAPITULO_66, CAPITULO_67, CAPITULO_68, CAPITULO_69,
        CAPITULO_70
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 70
    txt_cap70 = os.path.join(capitulos_dir, "70_Capitulo_70.txt")
    with open(txt_cap70, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_70['header'].upper()}: {CAPITULO_70['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_70['paragraphs']))
    print(f"[OK] Texto plano Capítulo 70 guardado en:\n  {txt_cap70}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 70
    docx_cap70 = os.path.join(capitulos_dir, "70_Capitulo_70.docx")
    create_single_chapter_document(CAPITULO_70, docx_cap70, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 70 guardado en:\n  {docx_cap70}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 70)
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
    compilar_capitulo_70()
