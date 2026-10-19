# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_67 = {
    "header": "Capítulo 67",
    "title": "Monedas Negras y Secretos de Medianoche: El Pacto con Julian",
    "paragraphs": [
        "La reunión oficial con el Conde Raymond fue tan insultantemente breve como patética.",
        "Apenas duró cinco minutos. El señor de Morvath, mientras desgarraba con los dientes el muslo grasiento de un capón asado y apuraba una copa de vino especiado, despachó la situación con una indolencia que rayaba en el descaro:",
        "—La orden es simple, cadetes de la capital —soltó el conde limpiándose la barba con la manga de seda—. En el pueblo de Sauceda hay labradores revoltosos que se niegan a pagar los diezmos atrasados, y en Pedregal Alto hay vagabundos que merodean las bocaminas. Vayan allí, cuelguen a un par de cabecillas para dar el ejemplo, despejen los caminos y mándenme el reporte firmado para que el Senado de Astraea me deje de fastidiar con auditorías.",
        "Valerius, con una sonrisa burlona y la mano apoyada en su espada enjoyada, nos señaló la salida con un gesto de desprecio:",
        "—Pueden alojarse en las caballerizas del patio bajo. No esperen sábanas limpias; no acostumbramos mimar a los lacayos del ejército.",
        "Vlad y yo mantuvimos la compostura militar, cuadrándonos con una reverencia seca y sin discutir. Discutir con un necio frente a sus guardias mercenarios solo habría puesto en alerta a quien manejaba los hilos.",
        "* * *",
        "La medianoche cayó sobre Morvath como un sudario helado y brumoso.",
        "Nos encontrábamos en el viejo cuartel de la guardia, una estructura de tablones húmedos y vigas ennegrecidas situada en el límite inferior del castillo. Borgar afilaba silenciosamente las aristas de su martillo, Faelar emplumaba flechas de caza rúnica junto a la chimenea apagada y Akame masticaba pacientemente un trozo de pan duro con los ojos clavados en la puerta.",
        "De pronto, Kaelan alzó la cabeza de golpe. Sus orejas de lince giraron hacia el callejón trasero y sus pupilas se dilataron en la penumbra:",
        "—Alguien se acerca —siseó el semihumano, flexionando las garras—. Pasos ligeros, una sola persona. No trae armadura de metal; no es un mercenario.",
        "La puerta trasera chirrió con suavidad y una figura encapuchada se deslizó al interior del recinto, atrancando el pestillo de madera tras de sí. Al retirarse la capucha de lana oscura, la luz de la vela reveló el rostro exhausto y demacrado de Julian de Morvath.",
        "—Gracias a los dioses que siguen despiertos... —susurró Julian con la respiración entrecortada.",
        "—¿Dónde dejaste a tu hermano pequeño? —pregunté sin rodeos, midiendo sus movimientos.",
        "—Leo está seguro, escondido en la recámara de su antigua nodriza —respondió Julian, bajando la voz al mínimo—. Tuve que esperar a que mi padre cayera inconsciente por el vino y a que Valerius se encerrara con las rameras de la taberna alta para poder escabullirme.",
        "* * *",
        "Vlad le acercó una banqueta de madera y Julian se sentó con pesadez, pasándose ambas manos temblorosas por el rostro:",
        "—Lo que les dijo mi padre en el salón es una farsa absoluta. Una mentira repugnante para tapar el pozo de mierda en el que está hundido este condado.",
        "—Te escuchamos, muchacho —dijo Vlad, cruzando los brazos con gesto severo—. Habla claro.",
        "—Mi padre no solo está cobrando diezmos abusivos; está exprimiendo a los campesinos hasta dejarlos en los puros huesos —reveló Julian con los ojos inyectados en sangre por la rabia contenida—. Todo el grano cosechado es confiscado y vendido a contrabandistas forasteros; cada moneda de plata, cada alhaja y cada buey de tiro son arrebatados por la fuerza. Mi padre y Valerius están acaparando todo el oro de la ciudad en las criptas del castillo mientras la gente muere de disentería y hambre en las zanjas.",
        "—¿Para qué demonios quiere un conde amontonar tanto oro si sus tierras se mueren? —gruñó Borgar frunciendo el ceño—. Un señor feudal necesita campesinos vivos para mantener su feudo.",
        "Julian tragó saliva con dificultad. Su mirada descendió hacia las tablas del suelo, cargada de un pavor genuino que le erizó la piel:",
        "—Porque creo... creo que mi padre se vendió a los demonios.",
        "Un silencio denso y sepulcral cayó sobre la estancia.",
        "Vance se estremeció, pegando la espalda a la pared, mientras Akame ladeaba la cabeza con la mirada ensombrecida.",
        "—¿Tienes pruebas o solo son sospechas? —inquirí con voz fría como el hielo.",
        "—No tengo un contrato firmado en pergamino negro, pero no soy ciego —respondió Julian con voz quebrada—. Desde hace tres meses, han comenzado a desaparecer personas en los pueblos periféricos. No hombres armados; mujeres jóvenes y niños pequeños. En Sauceda se llevaron a doce muchachas de sus chozas en mitad de la noche; en Pedregal Alto, cinco niños desaparecieron mientras jugaban cerca de los pozos. No piden rescates, no hay cadáveres abandonados en los bosques y la guardia de mi padre tiene órdenes estrictas de no investigar.",
        "Julian apretó los puños hasta clavarse las uñas en las palmas:",
        "—He visto carromatos cubiertos con lonas negras salir del castillo a la medianoche escoltados por los mercenarios personales de Valerius. Se dirigen hacia las minas abandonadas de Pedregal Alto. Y del fondo de esos túneles... he olido el miasma. Es un hedor dulzón, pútrido y caliente que no pertenece a este mundo.",
        "* * *",
        "Julian metió la mano en su jubón de cuero y extrajo un objeto pesado envuelto en un paño de terciopelo. Al desenvolverlo, sobre la mesa reposó un anillo de plata maciza con el sello heráldico de la Casa Morvath: un halcón bicéfalo posado sobre una espada invertida.",
        "—Este es mi sello personal como segundo heredero —dijo Julian, empujándolo hacia mis manos—. Si mi padre se entera de esto, me colgará del roble más alto por traición a la sangre. Pero prefiero morir como un traidor que seguir siendo cómplice de este matadero.",
        "Miré el anillo de plata y luego a los ojos oscuros y desesperados del muchacho.",
        "—¿Qué podemos hacer con esto? —pregunté.",
        "—Los mercenarios que vieron en los portones son perros sarnosos contratados por Valerius; ladrones y violadores con uniforme que cobran por aterrorizar al pueblo —explicó Julian con firmeza—. Pero en el cuartel bajo de Sauceda aún queda lo que fue la verdadera milicia cívica del condado. Son unos treinta veteranos leales a la tierra, hombres con familias e hijos que están desesperados, pero que no podían actuar sin una orden de la casa noble. Si les presentan este sello, la milicia los reconocerá como comandantes legítimos y los obedecerá a muerte.",
        "Julian se puso de pie, ajustándose la daga que llevaba al cinto:",
        "—Iré con ustedes. De incógnito, cubierto con una capa de explorador. Conozco cada túnel, cada sendero de los bosques y las entradas secundarias de las minas.",
        "* * *",
        "Vlad y yo intercambiamos una mirada de absoluto entendimiento táctico.",
        "—El chico tiene agallas —reconoció Vlad, asintiendo con respeto marcial—. Con la milicia local aseguramos el perímetro de los pueblos y evitamos que los mercenarios de Valerius nos apuñalen por la espalda.",
        "—Y si encontramos demonios o bestias corrompidas en esas minas... —Akame se puso de pie, acariciando con el pulgar la guardia dorada de *Murasame*—, simplemente los eliminaremos.",
        "—¡Por mi barba, cuenten con mi martillo! —bramó Borgar en un susurro áspero—. ¡Aplastaré la cabeza de cualquier bastardo que toque a un niño!",
        "—Mi arco cubrirá la noche —añadió Faelar con mirada gélida.",
        "Kaelan mostró sus colmillos con una sonrisa feroz desde la sombra:",
        "—Huelo el rastro de la sangre. Será una buena cacería.",
        "Incluso Vance von Holst, a pesar de que sus manos aún temblaban levemente, desenvainó su estoque dorado y apretó la mandíbula:",
        "—Si un noble de este imperio pactó con la oscuridad... mancilló el honor de la aristocracia. No permitiré que esa escoria se salga con la suya.",
        "Tomé el sello de plata de la mesa y lo guardé en el bolsillo interior de mi chaleco de combate.",
        "—El plan está trazado —dije, mirando a Julian y a mi escuadrón—. Partimos hacia Sauceda dos horas antes del amanecer, antes de que los mercenarios del conde despierten de su resaca. Vamos a devolverle este feudo a su gente.",
        "En la penumbra de las caballerizas, la Sección Uno selló su primer pacto de sangre.",
        "Y la noche de Morvath comenzó a temblar ante el acero que se avecinaba."
    ]
}

def compilar_capitulo_67():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 67              ")
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
        CAPITULO_65, CAPITULO_66, CAPITULO_67
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 67
    txt_cap67 = os.path.join(capitulos_dir, "67_Capitulo_67.txt")
    with open(txt_cap67, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_67['header'].upper()}: {CAPITULO_67['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_67['paragraphs']))
    print(f"[OK] Texto plano Capítulo 67 guardado en:\n  {txt_cap67}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 67
    docx_cap67 = os.path.join(capitulos_dir, "67_Capitulo_67.docx")
    create_single_chapter_document(CAPITULO_67, docx_cap67, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 67 guardado en:\n  {docx_cap67}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 67)
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
    compilar_capitulo_67()
