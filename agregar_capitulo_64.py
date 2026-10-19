# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_64 = {
    "header": "Capítulo 64",
    "title": "Nuevos Vientos en Astraea: Aulas Mixtas y la Hermandad Eterna",
    "paragraphs": [
        "El carruaje imperial frenó con un crujido seco frente a los monumentales portones de hierro forjado de la Gran Academia de Astraea.",
        "Detrás, las motos del Páramo apagaron sus motores con un bufido metálico. Dejábamos atrás el calor y la arena de Costa Esmeralda; el aire fresco del otoño capitalino nos daba la bienvenida al segundo semestre.",
        "Apenas pisamos los adoquines del patio central, los cadetes de los cursos superiores se abrieron a nuestro paso, murmurando con respeto. La Medalla del León de Oro de Seier y los ecos de la victoria en la Roca Carmesí habían borrado para siempre las burlas hacia los antiguos parias.",
        "—Vaya cambio —soltó Lighter ajustándose los guantes—. Parece que ahora los señoritos nobles aprendieron a agachar la cabeza.",
        "—Más les vale —replicó Caesar King, ajustando la correa de su pesado escudo a la espalda—. Si alguno vuelve a decir que la gente del páramo huele a chatarra, le incrusto el gladius en el pupitre.",
        "Yami expulsó una densa nube de humo de su puro, caminando con las manos en los bolsillos junto a ellos:",
        "—No gastes saliva, jefa. El miedo educa más rápido que cualquier profesor de modales.",
        "* * *",
        "El repique de la campana de bronce convocó a toda la academia a la Plaza de Armas.",
        "Frente a las escalinatas de mármol, el Instructor Jefe Gerald esperaba de pie, con los brazos a la espalda y el rostro severo como una pared de granito:",
        "—¡Firmes todos! —su voz de trueno silenció de golpe los murmullos de cientos de estudiantes—. ¡Iniciamos el segundo semestre académico con un cambio radical dictado por el Alto Consejo y el Senado de Astraea!",
        "Gerald recorrió las filas con la mirada antes de soltar la bomba:",
        "—Queda abolida la segregación de aulas por linaje, dinero o casta de ingreso. Se acabaron los privilegios feudales. Ya no existen la Clase A-1 ni la Clase F-2.",
        "Los nobles de familias condales se miraron entre sí, escandalizados.",
        "—A partir de hoy, todas las secciones serán mixtas —bramó Gerald—. Nobles, plebeyos, cadetes del páramo y forasteros compartirán los mismos bancos y los mismos entrenamientos. Quien no lo tolere, tiene los portones abiertos para largarse a su mansión.",
        "* * *",
        "El cristal rúnico del patio destelló, proyectando los listados de las nuevas secciones.",
        "Piper corrió la primera a buscar su nombre y pegó un grito que alertó a media plaza:",
        "—¡Nooo! ¡Oigan, esto es una estafa! ¡Me mandaron a la Sección Seis con la gente de la antigua Clase B! ¡¿Quién me va a cuidar las tuercas ahora?!",
        "—Deja de lloriquear, enana —se burló Caesar dándole un coscorrón cariñoso—. Si alguien te toca una llave inglesa, silba y voy a partirle el banco. A ver dónde caí yo... ¡¿Qué?! ¡¿Sección Cinco?!",
        "Natsuki buscó su nombre en la misma columna y parpadeó sorprendido:",
        "—Espera... Caesar, ¡a mí también me tocó la Sección Cinco!",
        "Caesar soltó una carcajada ronca, agarrando a Natsuki por el hombro con fuerza de tenaza:",
        "—¡Miren qué suerte! ¡Me tocó de compañero de banco el domador de cerdos! Prepárate, cuñado, que en clase de teoría táctica te voy a exigir el triple de flexiones.",
        "Lucy, con sus coletas rubias saltando de indignación, se abrió paso a empujones para ver la lista:",
        "—¡¿Cómo que la Cinco?! ¡A mí me mandaron a la Sección Dos! ¡Eso es trampa! —Lucy clavó sus ojos dorados en Natsuki, con las mejillas ardiendo de rubor y los puños cerrados—. ¡Oye, Subaru, más te vale no descuidarte ni hacer tonterías con los nobles de tu clase! ¡Y si alguna chica de la nobleza se te acerca a pedirte apuntes, me avisas para quemarle la falda!",
        "—Te lo prometo, Lucy —respondió Natsuki con una sonrisa tierna y tranquila, rascándose la nuca.",
        "—¡Menos charla y más disciplina! —rugió una voz imponente a su espalda.",
        "Erza Scarlet se cruzó de brazos, con su larga cabellera escarlata brillante reluciendo al sol sobre su armadura de placas plateadas:",
        "—Lucy, estás en mi misma aula, la Sección Dos. Y conmigo de compañera no toleraré distracciones amorosas ni berrinches. Nos sentaremos al frente y seremos las mejores de la clase.",
        "—¡P-pero Erza...! —protestó Lucy con un puchero resignado.",
        "—Dante Vance también está con nosotras —añadió Erza, señalando al trampero que limpiaba la recámara de su mosquete rúnico—. Mantendremos la compostura militar en todo momento.",
        "* * *",
        "En otro rincón del cristal rúnico, Ken soltó una maldición al ver su nombre:",
        "—¡Maldita sea! ¡A mí me metieron en la Sección Cuatro con un montón de presumidos de la vieja A-1! ¡¿Y Burnice dónde está?!",
        "—¡A mí me tocó la Sección Tres, Ken-Ken! —gritó Burnice White desde atrás, agitando una botella de licor chispeante con canela—. ¡No te preocupes! ¡Voy a prepararles bebidas picantes a los de la Tres para que no se duerman!",
        "A su lado, Lady Elysea de Vane se acomodó un mechón de su cabello plateado como luz de luna, mirando a Ken con sus ojos gélidos de color zafiro:",
        "—Cálmate, Ken. Yo estoy asignada a tu misma aula en la Sección Cuatro. Si los nobles de esa clase intentan faltarte al respeto por tu origen, mi estoque *Escarcha Silenciosa* les recordará cuál es su lugar.",
        "—¡Ja! Me gusta cómo hablas, muñeca de hielo —sonrió Ken, chocando los puños con chispas solares.",
        "* * *",
        "Yami, apoyado contra una columna, bostezó mientras leía su destino:",
        "—Sección Seis. Genial. Parece que me toca compartir aula con Kaoruko y Lighter.",
        "Kaoruko Waguri, con su cabello azabache cayéndole suave sobre los hombros y sus ojos cálidos llenos de dulzura, apretó su estuche de cuero contra el pecho:",
        "—Me aseguraré de llevar galletas y pasteles de manzana para todos, Yami-san... aunque me da pena no estar sentada junto a Seier...",
        "Lighter le dio una palmada en el hombro a la repostera:",
        "—No te preocupes, Kaoruko. Nadie se va a propasar en nuestra clase. Quien te mire feo se gana un jab de combustión.",
        "Mientras tanto, una chica con bata blanca manchada de reactivos, despeinada y con un matraz burbujeante en la mano se acercó a tropezones:",
        "—¡Fascinante! ¡Absolutamente caótico! —exclamó Agnes Tachyon, ajustándose las gafas protectoras mientras un humo violáceo salía de su frasco—. ¡Me asignaron a la Sección Cinco con Caesar y Natsuki! ¡El balance termodinámico de esa sala va a ser una maravilla para mis experimentos de fósforo!",
        "Natsuki tragó saliva, mirando el frasco con pánico evidente:",
        "—Por favor, señorita Agnes... prométame que no detonará nada dentro del aula.",
        "* * *",
        "Por su parte, Renji Abarai se tiraba de su coleta roja encendida mirando el cristal:",
        "—¡¿Es una broma?! ¡Me mandaron a la Sección Siete con Freya y Zemina! ¡¿Y Rukia dónde quedó?!",
        "—Estoy en la Sección Tres con Burnice, Hoshina y Pulchra, Renji —respondió Rukia Kuchiki con su calma habitual, con su cabello negro azabache y mechón frontal enmarcando su rostro menudo y ojos violetas—. No hagas un escándalo innecesario. Nos veremos en el almuerzo y en los entrenamientos libres.",
        "Lortelle Kecheln, con su cabello castaño claro con reflejos dorados y una sonrisa calculadora en sus ojos verdes felinos, cruzó los brazos satisfecha:",
        "—A mí también me tocó la Sección Siete, Renji. No pongas esa cara. Tener a la F-2 desparramada en todas las aulas es el mejor negocio que pudimos pedir. La Compañía Elte tendrá ojos y oídos en cada rincón de la academia.",
        "* * *",
        "Vlad se acercó a mi costado, impecable con su porte de caballero noble y su espada bastarda al cinto:",
        "—Seier, ¿ya viste la Sección Uno?",
        "Miré el encabezado del cristal.",
        "En la Sección Uno figuraban tres nombres muy claros: **Seier Silva**, **Vlad** y **Akame**.",
        "Akame dio un mordisco a un trozo de cecina que llevaba en la mano, con sus ojos rojos penetrantes fijos en la lista:",
        "—Buena combinación. Si algún noble busca pelea en nuestra aula, no durará mucho.",
        "—Totalmente de acuerdo —asintió Vlad con una sonrisa marcial—. Tenemos la vanguardia más pesada de la academia en la misma sala.",
        "Di dos pasos al frente y solté un silbido agudo que resonó por toda la plaza. De inmediato, los dieciséis guerreros que habíamos sobrevivido al asedio Skaven se reunieron a mi alrededor, formando un círculo compacto y blindado.",
        "—Escuchen bien todos —dije, mirando a cada uno a los ojos—. Las parejas no están juntas. Las aulas no tienen el cartel de la F-2. Pero un aula solo son cuatro paredes de ladrillo. La verdadera F-2 somos nosotros.",
        "Ken sonrió con ferocidad. Caesar levantó el pulgar. Lucy asintió decidida junto a Erza. Kaoruko me miró con devoción absoluta, y Elysea y Lortelle me dedicaron sonrisas cómplices.",
        "—En este semestre todos cumpliremos los diecisiete años —continué con voz firme—. Vamos a conquistar cada una de las aulas en las que nos metieron. Demostraremos que no importa dónde nos sienten: Astraea entera ahora nos pertenece.",
        "—¡¡POR LA CLASE F-2!! —rugieron todos al unísono, chocando puños y armas bajo el cielo otoñal de la capital.",
        "El segundo semestre había comenzado, y la leyenda apenas estaba empezando."
    ]
}

def compilar_capitulo_64():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 64 (REVISADO)   ")
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
        CAPITULO_60, CAPITULO_61, CAPITULO_62, CAPITULO_63, CAPITULO_64
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 64
    txt_cap64 = os.path.join(capitulos_dir, "64_Capitulo_64.txt")
    with open(txt_cap64, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_64['header'].upper()}: {CAPITULO_64['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_64['paragraphs']))
    print(f"[OK] Texto plano Capítulo 64 guardado en:\n  {txt_cap64}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 64
    docx_cap64 = os.path.join(capitulos_dir, "64_Capitulo_64.docx")
    create_single_chapter_document(CAPITULO_64, docx_cap64, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 64 guardado en:\n  {docx_cap64}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 64)
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
    compilar_capitulo_64()
