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
        "Detrás de nosotros, los motores rugientes de las motocicletas de Calydon se apagaron, dejando en el aire un leve rastro de vapor y ozono. Habíamos dejado atrás las aguas cristalinas y la arena dorada de la Costa Esmeralda; la brisa fresca del otoño capitalino nos daba la bienvenida a una nueva etapa.",
        "Apenas pusimos un pie sobre los adoquines del patio principal, el cambio en el ambiente fue palpable.",
        "Ya no había miradas de desprecio ni murmullos burlones hacia nuestros uniformes. Los cadetes de otros cursos que transitaban por los jardines se detenían en seco, abriéndonos paso con un respeto casi reverencial. Los ecos de la Batalla de la Roca Carmesí y la condecoración del León de Oro con Hojas de Roble habían sacudido hasta el último rincón de la capital.",
        "—Vaya, miren eso —soltó Lighter con una sonrisa de lado, ajustándose los guantes de cuero—. Parece que los señoritos nobles aprendieron modales mientras estábamos de pesca.",
        "—Más les vale —bufó Caesar King con su habitual paso pesado, apoyando su enorme escudo a la espalda—. Si alguno vuelve a llamarnos 'escoria del páramo', le enseñaré a volar sin escoba.",
        "Yami expulsó una densa calada de su puro, caminando con las manos en los bolsillos junto a Caesar:",
        "—No gastes saliva, jefa. El miedo educa más rápido que cualquier profesor de etiqueta.",
        "* * *",
        "El sonido bronco de una campana de bronce retumbó desde la torre central, convocando a todos los cadetes hacia la Plaza de Armas.",
        "Frente a las imponentes escalinatas de mármol, el Instructor Jefe Gerald aguardaba con su uniforme de gala impecable y los brazos cruzados tras la espalda. Su rostro curtido por cien batallas lucía más severo que de costumbre.",
        "—¡Firmes todos! —rugió Gerald, acallando al instante los cuchicheos de los cientos de estudiantes reunidos—. ¡Bienvenidos al inicio del segundo semestre académico!",
        "Gerald recorrió las filas con la mirada, deteniéndose apenas un segundo en nuestro grupo con un brillo de orgullo que no se molestó en disimular del todo.",
        "—Durante el receso estival, el Alto Consejo del Imperio y el Senado de Astraea han promulgado una reforma estructural inmediata y obligatoria —anunció con voz de trueno—. A partir de este ciclo, queda abolida la segregación de aulas por estatus nobiliario o linaje de cuna. Se acabaron los privilegios heredados y las aulas de casta. Ya no existen la Clase A-1 ni la Clase F-2.",
        "Una ola de murmullos indignados estalló entre los cadetes de familias condales y ducales.",
        "—¡Silencio! —bramó Gerald con tal fuerza que las baldosas vibraron—. A partir de hoy, todas las secciones serán heterogéneas y mixtas. Nobles, plebeyos, cadetes del Páramo Exterior y forasteros compartirán los mismos pupitres, las mismas asignaturas teóricas y los mismos barrotes de entrenamiento. Si alguien cree que su sangre azul vale más que el sudor y el acero de sus compañeros, las puertas de salida están abiertas.",
        "* * *",
        "Minutos después, la enorme pantalla de cristal rúnico en el centro de la plaza se iluminó, desplegando los listados de las nuevas secciones.",
        "Piper Wheel fue la primera en correr hacia el frente, entrecerrando los ojos para leer los nombres grabados en luz mágica. De pronto, soltó un chillido lastimero:",
        "—¡Nooo! ¡Oigan, esto es una tragedia! ¡Nos separaron a casi todos! ¡Me tocó en la Sección Seis con un montón de estirados de la vieja Clase B! ¡¿Quién me va a cuidar las herramientas y a probar mis galletas ahora?!",
        "Caesar soltó una carcajada estruendosa, dándole un suave tirón de orejas a la mecánica:",
        "—Deja el drama, enana. Si algún novato de la B se atreve a tocarte una sola tuerca, silbas y voy a partirle el pupitre en las costillas. Además, miren esto... a mí me mandaron a la Sección Tres. Y miren quién es mi compañero de banco.",
        "Yami miró el cristal, entrecerró los ojos y soltó un suspiro resignado:",
        "—Tenía que ser. Parece que tendré que aguantar tus rugidos también en clase de estrategia. Mientras no me tires tiza cuando esté durmiendo, sobreviviremos.",
        "—¿Dormir? ¡Ni lo sueñes, fumador! Te voy a tener despierto a codazos —le retó Caesar con una sonrisa pícara y desafiante que hizo sonreír a Yami de lado.",
        "A su lado, Ken chocó los puños con entusiasmo. Pequeñas llamaradas doradas brotaron de sus nudillos:",
        "—¡A mí me da exactamente igual el número de aula! ¡Me pusieron en la Sección Cuatro con Burnice y un par de presumidos de la antigua A-1! ¡Seier, te juro que los voy a hacer sudar fuego en cada combate de entrenamiento!",
        "Burnice aplaudió entusiasmada dando saltitos:",
        "—¡Siii! ¡Voy a prepararles mis cocteles chispeantes con canela y pimienta a todos los nuevos compañeros! ¡Vamos a encender esa clase desde el primer día!",
        "* * *",
        "Cerca de allí, Natsuki buscaba su nombre en el panel rúnico con nerviosismo. De pronto, sus ojos se abrieron y una sonrisa sincera iluminó su rostro:",
        "—¡Estoy en la Sección Dos! Y... espera... ¡Lucy también está en la Dos!",
        "Lucy, que estaba a su lado con los brazos cruzados y la barbilla en alto fingiendo desinterés, se puso roja como una manzana en cuanto escuchó su nombre. Trató de disimular con un mohín orgulloso:",
        "—N-no te emociones tanto, tonto. Solo significa que tendré que vigilarte de cerca para que no hagas el ridículo frente a los nobles. Y más les vale a esos engreídos no causarte problemas, o mis pequeños jabalíes les quemarán los libros de texto.",
        "Natsuki rió bajito, inclinándose hacia ella para susurrarle con cariño:",
        "—Sé que me vas a cuidar, pequeña fiera.",
        "—¡C-cállate y camina! —replicó Lucy dándole un codazo cómplice en las costillas, aunque sus dedos buscaron discretamente los de Natsuki por un instante.",
        "Renji Abarai se rascó la nuca mientras leía su asignación:",
        "—A mí me mandaron a la Sección Cinco con Lighter y Vlad. Nada mal. Pura fuerza de choque.",
        "—Pórtate bien, Renji —le aconsejó Rukia con una sonrisa serena y dulce—. A mí me tocó en la Sección Uno junto a Akame y Hoshina.",
        "—Ten por seguro que ningún noble se sobrepasará en esa aula —comentó Akame con tono tranquilo, masticando una tira de carne seca—. No mientras tengamos hambre de entrenar.",
        "* * *",
        "Lortelle Kecheln se colocó a mi derecha, acomodándose un mechón rubio tras la oreja con una mirada brillante de astucia:",
        "—¿Te das cuenta de lo que esto significa en realidad, Seier? El senado creyó que al dispersarnos debilitaría la cohesión de la F-2, pero en realidad nos acaban de regalar influencia directa en cada aula del campus. La red comercial y de información de la Compañía Elte se moverá el doble de rápido.",
        "Lady Elysea asintió con porte aristocrático impecable:",
        "—Concuerdo plenamente con Lortelle. En mi aula hay tres herederos que solían burlarse de nosotros. Ahora tendrán que pedirnos consejo si quieren aprobar las pruebas de campo.",
        "Kaoruko tomó mi mano con delicadeza, entrelazando sus dedos tibios con los míos mientras me miraba con una dulzura inquebrantable:",
        "—Lo mejor de todo es que tú y yo estamos juntos en la Sección Uno, Seier. No importa qué otros cadetes entren... mi lugar siempre será a tu lado.",
        "—Y el mío junto al de ustedes tres —le respondí, apretando su mano con devoción.",
        "* * *",
        "Di un paso al frente y llamé la atención de todos con un silbido agudo. El grupo entero —los catorce guerreros que habíamos desafiado a la muerte en las trincheras— formó un círculo apretado alrededor de mí.",
        "—Miren a su alrededor, muchachos —les dije con voz firme y serena—. Un aula de clases era solo un cartel de madera colgado en una puerta. La F-2 nunca fueron cuatro paredes de piedra. La F-2 somos nosotros.",
        "Todos asintieron con rostros decididos y miradas encendidas.",
        "—No importa si de ocho a dos de la tarde nos sentamos en pupitres distintos. En los patios de combate, en los comedores, en los calabozos y en la vida real, seguimos siendo la misma manada indestructible. En este nuevo semestre todos nosotros cumpliremos los diecisiete años. Vamos a demostrarle a este imperio que no importa dónde nos pongan... Astraea entera ahora camina a nuestro paso.",
        "—¡¡ASÍ SE HABLA!! —rugió Ken alzando el puño al cielo.",
        "—¡Aplastaremos a quien se interponga! —gritó Caesar chocando su escudo contra el puño de Yami.",
        "Extendí mi mano hacia el centro y todos los demás —Kaoruko, Lortelle, Elysea, Yami, Caesar, Ken, Burnice, Natsuki, Lucy, Piper, Lighter, Renji, Rukia, Akame, Pulchra y Hoshina— apilaron las suyas sobre la mía en un juramento silencioso y feroz.",
        "Las clases mixtas acababan de comenzar, pero la hermandad de la F-2 ya era eterna."
    ]
}

def compilar_capitulo_64():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 64              ")
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
