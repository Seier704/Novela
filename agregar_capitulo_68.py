# -*- coding: utf-8 -*-
import os
import sys

NOVELA_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Novela"
PROYECTO_DIR = r"C:\Users\fsilv\OneDrive\Documentos\GitHub\Proyecto-Minecraft\Novela"

sys.path.insert(0, NOVELA_DIR)
if os.path.exists(PROYECTO_DIR):
    sys.path.insert(0, PROYECTO_DIR)

from generar_libro import create_novel_document, create_single_chapter_document

CAPITULO_68 = {
    "header": "Capítulo 68",
    "title": "El Rastro del Cráneo y la Sangre: La Sombra de Khorne",
    "paragraphs": [
        "Dos horas antes del amanecer, la Sección Uno abandonó las murallas de Morvath bajo un manto de neblina gélida y cenicienta.",
        "Guiados por Julian a través de un túnel de drenaje en desuso que desembocaba más allá de los fosos de la guardia, nos internamos en el desfiladero que conducía hacia los riscos escarpados de Pedregal Alto. Nadie encendió una sola antorcha; el avance se hizo en silencio sepulcral, con los metales de las armas enfundados en cuero y tela para amortiguar cualquier chasquido.",
        "A medida que ganábamos altura entre las gargantas rocosas, el aire se tornó espeso, cargado de una pestilencia metálica que no dejaba lugar a dudas.",
        "—Huele a matadero —susurró Kaelan, deslizándose pegado a los peñascos con las pupilas dilatadas al máximo—. Sangre coagulada, azufre y ceniza humana. Está cerca, al doblar esa cresta.",
        "Faelar subió a un risco con la ligereza de una sombra élfica. Ajustó la cuerda de su arco largo y, tras asomarse apenas unos segundos sobre el borde de piedra caliza, descendió con el rostro ensombrecido:",
        "—Hay una explanada frente a la bocamina principal. Cincuenta hombres armados. Pero no son simples bandidos de bosque.",
        "* * *",
        "Nos arrastramos sobre la panza hasta el borde del acantilado para observar el campamento desde la penumbra superior.",
        "Lo que vimos en la hondonada congeló la sangre de Vance y arrancó una maldición ahogada de Borgar.",
        "Frente a la entrada ennegrecida de la mina abandonada se alzaban estacas de fresno coronadas por cráneos humanos descarnados. En el centro de la explanada ardía un brasero colosal de bronce brujo, sobre el cual se proyectaba una runa tallada en roca viva: el símbolo de ocho filos y la calavera estilizada del Dios de la Sangre, **Khorne**.",
        "Los hombres que rodeaban el fuego no vestían libreas imperiales. Eran fanáticos del Caos con los torsos desnudos cubiertos de cicatrices rituales, corazas herrumbrosas tachonadas de púas y enormes hachas de carnicero con las hojas ennegrecidas de sangre seca.",
        "—Por todos los dioses... —tartamudeó Vance, temblando mientras miraba las estacas—. Un culto al Caos... una secta de sangre en pleno territorio del imperio.",
        "—El Culto a Khorne —masculló Vlad con una frialdad gélida que cortaba el viento—. Fanáticos que rinden tributo a la matanza y la barbarie. No negocian, no se rinden y solo buscan masacrar para apilar cráneos ante su trono de hierro.",
        "—Y miren quiénes están negociando con ellos —intervino Akame, señalando con el dedo hacia el flanco derecho del campamento.",
        "* * *",
        "Dos carromatos pesados, tirados por mulas negras y con lonas manchadas, se detuvieron junto al brasero.",
        "De los carros descendió un oficial con el jubón de cuero de la guardia personal de Valerius de Morvath. El hombre se quitó el yelmo y descargó un cofre de hierro pesado sobre una mesa de sacrificios, abriendo la cerradura para mostrar cientos de monedas de oro y lingotes de plata confiscados a los campesinos.",
        "Frente al oficial emergió de la mina una figura titánica: un sacerdote guerrero de Khorne, cubierto con una máscara de hierro sin ojos y una coraza forjada con calaveras de bronce.",
        "Julian apretó los dientes con tal fuerza que una gota de sangre brotó de su labio inferior:",
        "—Es el capitán de mercenarios de mi hermano Valerius... El perro personal de mi familia.",
        "Faelar aguzó el oído canalizando un susurro de maná de viento para atrapar la conversación que subía desde el foso:",
        "—Dice el oficial que el oro está completo —tradujo el elfo en voz baja—. Y el sacerdote de la máscara le reclama por los tributos vivos. Exige más sangre fresca para encender el portal de la mina.",
        "—El oficial le está respondiendo... —continuó Faelar con el ceño fruncido—. Le asegura que el pueblo de **Sauceda** ha sido seleccionado para la recolección de esta noche. Van a asaltar la aldea al anochecer. Se llevarán a sesenta mujeres y niños para degollarlos en el altar de la mina.",
        "* * *",
        "—¡Malditos engendros! —bramó Borgar en un susurro cargado de furia ciega, aferrando el mango de su martillo rúnico—. ¡Bajemos ahora mismo y reventemos a esos sacerdotes de mierda contra la piedra!",
        "Vance desenvainó medio palmo de su estoque, con los ojos inyectados en rabia:",
        "—¡Concuerdo con el enano! ¡Están vendiendo a su propia gente! ¡Es traición a la Corona y al honor de la nobleza!",
        "Puse mi mano pesada sobre el hombro de Borgar y la otra sobre el brazo de Vance, clavando mis ojos en ambos con una firmeza que los paralizó:",
        "—Nadie se mueve. Ni un solo paso.",
        "—¡¿Pero Seier, van a masacrar a Sauceda?! —protestó Vance con la voz rota.",
        "—Usa la cabeza, Holst —le espeté en un susurro rotundo—. Si atacamos aquí ahora, mataremos a treinta cultistas, pero el oficial de Valerius morirá o escapará en el caos. El conde Raymond quemará los archivos en el castillo, dirá que nosotros inventamos todo para derrocarlo y el Senado de Astraea nos colgará por traición e insubordinación militar. La palabra de un conde vale más que la de siete cadetes ante la corte imperial.",
        "Vlad asintió de inmediato con la mirada afilada de un comandante de asedio:",
        "—Seier tiene toda la razón. Necesitamos la prueba irrefutable. El documento, la orden sellada o capturar a los mercenarios y cultistas en flagrancia dentro del pueblo.",
        "Julian sacó su daga y señaló un mapa de cuero arrugado:",
        "—Ese cofre de hierro contiene las cartas de entrega firmadas por mi padre y por Valerius con el sello de cera negra de la casa. Si el oficial mercenario va a Sauceda a coordinar el ataque con los cultistas, llevará la orden condal consigo para exigir que los campesinos no opongan resistencia armada.",
        "* * *",
        "Miré a Akame, que ya había memorizado la ruta de escape de los carromatos.",
        "—El plan cambia —dictaminé en voz baja—. Dejaremos que los cultistas y los mercenarios salgan confiados hacia Sauceda creyendo que el pueblo está indefenso. Nosotros nos adelantaremos por los atajos del bosque que conoce Julian.",
        "—Llegaremos dos horas antes que ellos —aseguró Julian con determinación férrea.",
        "—Exacto —continué—. En Sauceda usaremos tu sello para movilizar a los treinta veteranos de la milicia cívica. Los atrincheraremos en las granjas y en los graneros. Cuando los cultistas de Khorne y los hombres de Valerius entren a secuestrar a las familias, cerraremos el cerco, recuperaremos los documentos sellados del conde y no dejaremos a un solo hereje vivo.",
        "Akame asintió con una calma gélida, acariciando el pomo de *Murasame*:",
        "—Eliminar a los demonios en el pueblo. Asegurar la prueba. Ejecutar al conde bajo la ley del Rey.",
        "—Así se habla —sonrió Borgar mostrando los dientes bajo la barba cobriza—. Emboscada en el barro. Mi martillo tiene hambre de calaveras de hojalata.",
        "—Cubriré los tejados de Sauceda con flechas perforantes —dijo Faelar tensando la cuerda.",
        "Kaelan flexionó sus garras con un ronquido bajo:",
        "—Cortaré los tendones de los que intenten huir.",
        "Vance envainó su estoque, respirando hondo y asintiendo con una madurez que jamás había mostrado en las aulas de la capital:",
        "—Defenderé a los campesinos junto a la milicia. Por el verdadero honor de Solaria.",
        "Nos deslizamos riscos abajo en completo sigilo, dejando atrás el humo rojizo y las blasfemias de las bocaminas.",
        "El destino de Morvath se decidiría en las calles de adobe de Sauceda.",
        "Y la Sección Uno estaba lista para bañar en sangre a los que osaban rezarle al Caos."
    ]
}

def compilar_capitulo_68():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 68              ")
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
        CAPITULO_65, CAPITULO_66, CAPITULO_67, CAPITULO_68
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 68
    txt_cap68 = os.path.join(capitulos_dir, "68_Capitulo_68.txt")
    with open(txt_cap68, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_68['header'].upper()}: {CAPITULO_68['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_68['paragraphs']))
    print(f"[OK] Texto plano Capítulo 68 guardado en:\n  {txt_cap68}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 68
    docx_cap68 = os.path.join(capitulos_dir, "68_Capitulo_68.docx")
    create_single_chapter_document(CAPITULO_68, docx_cap68, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 68 guardado en:\n  {docx_cap68}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 68)
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
    compilar_capitulo_68()
