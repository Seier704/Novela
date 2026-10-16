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

CAPITULO_43 = {
    "header": "Capítulo 43",
    "title": "Veneno en las Entrañas",
    "paragraphs": [
        "El suelo tembló como si las placas tectónicas estuvieran desgarrándose bajo nuestros pies.",
        "Desde la cortina de humo sulfuroso del cañón, cuatro moles titánicas de más de cuatro metros de altura emergieron al galope. Eran las temidas **Ratas Ogro**: abominaciones monstruosas nacidas de la mutación forzada y la magia oscura de piedra bruja. Sus torsos hipertrofiados estaban cubiertos de costurones purulentos, placas de hierro clavadas directamente sobre su carne y garras metálicas oxidadas que chorreaban ácido verdoso.",
        "—¡Abran fuego pesado! ¡No permitan que toquen el rastrillo exterior! —bramó Gerald desenvainando un mandoble de dos manos encantado con maná de viento.",
        "Una lluvia de flechas incendiarias, lanzas de fuego y proyectiles de hielo cayó desde las almenas sobre las bestias. Sin embargo, el dolor solo enfurecía más a las criaturas; con la mirada inyectada en una demencia carmesí, la primera Rata Ogro embistió directamente contra el portón levadizo.",
        "*¡¡¡BOOOOMMMM!!!*",
        "El impacto de toneladas de músculo y metal hizo astillas los travesaños de roble reforzado. La pesada reja de hierro se dobló hacia adentro con un chirrido agónico de goznes destrozados, abriendo una brecha de cinco metros en la entrada exterior.",
        "El pánico amenazó con quebrar a los cadetes más jóvenes, pero Vlad no concedió un solo segundo a la vacilación.",
        "—¡Compañía este y norte, no cedan terreno! —rugió Vlad, cuya voz resonaba con la serenidad helada de un general experimentado—. ¡Lanceros pesados, formen embudo en el cuello del patio! ¡Dirijan a las bestias a la zona de aniquilación!",
        "La disciplina del joven caballero del Ducado transformó el caos en una trampa mortal.",
        "* * *",
        "Poco a poco, con sangre fría y un trabajo en equipo implacable, comenzamos a darle vuelta a la situación.",
        "En lugar de enfrentarlas a campo abierto, Vlad coordinó a los cadetes para que formaran una muralla de escudos entrelazados en los flancos, forzando a dos de los gigantes a entrar en fila india hacia el patio adoquinado. Desde los balcones superiores, Agnes Tachyon vio su oportunidad dorada.",
        "—¡A ver cómo digieren esto, monstruosidades sin refinar! —chilló Agnes, arrojando tres frascos de vidrio sellados con cera negra sobre el lomo de la primera bestia.",
        "Los viales se estrellaron contra las placas de piedra bruja del monstruo. Al entrar en contacto con el reactivo de Agnes, los injertos mutantes de la criatura comenzaron a burbujear violentamente, desatando una reacción química endotérmica que cristalizó los tejidos musculares del ogro, haciéndolo aullar de dolor y reduciendo su movilidad a la mitad.",
        "—¡Ahora, Silva! ¡Rompe su núcleo! —me gritó Agnes desde arriba.",
        "No necesité que me lo repitiera.",
        "Exploté hacia adelante con un pisotón de Muay Thai que agrietó los adoquines bajo mis botas. La Rata Ogro lanzó un zarpazo descendente capaz de partir un carruaje en dos; me deslicé por debajo de su garra en un ángulo oblicuo con el juego de piernas del Jeet Kune Do, canalizando mi maná de **Tierra** para densificar mi masa muscular y mi maná de **Fuego** para generar una propulsión explosiva.",
        "Planté mi talón izquierdo y descargué una patada baja devastadora contra la parte posterior de su rodilla lisiada.",
        "*¡¡KRAAAAK!!*",
        "El crujido de tendones y hueso rompiéndose resonó en todo el patio. La mole cayó sobre una rodilla con un alarido gutural, dejando su cabeza expuesta a mi altura. Sin perder un ápice de inercia, giré sobre mi eje y clavé un codazo ascendente envuelto en llamas comprimidas directo en su mandíbula acorazada.",
        "El impacto perforó las placas de metal y desintegró la base de su cráneo. La bestia cayó inerte sobre el suelo, levantando una densa polvareda.",
        "A pocos metros de allí, Ken y Yami combinaban sus estilos de forma demoledora contra la segunda bestia. Ken, envuelto en el plumaje ígneo de su espíritu fénix, quemó los ojos del titán con una ráfaga de fuego sagrado, mientras Yami, deslizándose como una sombra entre las piernas del monstruo, ejecutó un corte ascendente de oscuridad que cercenó limpiamente el brazo derecho de la criatura. Erza Scarlet, con una armadura de placas plateadas y una lanza de choque, se arrojó desde la muralla para ensartar el corazón de la Rata Ogro con una estocada terminal.",
        "La tercera bestia fue derribada por Gerald y Vlad, quienes decapitaron a la aberración en una maniobra sincronizada de mandoble y espada bastarda.",
        "Un grito de triunfo comenzó a elevarse entre los estudiantes. Tres de las cuatro pesadillas yacían muertas en el patio y la última retrocedía acribillada por los tiradores de Lighter y Hoshina.",
        "Habíamos revertido el golpe. La victoria parecía al alcance de la mano.",
        "Y fue exactamente en ese instante de alivio cuando el verdadero horror se desató.",
        "* * *",
        "*¡¡BOOOOOOMMMM!!*",
        "La explosión no provino de los muros exteriores.",
        "Provino del corazón mismo de la fortaleza.",
        "Un estruendo sordo y subterráneo sacudió el suelo del patio interior, justo debajo de la capilla y las bóvedas inferiores donde se encontraban los almacenes médicos y el hospital de campaña.",
        "Los adoquines del centro del recinto se hundieron en un socavón humeante de diez metros de diámetro. Desde las profundidades de la brecha no brotó tierra común, sino una densa, asfixiante y verdosa marea de **Gas de Plaga**: el infame *Viento de Piedra Bruja* de los monjes Skaven.",
        "—¡COF... COF! ¡¡ES VENENO!! —gritó un cadete de la guardia médica mientras caía al suelo convulsionando con espuma sanguinolenta en la boca.",
        "El humo corrosivo comenzó a expandirse con rapidez aterradora, disolviendo las vendas, carcomiendo la armadura de los soldados y amenazando con intoxicar a las decenas de estudiantes heridos que yacían en las camillas.",
        "Y entre la niebla venenosa, trepando por el cráter con una velocidad inhumana, emergieron decenas de sombras cubiertas de harapos pestilentes: Monjes de la Plaga armados con mayales de incienso tóxico y Acechadores de la Muerte con guadañas envenenadas.",
        "Al frente de ellos, sobre una roca humeante, se alzó un sacerdote de guerra Skaven, empuñando un báculo coronado por una calavera de carnero que emitía pulsos mágicos verdosos.",
        "—¡Chillad-morid, humanos ciegos! —graznó el brujo con una risa sibilante—. ¡La muralla era el cebo-distracción! ¡Este nido arde-muere desde adentro!",
        "* * *",
        "El terror se apoderó de las filas.",
        "La situación había pasado de una defensa exitosa a una pesadilla de dos frentes.",
        "Por la brecha exterior del portón destrozado, miles de tropas de refresco Skaven volvían a cargar alentadas por la explosión interna. Y en la retaguardia, donde estaban nuestros compañeros heridos, las enfermeras y los suministros de agua y pociones, el veneno y los asesinos amenazaban con masacrarlo todo en cuestión de minutos.",
        "—¡Maldita sea! ¡Nos engañaron! —rugió Gerald, teniendo que quedarse en la puerta principal junto a una docena de veteranos para evitar que la marea exterior inundara el patio.",
        "—¡Freya! ¡Elysea! —el grito de Kaoruko me desgarró el pecho.",
        "Ambas chicas estaban atrapadas en el ala médica, intentando levantar barreras de viento purificador con sus báculos mientras los monjes de la plaga avanzaban hacia ellas con los mayales en alto.",
        "Vlad apretó los dientes, evaluando el tablero con una angustia controlada:",
        "—No podemos dividir la línea exterior sin que el frente caiga... pero si el hospital colapsa, el veneno asfixiará a toda la fortaleza en menos de cinco minutos.",
        "Kaoruko me miró, con el rostro pálido pero la mirada ardiendo con una determinación feroz.",
        "—Seier...",
        "—Nosotros vamos a la retaguardia —le respondí, y el fuego y la tierra en mi interior comenzaron a rugir como un volcán a punto de reventar—. Ken, Yami, Vlad... ¡mantengan esa maldita puerta! ¡Nosotros limpiaremos el nido de adentro!",
        "La noche se había convertido en un infierno de veneno y acero.",
        "Y cada segundo que pasaba, el aire se volvía más escaso y mortal."
    ]
}

def compilar_capitulo_43():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 43              ")
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
        CAPITULO_40, CAPITULO_41, CAPITULO_42, CAPITULO_43
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 43
    txt_cap43 = os.path.join(capitulos_dir, "43_Capitulo_43.txt")
    with open(txt_cap43, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_43['header'].upper()}: {CAPITULO_43['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_43['paragraphs']))
    print(f"[OK] Texto plano Capítulo 43 guardado en:\n  {txt_cap43}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 43
    docx_cap43 = os.path.join(capitulos_dir, "43_Capitulo_43.docx")
    create_single_chapter_document(CAPITULO_43, docx_cap43, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 43 guardado en:\n  {docx_cap43}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 43)
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
    compilar_capitulo_43()
