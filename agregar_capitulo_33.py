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

CAPITULO_33 = {
    "header": "Capítulo 33",
    "title": "Hogueras en la Frontera",
    "paragraphs": [
        "El convoy militar de carruajes blindados abandonó las murallas doradas de Solaria antes del amanecer, escoltado por batallones de caballeros grifos de la guardia fronteriza.",
        "El trayecto hacia el sur duró dos jornadas completas a través de páramos rocosos y desfiladeros áridos, hasta que el aire comenzó a adquirir ese regusto metálico y sulfuroso tan característico de las zonas con fisuras de maná subterráneo. Al caer la segunda tarde, el convoy coronó la última meseta y el campamento avanzado apareció ante nuestros ojos: el Fuerte de la Niebla Roja.",
        "Era una fortaleza militar excavada directamente en la roca volcánica al pie de la Cordillera Escarlata. Desde sus almenas de basalto negro se divisaba la entrada monumental del calabozo: una grieta gigantesca en la ladera de la montaña, de cuyas fauces oscuras brotaba un vaho carmesí espeso que teñía el horizonte de un tono sanguinolento.",
        "Allí se concentrarían todos los escuadrones mixtos para la última noche de preparativos antes del descenso.",
        "* * *",
        "A pesar de que el alto mando imperial había dispersado formalmente a la Clase F-2 en distintos grupos, en cuanto los carruajes descargaron los suministros en el patio de armas, la separación administrativa demostró ser una ilusión.",
        "Apenas cayó la noche y las grandes hogueras de leña de pino se encendieron en el patio central del fuerte, el espíritu de nuestra clase volvió a converger como un río buscando su cauce natural.",
        "Cerca de los talleres de herrería, el Escuadrón 3 ya estaba dando un espectáculo monumental.",
        "—¡No, no, señor caballero de la B-1! —exclamaba Burnice agitando un soplete encendido frente al rostro pálido de un cadete con armadura ceremonial—. ¡Si le agregamos una válvula de compresión de queroseno rúnico a tu peto, tus estocadas van a salir con un empuje de trescientos caballos de fuerza! ¡Vas a ser como un cometa humano!",
        "—¡P-Pero señorita White, las normas de la orden prohíben modificar las corazas familiares con tanques de combustible inflamable! —tartamudeaba el pobre noble, retrocediendo con los ojos desorbitados mientras Ken, sentado sobre un barril con su bandana roja chueca, se atiborraba de carne asada mientras reía a carcajadas.",
        "—¡Hazle caso, novato! —gritaba Ken con la boca llena—. ¡Si Burnice dice que necesitas fuego en la espalda, es porque necesitas fuego en la espalda!",
        "A unos metros de ellos, en una mesa larga de piedra bajo el alero del cuartel, el Escuadrón 5 mostraba una dinámica completamente distinta.",
        "Los dos nobles de la Clase A-1 que habían sido asignados con Lighter y Subaru Hoshina habían intentado asumir el liderazgo al inicio del viaje con su habitual prepotencia aristocrática. Sin embargo, la lección había sido rápida y silenciosa: bastó que uno de ellos intentara levantarle la voz a Hoshina para que Lighter dejara caer su pesado puño vendado sobre la mesa de roble, partiéndola en dos con un chasquido sordo mientras sus ojos avellana brillaban detrás de sus gafas oscuras con una frialdad de depredador de Calydon.",
        "Ahora, los dos nobles comían en silencio sepulcral al fondo del banco, mientras Lighter y Hoshina compartían una fuente de estofado caliente.",
        "Hoshina le soplaba suavemente a una cucharada antes de ofrecérsela al muchacho con una sonrisa tímida y radiante, ajena por completo a la mirada atónita de los soldados veteranos que jamás habían visto a la temida 'chica de hielo' mostrar semejante devoción hacia un hombre.",
        "* * *",
        "En el otro extremo del patio, la disciplina militar rozaba la perfección marcial.",
        "Erza Scarlet y Akame habían sentado a los otros tres integrantes de su escuadrón en círculo sobre la gravilla. Mientras Akame afilaba su espada con una piedra de agua con movimientos rítmicos e hipnóticos que ponían los pelos de punta a sus compañeros, Erza revisaba cada cantimplora, ración seca y piedra de pedernal con la minuciosidad de un general de asedio.",
        "—Si alguien gasta una sola ración de comida sin autorización previa durante los primeros tres niveles del calabozo —advirtió Erza con voz de hierro—, cargará el fardo de equipo pesado de todo el grupo durante los siguientes cinco kilómetros. ¿Queda claro?",
        "Los tres cadetes asintieron con la cabeza con los rostros pálidos, sin atreverse a pestañear.",
        "Más allá, el drama cómico habitual se desarrollaba junto a las bodegas de intendencia.",
        "—¡CAPITÁN YAMI! ¡ESTA MOCHILA PESA OCHENTA KILOS! —chillaba Subaru Natsuki, tambaleándose bajo una montaña de petates militares y cajas de provisiones—. ¡¿POR QUÉ TENGO QUE CARGAR YO TODO EL TABACÓ Y LAS BOTELLAS DE SAKE DEL ESCUADRÓN?!",
        "Yami, recostado sobre un fardo de heno con los ojos entrecerrados y un cigarrillo humeante entre los labios, expulsó una voluta azulada hacia el cielo nocturno.",
        "—Porque si se rompe una sola botella, mocoso, te usaré de cebo para los gusanos de tierra en el primer pasillo. Considera el peso como tu entrenamiento muscular de la semana.",
        "Renji y Rukia pasaron por el lado cargando paquetes de flechas rúnicas. Al ver a Natsuki en apuros, Renji le dio una patada suave al bulto para equilibrarlo entre risas, mientras Rukia regañaba a su compañero por no tener más empatía con el pobre muchacho.",
        "* * *",
        "En el centro del patio, alrededor de la hoguera más grande, el **Escuadrón 7** se había instalado para cenar.",
        "Freya Skuld estaba en su elemento absoluto. La giganta norteña había clavado una pata entera de buey salvaje sobre las brasas y la cortaba con una daga de caza descomunal, ofreciendo pedazos humeantes a todos los que pasaban por su lado.",
        "—¡Come más, Silva! —bramó Freya, estampándome un pedazo de carne del tamaño de un escudo en el plato de madera con una sonrisa radiante—. ¡Un hombre de noventa kilos no puede sostener la vanguardia contra bestias de caparazón blindado con el estómago vacío!",
        "—No me quejo de la comida, Freya —respondí con una media sonrisa, devorando la carne bien sazonada—. Tiene buena mano tu tribu para los asados.",
        "Dante Vance estaba sentado sobre una caja de municiones, limpiando el cerrojo de su mosquete rúnico mientras bebía de una petaca de sidra.",
        "—Admito que este campamento es mucho más ruidoso de lo que esperaba —comentó el tirador con su tono socarrón habitual, mirando hacia donde Burnice intentaba colocarle una turbina en la espalda a su compañero—. En la Clase B-1 todos están encerrados en sus tiendas rezando a los dioses de la guerra o repasando esquemas teóricos. Ustedes parecen estar preparándose para una boda campesina.",
        "—La risa templa el espíritu antes del peligro, Dante —habló Elysea de Vane con su tono cristalino y elegante.",
        "La elfa de hielo estaba sentada sobre una manta de lana limpia, sosteniendo un cuenco de té que Kaoruko le había servido con esmero. Había dejado a un lado su habitual porte distante y observaba el bullicio de la Clase F con una fascinación sincera en sus ojos celestes.",
        "—En la Clase A nos enseñaron que el deber se cumple con frialdad y jerarquía —continuó Elysea, mirando a Ken y a Burnice riendo junto a las llamas—. Pero ver cómo estos muchachos confían sus vidas entre sí sin importarles los linajes... me hace comprender por qué lograron lo imposible en el coliseo.",
        "Kaoruko, sentada a mi lado muy pegada a mi brazo, sonrió con dulzura.",
        "—La F-2 es como una gran familia, Lady Elysea —dijo Kaoruko con voz suave—. Todos llegamos aquí con cicatrices o siendo considerados menos que los demás. Pero juntos aprendimos que nadie tiene que pelear solo.",
        "Elysea contempló a la chica azabache durante un momento y luego asintió con una reverencia delicada.",
        "—Es una hermosa forma de luchar, Waguri-san. Espero ser digna de esa confianza mañana en las cavernas.",
        "* * *",
        "Avanzada la medianoche, cuando el fuego de las hogueras se redujo a brasas titilantes y la mayoría de los cadetes se retiraron a las tiendas de campaña para dormir, me puse de pie y caminé hacia el mirador exterior del fuerte.",
        "El viento que soplaba desde la garganta de la montaña era gélido, cargado con el olor a ceniza y maná corrompido que brotaba de las profundidades del calabozo.",
        "El crujido de pasos sobre la piedra anunció una presencia familiar.",
        "Me giré y encontré a Kaoruko caminando hacia mí, envuelta en una capa de lana gruesa con capucha. Sus mejillas estaban teñidas por el frío y sus ojos oscuros reflejaban el resplandor de las estrellas de la cordillera.",
        "—Sabía que estarías aquí, Seier-kun —dijo en un murmullo, colocándose a mi lado y apoyando sus antebrazos sobre la barandilla de roca.",
        "—Solo estaba estudiando el terreno —respondí, pasando un brazo sobre sus hombros para resguardarla del viento cortante.",
        "Kaoruko se acurrucó contra mi costado con un suspiro de alivio, apoyando su cabeza en mi pecho y escuchando el latido acompasado de mi corazón.",
        "—Dicen que las Cavernas Carmesíes son muy traicioneras... —susurró ella, apretando la tela de mi chaqueta—. Hay criaturas que atacan en manadas por los techos y el miasma desgasta el cuerpo sin que te des cuenta.",
        "—Lo sé —le dije con calma, levantando su barbilla con un dedo para que me mirara a los ojos—. Pero tenemos a Elysea cubriendo el espacio con hielo, a Dante asegurando el camino con sus trampas, a Freya sosteniendo la línea y a mí en el frente. Y lo más importante... te tengo a ti cubriendo mi espalda.",
        "Kaoruko me miró con una devoción tan pura que el aire frío a nuestro alrededor pareció entibiarse. Sonrió con ternura y se alzó ligeramente sobre las puntas de los pies, depositando un beso suave y prolongado sobre mis labios.",
        "—Siempre estaré a tu lado, Seier-kun —prometió en un susurro contra mi boca—. No importa qué tan profundo sea el abismo.",
        "* * *",
        "El toque de queda del fuerte sonó a la quinta hora del amanecer.",
        "Un cuerno de guerra de sonido grave y pesado cortó la niebla de la mañana, despertando a los escuadrones y marcando el fin de la espera.",
        "En el patio de armas, los cinco integrantes del Escuadrón 7 nos reunimos con el equipo de combate ajustado:",
        "Dante cargaba su mosquete rúnico y sus bolsas de perdigones de viento; Freya sostenía su lanza de trueno y su escudo de roble con una sonrisa fiera; Elysea ajustaba el tahalí de su estoque de hielo con mirada calculadora; Kaoruko revisaba sus frascos de bálsamo purificador; y yo me vendaba los nudillos de hierro mientras sentía el maná de tierra y fuego vibrar en mis venas.",
        "Frente a nosotros, las monumentales puertas de hierro que sellaban la grieta del calabozo comenzaron a elevarse con un chirrido de cadenas que estremeció la montaña.",
        "Una bocanada de niebla carmesí y un aullido subterráneo lejano brotaron del interior de la tierra, invitándonos a entrar.",
        "Miré a mis compañeros de la F-2 que formaban en los otros escuadrones: Ken me dedicó un pulgar en alto; Lighter levantó su mentón con un guiño cómplice junto a Hoshina; Erza alzó su espada; y Yami asintió desde el fondo.",
        "—Escuadrón 7 —ordené, dando el primer paso hacia la penumbra del calabozo—. En marcha.",
        "La oscuridad de las Cavernas Carmesíes nos tragó por completo.",
        "Y la verdadera cacería en las entrañas del imperio acababa de comenzar."
    ]
}

def compilar_capitulo_33():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 33              ")
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
        CAPITULO_30, CAPITULO_31, CAPITULO_32, CAPITULO_33
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 33
    txt_cap33 = os.path.join(capitulos_dir, "33_Capitulo_33.txt")
    with open(txt_cap33, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_33['header'].upper()}: {CAPITULO_33['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_33['paragraphs']))
    print(f"[OK] Texto plano Capítulo 33 guardado en:\n  {txt_cap33}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 33
    docx_cap33 = os.path.join(capitulos_dir, "33_Capitulo_33.docx")
    create_single_chapter_document(CAPITULO_33, docx_cap33, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 33 guardado en:\n  {docx_cap33}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 33)
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
    compilar_capitulo_33()
