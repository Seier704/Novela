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

CAPITULO_46 = {
    "header": "Capítulo 46",
    "title": "El Rugido de Calydon",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "hijos_de_calydon.png"),
            "caption": "Ilustración: Los Hijos de Calydon (Operación de Rescate)"
        }
    ],
    "paragraphs": [
        "El suelo temblaba con un ritmo pesado y cadencioso, anunciando la llegada del amo absoluto del nido.",
        "Montado sobre una abominación cuadrúpeda cubierta de púas de hierro y colmillos envenenados, el gigantesco **Señor de la Guerra Skaven** detuvo su avance a escasos cincuenta metros de la entrada destrozada de la fortaleza. Su coraza de bronce brujo irradiaba un vapor verdoso y en su garra derecha sostenía una alabarda rúnica de tres metros cuya hoja goteaba ponzoña.",
        "Sin embargo, lo que detuvo en seco el corazón de todos los defensores no fue el tamaño del monstruo.",
        "Fue lo que traían detrás de él.",
        "Un contingente de Acechadores Skavens emergió de una zanja lateral arrastrando a seis prisioneros encadenados por el cuello con cables de púas. Eran cadetes y personal auxiliar capturados durante el colapso del ala médica y los almacenes subterráneos.",
        "Y entre ellos, de rodillas sobre el barro pero con la mirada ardiendo de orgullo y desprecio, estaba **Lortelle Kecheln**.",
        "Su chaqueta dorada de la Clase C estaba rasgada, un hilo de sangre seca le manchaba la comisura de los labios y tenía las muñecas atadas a la espalda. Detrás de ella, un verdugo Skaven jorobado apoyaba una pesada cuchilla mellada sobre la curva de su cuello.",
        "Una ola de frío visceral me recorrió la espina dorsal.",
        "Sentí cómo el aire desaparecía de mis pulmones. La chica que apenas unas horas atrás había dormido acurrucada en mi pecho, buscando refugio en mis brazos y confesándome con lágrimas en los ojos sus miedos más profundos, estaba a un solo tajo de ser ejecutada frente a nosotros.",
        "La sangre me ardió en las sienes con una furia volcánica. El maná de fuego en mis venas amenazó con estallar y mis piernas se tensaron por instinto para lanzarme hacia adelante como una fiera.",
        "—¡Seier, no! —la mano firme y suplicante de Kaoruko se clavó en mi antebrazo, reteniéndome con desesperación—. ¡Si cargas de frente, la degüellan antes de que des tres pasos!",
        "Apreté la mandíbula con tanta fuerza que el sabor metálico de la sangre llenó mi boca.",
        "Kaoruko tenía razón. En una situación de rehenes contra un enemigo sádico, la desesperación era el billete directo a la morgue. Cerré los ojos una fracción de segundo, canalizando la disciplina gélida del Jeet Kune Do y el control mental del Muay Thai: *Calma. Respira. Piensa. La ira ciega mata; la precisión rescata.*",
        "—¡Rendid-entregad el bastión, gusanos de superficie! —bramó el Señor de la Guerra con una voz que sonaba como grava triturada bajo una rueda de hierro—. ¡O la cachorra de la Compañía Elte y sus amigos morirán chillando-sangrando aquí mismo!",
        "Gerald y Vlad intercambiaron miradas sombrías. Ninguno de los dos sabía cómo actuar sin provocar una masacre inmediata.",
        "Nadie tenía una salida.",
        "Nadie... excepto las sombras.",
        "* * *",
        "En lo alto de la cornisa de piedra que bordeaba el flanco izquierdo de los captores, una silueta felina se movió sin emitir el más mínimo sonido.",
        "Era **Pulchra**.",
        "La ágil y letal chica felina, aliada incondicional de los Hijos de Calydon y de la Clase F-2, se deslizaba por la roca como el humo en la niebla. Sus orejas triangulares se movían captando cada respiración y el brillo ámbar de sus ojos rasgados calculaba los ángulos con una frialdad matemática.",
        "Sostenía dos dagas curvas gemelas invertidas en sus manos.",
        "Sin que el Señor de la Guerra ni los guardias de primera línea se percataran, Pulchra se descolgó boca abajo por una saliente de basalto, cayendo a espaldas del primer guardia con la suavidad de una pluma.",
        "*Chik.*",
        "Un corte limpio y transversal atravesó la laringe del primer Skaven. Pulchra sostuvo el cuerpo inerte con un brazo, amortiguando la caída de su armadura, y antes de que el cadáver tocara el fango, ya se había deslizado detrás del segundo captor.",
        "*Chik... chik.*",
        "Dos gargantas más fueron abiertas en un mutismo aterrador. Pulchra danzaba entre los captores de retaguardia como un espectro invisible, eliminando a los carceleros uno tras otro sin darles tiempo de soltar un solo quejido.",
        "Sin embargo, el verdugo principal que sostenía a Lortelle comenzó a impacientarse ante el silencio de la fortaleza.",
        "—¡No hay trato-acuerdo! —chilló la rata gigante, levantando la pesada cuchilla a dos manos—. ¡Empieza la sangre!",
        "La hoja descendió a toda velocidad hacia el cuello de Lortelle.",
        "Lortelle cerró los ojos con los dientes apretados, negándose a mostrar miedo.",
        "Y entonces, el infierno mecánico cayó del cielo.",
        "* * *",
        "*¡¡¡KRRRAAAAAAANNNNNGGGGGG!!!*",
        "Una lluvia de chispas incandescentes iluminó la noche.",
        "El golpe de la cuchilla no tocó la piel de la chica; se estrelló con un impacto ensordecedor contra una gigantesca rueda-hacha dentada de metal reforzado, dos veces más grande que la persona que la empuñaba.",
        "—¡A las chicas bonitas no se las toca sin pedir permiso, rata mugrienta! —bramó una voz chillona y llena de chulería.",
        "Era **Piper Wheel**.",
        "La pequeña mecánica y guerrera de los Hijos de Calydon había aterrizado como un meteorito justo delante de Lortelle, deteniendo el tajo del verdugo con la mole de su arma giratoria. El motor de su hacha de combustión rugió a máxima potencia: los dientes de acero comenzaron a rotar a miles de revoluciones por minuto, partiendo la cuchilla del Skaven en mil pedazos y rebanando el pecho del verdugo en una explosión de humo y vísceras.",
        "Pero no venía sola.",
        "—¡¡ABRAN PASO A LA REINA!! —resonó un grito de guerra imponente.",
        "Desde la nube de polvo, **Caesar King** cargó en primera línea como un tren blindado desbocado. Con su gigantesco escudo de torre reluciendo bajo el resplandor de las llamas y su gladius en alto, Caesar embistió con toda su fuerza contra el costado de la montura del Señor de la Guerra.",
        "*¡¡¡BOOOOOOMMMM!!!*",
        "El choque de titanio contra hueso fue colosal. Caesar empujó a la mole acorazada y a su jinete más de diez metros hacia atrás, barriendo a decenas de guardias y abriendo una brecha gigantesca en las líneas enemigas.",
        "Casi al mismo tiempo, **Lucy** apareció sobre un montículo de escombros, blandiendo su bate de béisbol envuelto en llamas y lanzando una andanada de proyectiles explosivos que estallaron en el suelo alrededor de los rehenes, creando una cortina de humo y fuego protector.",
        "—¡Corran ahora, idiotas! ¡No los cubriremos dos veces! —gritó Lucy con su habitual arrogancia.",
        "Pulchra cortó las cadenas de Lortelle y de los demás prisioneros en un segundo, empujándolos con agilidad hacia nuestras líneas.",
        "—¡Vuelvan al portón! —ordenó la felina mientras cubría la retirada con giros mortales de sus dagas.",
        "* * *",
        "Lortelle corrió hacia nosotros a trompicones, respirando con dificultad.",
        "Di tres zancadas hacia adelante y la atrapé en el aire, rodeándola con mis brazos y atrayéndola contra mi pecho con una fuerza desesperada. Sentí su corazón latiendo como un tambor desbocado contra mi esternón.",
        "—¿Estás bien? —le pregunté en voz baja, con un nudo en la garganta mientras le retiraba el cabello empapado de la frente.",
        "Lortelle levantó la mirada hacia mí. Sus ojos verdes estaban empañados de alivio, pero su sonrisa pícara y hermosa no tardó en asomar entre el hollín de sus mejillas.",
        "—Te dije... que la Compañía Elte siempre vuelve a sus inversiones, Silva —susurró con voz temblorosa, aferrándose un segundo a mi solapa antes de que Kaoruko y Elysea la tomaran de los brazos para escoltarla hacia la seguridad de los muros.",
        "Kaoruko me miró a los ojos, asintiendo con una expresión decidida que no necesitaba palabras.",
        "Giré lentamente sobre mis talones.",
        "El alivio se transformó en un frío glacial. Toda la contención, la paciencia y el dolor contenido de las últimas horas se condensaron en el centro de mi pecho en una furia pura, concentrada y letal.",
        "* * *",
        "Al frente, el Señor de la Guerra Skaven se reincorporaba furioso sobre su montura, bramando maldiciones mientras miles de ratas volvían a apiñarse a su alrededor.",
        "Pero el tablero de batalla acababa de cambiar de forma definitiva.",
        "Los rehenes estaban a salvo.",
        "Los Hijos de Calydon —Caesar, Piper, Lucy y Lighter— estaban desplegados en formación de vanguardia. Vlad sostenía su espada bastarda lista para la orden de carga. Yami sonreía con su katana oscura desenvainada, Ken ardía en llamas y Subaru ajustaba su guardia con su aura plateada al máximo.",
        "Y lo más importante de todo: la mente colmena de los Skavens dependía exclusivamente de la bestia que teníamos enfrente.",
        "Aquellas criaturas eran cobardes por naturaleza, movidas únicamente por el terror a su amo. Si el Señor de la Guerra caía decapitado, la mente colectiva del enjambre colapsaría en cuestión de segundos y los miles de asaltantes correrían en pánico por sus vidas hacia las profundidades.",
        "—¡TODOS A LA VEZ! —bramé con un rugido que hizo vibrar el aire del cañón, canalizando todo el maná de tierra y fuego en mis puños y piernas—. ¡¡¡CORTEN LA CABEZA DE LA COLMENA!!!"
    ]
}

def compilar_capitulo_46():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 46              ")
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
        CAPITULO_45, CAPITULO_46
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 46
    txt_cap46 = os.path.join(capitulos_dir, "46_Capitulo_46.txt")
    with open(txt_cap46, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_46['header'].upper()}: {CAPITULO_46['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_46['paragraphs']))
    print(f"[OK] Texto plano Capítulo 46 guardado en:\n  {txt_cap46}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 46
    docx_cap46 = os.path.join(capitulos_dir, "46_Capitulo_46.docx")
    create_single_chapter_document(CAPITULO_46, docx_cap46, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 46 guardado en:\n  {docx_cap46}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 46)
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
    compilar_capitulo_46()
