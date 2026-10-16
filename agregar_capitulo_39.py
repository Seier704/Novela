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

CAPITULO_39 = {
    "header": "Capítulo 39",
    "title": "El Refugio más Cálido",
    "paragraphs": [
        "El resplandor dorado del pequeño fuego de Lortelle proyectaba sombras alargadas y danzantes sobre los muros de arenisca de la gruta.",
        "Afuera de nuestra pequeña hendidura, el viento helado del cañón subterráneo seguía aullando con su lamento mineral, arrastrando el murmullo lejano de las cascadas gélidas y el goteo incesante de las estalactitas. Pero dentro de aquel rincón oculto, la losa de pizarra calentada por las profundidades y la proximidad de nuestros cuerpos creaban un oasis de calor que parecía desafiar las leyes de aquel mundo hostil.",
        "Apoyé la espalda contra la roca pulida, acomodando el lecho de arena seca debajo de nosotros.",
        "Lortelle, todavía temblando levemente por los últimos vestigios del frío, se movió con lentitud. En lugar de apartarse a un rincón, buscó mi presencia con una necesidad instintiva de abrigo. Deslizó sus piernas desnudas a lo largo de mi regazo, doblando las rodillas para acomodarse sobre mis muslos, y apoyó su mejilla y su sien directamente sobre el centro de mi pecho desnudo.",
        "Para protegerla de las corrientes traicioneras, pasé mis dos brazos alrededor de su figura menuda.",
        "La rodeé con una firmeza protectora, atrayéndola por completo contra mi torso y ciñendo los bordes de la capa gruesa sobre sus hombros para sellar cualquier fuga de aire. Esperaba que la joven heredera de la Compañía Elte hiciera algún comentario sarcástico o intentara recomponer su dignidad nobiliaria.",
        "Pero Lortelle no protestó.",
        "Al contrario: soltó un largo suspiro tembloroso que le recorrió el cuerpo entero, relajando la tensión de sus músculos como si una carga insoportable hubiera sido retirada de sus hombros de golpe. Sus pequeñas manos delgadas se aferraron a mi cintura y su cuerpo menudo se moldeó contra el mío con una entrega absoluta.",
        "* * *",
        "El silencio de la caverna se volvió tan íntimo que solo se escuchaba el crepitar manso de la llama mágica y el latido fuerte y constante de mi corazón bajo su oído.",
        "En la penumbra, la voz de Lortelle rompió la calma con un tono tan bajo, vulnerable y desarmado que parecía un susurro arrancado de lo más hondo de su ser:",
        "—¿Sabes, Seier?... —murmuró ella, con la mirada perdida en las brasas titilantes—. Esto es lo más cálido que he sentido en toda mi vida.",
        "No la interrumpí. Dejé que mis manos descansaran sobre su espalda, sintiendo la suavidad de su piel y el calor reconfortante que comenzaba a florecer entre los dos.",
        "—Desde que era una niña pequeña en los distritos mercantiles de la capital... jamás conocí lo que era estar en paz —continuó Lortelle, y por primera vez, su voz de dieciséis años perdió todo rastro de arrogancia y fingimiento—. En mi familia no hay padres ni hermanos; solo hay socios comerciales, contratos y rivales de sangre. Mi propio tío intentó envenenar mi reputación cuando apenas tenía diez años para quedarse con mi herencia... y las pocas personas a las que alguna vez llamé amigas terminaron vendiendo mis secretos por un puñado de monedas de plata.",
        "Apretó sus pequeños dedos contra mi piel, hundiendo un poco más el rostro en mi clavícula.",
        "—Siempre me han traicionado, Seier. Cada persona que alguna vez me miró a los ojos y me dijo que cuidaría de mí... terminó poniéndome un precio en la espalda. Aprendí a calcular, a sonreír con veneno y a no dormir jamás con los dos ojos cerrados porque sabía que la menor debilidad significaba mi ruina.",
        "Lortelle levantó despacio la cabeza.",
        "Sus ojos verde esmeralda estaban húmedos por un brillo cristalino y sus labios entreabiertos temblaban con una emoción pura.",
        "—Y de pronto... caigo a un abismo de cien metros, y un chico al que apenas conocía de vista se tira al vacío conmigo. Me salvas la vida, recibes los golpes de las rocas para que yo no me rompa, me cargas durante horas en la oscuridad y ahora me tienes aquí entre tus brazos, dándome tu propio calor sin pedirme una sola moneda, ni un favor, ni una promesa...",
        "Lortelle me miró con una devoción tan intensa y sincera que sentí un vuelco en el pecho.",
        "—Estar contigo... se siente tan seguro que casi me hace olvidar todo el dolor que viví antes.",
        "Levanté mi mano derecha y le acaricié la mejilla con el pulgar, secando una pequeña lágrima que comenzaba a resbalar hacia su barbilla.",
        "—En mi mundo me enseñaron que un verdadero guerrero usa su fuerza para proteger a los suyos, Lortelle —le respondí en voz baja, sosteniéndole la mirada—. No necesitas pagar nada por estar viva. Y mientras estés conmigo, nadie va a traicionarte.",
        "Una pequeña sonrisa, la más dulce, hermosa y despojada de malicia que jamás había mostrado, floreció en los labios de la pelirroja.",
        "—Eres un tonto, Seier Silva... —susurró con ternura inmensa—. Pero es el mejor tipo de tonto que he conocido.",
        "Se acurrucó de nuevo contra mi pecho, escondiendo el rostro bajo mi mentón y cerrando los ojos.",
        "En cuestión de minutos, el agotamiento físico y el alivio emocional finalmente vencieron sus defensas, y Lortelle cayó en un sueño profundo y tranquilo.",
        "* * *",
        "Las horas de la noche subterránea transcurrieron con una lentitud solemne.",
        "A pesar del cansancio acumulado en mis músculos tras las batallas y el impacto de la caída, no me permití cerrar los ojos.",
        "Mi instinto protector forjado en diez años de supervivencia marcial y disciplina de Muay Thai me mantuvo en un estado de vigilia serena y constante. La oscuridad de las cavernas profundas no era un lugar para descuidos; cualquier cambio en la presión del aire, el menor crujido de gravilla en el cañón exterior o el aroma a azufre de los Skavens activaban mis sentidos al instante.",
        "Mantuve mi espalda apoyada contra la pared caliente, con la mano izquierda apoyada sobre la empuñadura de mi daga de combate y mi brazo derecho sosteniendo con firmeza a la chica sobre mi regazo.",
        "A mitad de la noche, una pequeña pesadilla pareció turbar el descanso de Lortelle.",
        "La muchacha soltó un quejido ahogado en sueños, arrugando el ceño con dolor, y dos lágrimas silenciosas brotaron de sus párpados cerrados, resbalando por su mejilla amoratada como la memoria viva de años de soledad y desconfianza.",
        "La contemplé en silencio bajo la penumbra ámbar.",
        "Detrás de esa fachada de mujer fatal y comerciante astuta de la Clase C, no era más que una joven de dieciséis años que había tenido que madurar a la fuerza en un mundo despiadado que jamás le ofreció un refugio sincero.",
        "Acerqué mi mano libre a su rostro con una suavidad infinita.",
        "Con la yema de los dedos, le sequé con delicadeza las lágrimas que manchaban su mejilla. Luego comencé a acariciarle lentamente los mechones de cabello castaño rojizo, peinándole la coronilla con movimientos rítmicos y pacientes mientras le susurraba palabras de calma casi inaudibles:",
        "—Tranquila... estás a salvo. Estoy aquí.",
        "El tacto cálido de mis manos y el sonido grave de mi voz hicieron su efecto al instante.",
        "El ceño fruncido de Lortelle se relajó poco a poco. Dejó escapar un suspiro dulce y plácido, sonrió suavemente entre sueños y hundió aún más su rostro contra mi pecho, aferrándose a mi piel como si fuera su ancla en mitad de la tormenta.",
        "La pequeña llama mágica sobre las estalagmitas terminó de consumir su maná lentamente, dejando nuestras ropas completamente secas y tibias antes de extinguirse en un hilito de humo dorado.",
        "La caverna quedó iluminada únicamente por el resplandor azulado y lejano de los musgos fosforescentes del exterior.",
        "Permanecí despierto velando su sueño durante el resto de la noche, con el peso dulce de Lortelle sobre mi cuerpo y la certeza absoluta de que el destino acababa de forjar un lazo inquebrantable entre los dos.",
        "El abismo seguía rugiendo afuera en la oscuridad.",
        "Pero mientras amanecía en el corazón de la tierra, la chica que jamás había confiado en nadie dormía en mis brazos sabiendo que jamás volvería a estar sola."
    ]
}

def compilar_capitulo_39():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 39              ")
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
        CAPITULO_35, CAPITULO_36, CAPITULO_37, CAPITULO_38, CAPITULO_39
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 39
    txt_cap39 = os.path.join(capitulos_dir, "39_Capitulo_39.txt")
    with open(txt_cap39, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_39['header'].upper()}: {CAPITULO_39['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_39['paragraphs']))
    print(f"[OK] Texto plano Capítulo 39 guardado en:\n  {txt_cap39}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 39
    docx_cap39 = os.path.join(capitulos_dir, "39_Capitulo_39.docx")
    create_single_chapter_document(CAPITULO_39, docx_cap39, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 39 guardado en:\n  {docx_cap39}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 39)
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
    compilar_capitulo_39()
