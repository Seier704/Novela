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

CAPITULO_48 = {
    "header": "Capítulo 48",
    "title": "Honor, Cenizas y Vientos de Descanso",
    "illustrations": [
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "kaoruko_waguri.png"),
            "caption": "Ilustración: Kaoruko Waguri (La Calidez Tras la Tormenta)"
        },
        {
            "image": os.path.join(NOVELA_DIR, "imagenes", "lortelle_kecheln.png"),
            "caption": "Ilustración: Lortelle Kecheln (Recuperada y Lista para el Verano)"
        }
    ],
    "paragraphs": [
        "El silencio que siguió a la tempestad fue el más denso que jamás hubiera respirado en mi nueva vida.",
        "Durante los tres días posteriores a la carnicería en el cañón de la Roca Carmesí, el hedor a azufre y sangre negra fue sustituido poco a poco por el aroma amargo del incienso purificador de la corte y la madera de pino ardiendo en las laderas.",
        "La guerra había terminado, pero el precio del acero siempre se cobra por adelantado.",
        "En la explanada exterior, más de cuarenta piras funerarias se alzaban hacia el cielo gris de la mañana. Guardias veteranos de la guarnición, centinelas fronterizos y catorce cadetes de diversas clases que no alcanzaron a ver la llegada de los jinetes del León descansaban envueltos en sudarios de lino con el estandarte imperial cubriendo sus pechos.",
        "Se celebró una vigilia solemne. El Instructor Jefe Gerald, con el brazo izquierdo vendado y el rostro sombrío como una losa de granito, clavó su espada de oficial en la tierra mientras los clarines de la caballería real entonaban el toque de retreta fúnebre.",
        "A su lado, Vlad y yo nos mantuvimos en posición de firmes. Tomé un puñado de tierra arcillosa y cenizas de la montaña, esparciéndolas sobre el fuego sagrado como solía hacer mi padre Ronald Silva en los desfiladeros del sur.",
        "—Descansen en paz, hermanos de acero —murmuré con voz ronca—. Su guardia ha terminado.",
        "Kaoruko, vestida con una sobreveste blanca y los ojos cristalizados por el dolor de la pérdida, entonó una plegaria dulce y melodiosa con su maná de agua, haciendo descender una llovizna suave que purificaba el aire de las piras. Detrás de ella, Freya Skuld golpeó rítmicamente el borde de su escudo de roble con su lanza tribal, entonando un cántico bárbaro de la Costa Helada para guiar las almas hacia los salones celestiales del norte, mientras Rukia Kuchiki congelaba diminutos capullos de flores silvestres para depositarlos al pie de cada monumento.",
        "Fue un duelo compartido sin distinción de clases nobles ni plebeyas: allí no había alumnos de la A ni desterrados de la F. Solo camaradas que habían sangrado en el mismo barro.",
        "* * *",
        "Al quinto día, los carruajes de oro y terciopelo del Alto Mando Real y los emisarios de la Corona llegaron finalmente a la fortaleza escoltados por una centuria de coraceros de Solaria.",
        "El Gran Patio de Armas de la Roca Carmesí, despejado de escombros y engalanado con los pabellones dorados del Reino, fue el escenario de la ceremonia oficial de condecoraciones y recompensas.",
        "Un heraldo de la corte, portando un pergamino con el gran sello de cera escarlata de la Cancillería Imperial, alzó la voz ante toda la guarnición formada en cuadro:",
        "—¡Por decreto de Su Majestad el Rey y en reconocimiento al heroísmo supremo demostrado en la defensa del bastión fronterizo frente a la plaga subterránea!",
        "Las condecoraciones comenzaron a entregarse bajo el redoble marcial de los tambores.",
        "El **Instructor Principal Gerald** recibió la *Gran Cruz de la Frontera*, consagrando su veteranía de tres décadas al servicio del reino. **Vlad** fue galardonado con la *Estrella de Plata del Ducado*, siendo reconocido formalmente por los mariscales de la capital con el rango de Capitán Táctico de Campaña por haber organizado el cuello de botella que salvó cientos de vidas.",
        "Luego llegó el turno de la Clase F-2.",
        "La multitud de nobles y oficiales cortesanos contuvo el aliento cuando los nombres de los marginados resonaron en la explanada con honores que jamás en la historia de la Academia se habían otorgado a la clase más baja.",
        "A **Caesar King, Lighter, Piper Wheel, Lucy, Burnice White y Pulchra** —los Hijos de Calydon, reconocidos ahora de forma definitiva e irrevocable como estudiantes y combatientes de pleno derecho de la Clase F-2— se les concedió la *Medalla del Valor Cívico* y una gratificación colosal de tres mil monedas de oro del tesoro real para la reconstrucción de sus talleres y la mejora de sus equipos mecánicos.",
        "—¡Miren nada más! ¡Oro de verdad para gasolina y dinamita! —exclamó Lucy con una sonrisa altiva, mordiendo una de las monedas mientras Caesar reía con orgullo y Lighter chocaba los puños con Piper.",
        "A **Ken, Erza Scarlet, Akame, Yami Sukehiro, Hoshina y Natsuki** se les impuso la *Cruz al Mérito Marcial de Primera Clase*. Natsuki, con el pecho inflado y las mejillas rojas de emoción, casi rompe a llorar al sentir el peso del metal reluciente sobre su chaqueta; Yami, fiel a su estilo con apenas dieciséis años pero con el descaro de siempre, aceptó su medalla sosteniendo su cigarrillo apagado en la mano libre con un asentimiento relajado, mientras Akame miraba discretamente hacia la cocina preguntando en voz baja si la ceremonia incluía raciones dobles de asado.",
        "* * *",
        "Y entonces, el heraldo pronunció mi nombre.",
        "—¡Seier Silva! ¡Heredero de la Casa Baronal Silva!",
        "Di tres pasos al frente con mi paso pesado y firme. Los noventa kilos de músculo compacto y las marcas de hollín en mi rostro hacían un contraste salvaje frente al terciopelo pulcro de los aristócratas enviados desde la capital.",
        "El Comandante de la Orden del León se adelantó en persona, portando un cofre de terciopelo morado.",
        "—Seier Silva —habló el veterano caballero con una voz profunda que denotaba un respeto marcial genuino—. Quien corta la cabeza de la colmena salva al reino entero. Vuestra sangre proviene de héroes que nunca doblaron la rodilla, y hoy habéis demostrado que el fuego de la Casa Silva arde más implacable que nunca.",
        "Me prendió sobre el pecho la máxima distinción otorgada a un cadete en cincuenta años: la **Gran Medalla del León de Oro con Hojas de Roble**.",
        "Pero la recompensa no se limitó al metal.",
        "—Por orden real —continuó el heraldo desenrollando el edicto con letras doradas—, la Corona restituye a la Casa Silva la titularidad y recaudación de las canteras de hierro de la Cordillera Sur, con una asignación perpetua de diez mil monedas de plata anuales y el título hereditario de Caballero Defensor de la Frontera.",
        "Un murmullo de impacto recorrió a los cortesanos.",
        "Apreté el medallón de bronce de mis padres que descansaba bajo mi camisa. Las canteras que el Marqués Delacroix había intentado arrebatarnos con su traición volvían a mis manos por la fuerza de mis propios puños. La justicia tardaba, pero comenzaba a cobrarse en acero y títulos legítimos.",
        "Antes de romper filas, una presencia dorada y elegante se deslizó entre los oficiales de la retaguardia.",
        "**Lortelle Kecheln**, con su uniforme impecable de la Clase C y una capa de terciopelo verde esmeralda con ribetes de piel, subió al estrado con una sonrisa que desarmó a toda la comitiva.",
        "—La Corona ha cumplido con su deber... y la **Compañía Mercantil Elte** jamás se queda atrás en saldar sus cuentas —declaró Lortelle con una voz melodiosa que acaparó todas las miradas.",
        "Dos de sus asistentes depositaron un cofre de hierro reforzado que requirió cuatro brazos para cargarlo. Al abrirse, el resplandor de decenas de lingotes de oricalco puro y gemas mágicas condensadas dejó mudos a los presentes.",
        "—Una donación especial para los fondos de armamento y suministros vitalicios de la Clase F-2 —dijo la chica mercader, guiñándome un ojo de forma sutil mientras me entregaba un pagaré sellado con cera dorada—. Y para ti, Silva... una línea de crédito ilimitada en cualquiera de nuestras sucursales del continente. Mis inversiones siempre reciben el mejor trato.",
        "A pocos metros, Kaoruko cruzó los brazos bajo su pecho haciendo un puchero adorable que me arrancó una media sonrisa, mientras Elysea soltaba un carraspeo aristocrático fingiendo desinterés mientras ajustaba los guantes de su estoque.",
        "* * *",
        "Los días siguientes fueron una marcha triunfal de regreso hacia la capital.",
        "Con los túneles Skaven volados y sellados con cargas alquímicas de Agnes y los 400 caballeros asegurando los desfiladeros, los escuadrones mixtos emprendieron el retorno a lomo de carromatos y caballos descansados.",
        "Cuando cruzamos las gigantescas murallas blancas de Solaria dos semanas después, el espectáculo nos dejó sin habla.",
        "Las campanas de la catedral repicaban a fiesta. Miles de ciudadanos abarrotaban las avenidas principales, arrojando pétalos de flores y vitoreando a los «Héroes de la F-2» y a los cadetes de la frontera. Los nobles que meses atrás nos miraban con desprecio en las tabernas debían apartarse en las aceras para contemplar el paso victorioso de nuestra columna.",
        "El semestre había concluido oficialmente.",
        "Las semanas posteriores transcurrieron con una serenidad casi irreal dentro de los muros de la Academia. Las heridas terminaron de cerrar bajo los ungüentos de Agnes y los cuidados de la enfermería; los exámenes prácticos fueron convalidados con honores máximos para todos los supervivientes y las aulas vacías comenzaron a oler al polvo del fin de curso.",
        "El verano había llegado al reino.",
        "El solsticio traía consigo el ansiado periodo de **vacaciones de mitad de año**: un receso completo de seis semanas libres de uniformes, alarmas de madrugada y olor a pólvora rúnica.",
        "* * *",
        "La tarde previa a la apertura de las puertas de la Academia, me encontraba sentado en el pretil de piedra de los jardines interiores de la F-2, limpiando el sudor de mi último entrenamiento con una toalla de lino.",
        "El aire cálido del estío mecía las copas de los robles.",
        "—¿En qué tanto piensas, campeón condecorado? —la voz dulce y risueña de Kaoruko me sacó de mis pensamientos.",
        "Venía caminando sobre el césped con un vestido ligero de verano de color vainilla y una canasta de mimbre recién horneada que desprendía un aroma embriagador a tarta de manzanas y canela. Se sentó a mi lado, apoyando con naturalidad su cabeza sobre mi hombro fornido.",
        "—En que hace cuatro meses creía que esta academia solo era un nido de víboras —le respondí, rodeando su cintura con mi brazo derecho y atrayéndola contra mi pecho—. Y ahora tengo un aula llena de hermanos... y a la mujer más hermosa del imperio a mi lado.",
        "Kaoruko se sonrojó intensamente, enterrando el rostro en mi cuello con un gemido tierno de pura felicidad.",
        "—Seier-kun... eres un tramposo diciendo esas cosas —susurró, besando la curva de mi mandíbula antes de abrir la canasta—. Por cierto... las vacaciones empiezan mañana. Mi familia en el pueblo fronterizo de Valenwood envió una carta... quieren conocer al hombre que salvó mi vida y del que no paro de hablar. Tienen manzanos, ríos cristalinos y una cocina enorme. ¿Vendrías conmigo unas semanas?",
        "Antes de que pudiera responderle, un par de pasos seguros sobre la grava interrumpieron la escena.",
        "Lady Elysea de Vane se acercó vestida con una elegante túnica de viaje de seda celeste, con su cabello plateado recogido en una trenza lateral que acentuaba su cuello esbelto.",
        "—No monopolices tanto a nuestro líder, Waguri —dijo la elfa con una sonrisa altiva pero teñida de un brillo afectuoso evidente—. La Casa de Vane posee una villa en los lagos glaciares del norte. Aguas termales privadas, sparrings de esgrima de alto nivel y el mejor vino de escarcha del continente. Silva tiene una invitación formal firmada por mi padre.",
        "—¡Ni se te ocurra, Lady Hielo! —saltó una tercera voz con una carcajada pícara desde el arco de glicinas.",
        "Lortelle Kecheln apareció abanicándose con indolencia, luciendo un sombrero de paja de ala ancha y gafas oscuras de cristal ahumado.",
        "—Silva pasará la primera quincena en la costa sur con la delegación de la Compañía Elte. Tenemos barcos privados, banquetes de mariscos y muchos negocios pendientes que revisar en mi camarote privado... con llave por dentro, por supuesto.",
        "Kaoruko infló los mofletes con un puchero fiero, Elysea arqueó una ceja con dignidad aristocrática y Lortelle sonrió con esa malicia deliciosa que siempre prometía peligro.",
        "Desde el otro extremo del patio, los gritos de Ken y Lighter encendiendo una gigantesca barbacoa de carne de jabalí junto a Caesar, Piper, Burnice, Natsuki y Hoshina llenaron el aire de risas, humo blanco y cánticos festivos.",
        "Miré hacia el cielo azul despejado.",
        "Las batallas, los Skavens y las conspiraciones de la corte estarían esperando más adelante. Pero por ahora, el calor del verano, la carne asada y la compañía de quienes amaba eran el único combate que valía la pena librar.",
        "Sonreí por primera vez en mucho tiempo con una paz absoluta en el pecho.",
        "Las vacaciones habían comenzado."
    ]
}

def compilar_capitulo_48():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 48              ")
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
        CAPITULO_45, CAPITULO_46, CAPITULO_47, CAPITULO_48
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 48
    txt_cap48 = os.path.join(capitulos_dir, "48_Capitulo_48.txt")
    with open(txt_cap48, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_48['header'].upper()}: {CAPITULO_48['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_48['paragraphs']))
    print(f"[OK] Texto plano Capítulo 48 guardado en:\n  {txt_cap48}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 48
    docx_cap48 = os.path.join(capitulos_dir, "48_Capitulo_48.docx")
    create_single_chapter_document(CAPITULO_48, docx_cap48, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 48 guardado en:\n  {docx_cap48}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 48)
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
    compilar_capitulo_48()
