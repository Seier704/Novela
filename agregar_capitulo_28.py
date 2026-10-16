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

CAPITULO_28 = {
    "header": "Capítulo 28",
    "title": "Llama y Escarcha",
    "paragraphs": [
        "La noche cayó sobre los dormitorios de la Clase F-2 envuelta en un bullicio festivo y aromático.",
        "Tal como había prometido, Burnice había asaltado las cocinas del complejo para hornear tres bandejas colmadas de pasteles de manzana con canela y azúcar glasé. Ken, con el torso al descubierto y una sonrisa radiante, devoraba los hojaldres de dos en dos mientras recreaba en el aire su patada giratoria contra la alabarda de los Valenwood, arrancando risas y vítores de Natsuki, Renji y Caesar King.",
        "Sin embargo, el artífice del segundo nocaut de la tarde no estaba en el salón.",
        "Aprovechando el alboroto general, Lighter se había escabullido en silencio por las escaleras traseras para buscar la tranquilidad del aire nocturno.",
        "* * *",
        "En la terraza superior del edificio, donde las gárgolas de piedra custodiaban la vista panorámica de la capital iluminada por faroles de aceite, el viento soplaba limpio y fresco.",
        "Lighter estaba apoyado contra la barandilla de mármol, con su chaqueta de cuero entreabierta sobre los hombros y las gafas oscuras descansando en el bolsillo de su pecho. Entre los dedos de su mano derecha, vendada con tela áspera y con los nudillos enrojecidos por el combate, sostenía su viejo chisquero mecánico, haciéndolo chasquear perezosamente contra la brisa: *clic-fshhh*.",
        "El suave crujido de una bota sobre las losas de piedra rompió el murmullo de la noche.",
        "Lighter no se sobresaltó. Sonrió de lado sin girar la cabeza de inmediato.",
        "—Tienes el paso más silencioso de toda la academia, señorita hielo —dijo el boxeador con su tono grave y rasposo—. Si no fuera por el aroma a menta silvestre que dejas al caminar, pensaría que una sombra vino a degollarme.",
        "Subaru Hoshina emergió de la penumbra del arco de la escalera.",
        "Vestía un kimono corto de noche de color azul profundo, atado con un lazo blanco alrededor de su cintura delgada. Su larga melena plateada caía suelta sobre sus hombros, brillando bajo la luz plateada de la luna llena. Entre sus manos sostenía un frasco pequeño de cerámica vidriada y un cuenco de té tibio.",
        "—Un verdadero asesino no usa menta —replicó ella con una voz suave, acercándose con pasos lentos hasta detenerse junto a él en el borde de la terraza—. Y dudo mucho que alguien en esta academia sea capaz de degollarte con la guardia que tienes.",
        "—Nunca se sabe —bromeó Lighter, cerrando la tapa del chisquero con un chasquido metálico—. La complacencia en el ring te cuesta la mandíbula.",
        "Hoshina no respondió a la broma.",
        "Sus ojos verdes, habitualmente afilados y desconfiados, se posaron sobre la mano derecha del muchacho. Con una delicadeza inusual, tomó asiento en el banco de piedra adosado al muro y le hizo un gesto con la cabeza.",
        "—Dame tu mano.",
        "Lighter alzó una ceja, pero no opuso resistencia. Se sentó a su lado sobre la piedra fresca y le extendió el brazo.",
        "Hoshina desató con paciencia el nudo de la venda de algodón. Sus dedos delgados y fríos rozaron la piel áspera y callosa de Lighter, retirando la tela para revelar los nudillos hinchados y los moretones superficiales causados por el impacto contra la coraza de los hermanos Valenwood.",
        "Abrió el frasco de ungüento de árnica y flor de hielo que Kaoruko le había preparado. Tomó una pequeña porción con la yema de los dedos y comenzó a esparcirla sobre la piel lastimada con movimientos circulares y meticulosos.",
        "Lighter se quedó inmóvil. El calor del ungüento combinado con el tacto suave y tembloroso de Hoshina le envió un cosquilleo inesperado a lo largo del antebrazo.",
        "—Hiciste un combate impecable hoy —susurró ella sin levantar la vista de sus manos—. Esa finta por debajo de la alabarda fue... perfecta.",
        "—Los gemelos eran predecibles —dijo él, intentando sonar despreocupado—. Mucho peso, poco juego de cintura.",
        "—Para ti todo parece un simple juego —murmuró Subaru, deteniendo el masaje por un instante mientras sus dedos se entrelazaban con los de él sobre el regazo—. Pero cuando vi esas dos alabardas de acero caer hacia ti al mismo tiempo... sentí que el pecho se me congelaba.",
        "Lighter guardó silencio, observando el perfil delicado de la muchacha bajo la penumbra.",
        "—Durante años, lo único que me importaba en este mundo era proteger a Kaoruko y sobrevivir un día más —continuó Hoshina con un hilo de voz cargado de una vulnerabilidad que jamás había mostrado a nadie—. Todo el mundo fuera de ella me parecía una amenaza. Pero hoy, mientras estabas en la arena... no pensé en la clase, ni en los puntos, ni en el honor. Solo sentí un terror asfixiante de que te lastimaran.",
        "Subaru levantó finalmente la mirada hacia los ojos avellana de Lighter. Sus pupilas brillaban con lágrimas contenidas y un rubor carmesí encendía sus mejillas níveas.",
        "—Te has vuelto demasiado importante para mí, Lighter. Y eso me asusta... me asusta muchísimo.",
        "Lighter sintió un vuelco profundo en el pecho.",
        "El peleador solitario de las Tierras Yermas, el hombre que solo confiaba en el peso de sus puños y en el rugido de los motores de Calydon, sintió cómo toda su coraza de rudeza se desmoronaba en un segundo ante la sinceridad de la chica de hielo.",
        "Con su mano izquierda, Lighter acarició suavemente la mejilla de Subaru. Sus dedos anchos apartaron un mechón plateado detrás de su oreja, sintiendo la piel suave y tibia de la muchacha temblar levemente bajo su toque.",
        "—En los caminos de Calydon aprendí que la gente suele ir y venir, Hoshina —dijo Lighter con una voz baja, tierna y llena de una calidez que solo ella había logrado despertar—. Aprendí a no apegarme a nada que no pudiera defender con los puños. Pero desde la primera vez que te vi en Solaria... supe que no quería que fueras a ninguna parte.",
        "Subaru contuvo el aliento.",
        "—No tienes que tener miedo de nada conmigo —susurró él, acortando la distancia entre sus rostros—. No pienso dejar que nadie rompa la guardia que construimos juntos.",
        "Lighter se inclinó hacia adelante.",
        "Hoshina no retrocedió. Cerró los ojos despacio, dejando escapar un suspiro entrecortado mientras sus manos se aferraban a las solapas de la chaqueta de cuero del muchacho.",
        "Sus labios se encontraron bajo el resplandor plateado de la luna.",
        "Fue un beso lento, dulce y profundo al principio, que pronto se encendió con una intensidad cálida y voraz. El calor de Lighter derritió la última brizna de frialdad y soledad que Hoshina había guardado durante años en su alma; y para él, la suavidad de Subaru fue el primer remanso de paz verdadera que había conocido en toda su vida.",
        "Permanecieron abrazados en el banco de piedra durante un tiempo sin reloj, con la cabeza de ella descansando plácidamente contra el hombro del boxeador y el brazo de Lighter rodeando su cintura con una firmeza protectora, contemplando las estrellas en un silencio perfecto.",
        "* * *",
        "A la mañana siguiente, la academia amaneció envuelta en un clima de festividad.",
        "Por decreto de la dirección imperial, el comité del torneo había otorgado un día completo de descanso obligatorio a todas las clases clasificadas para permitir la recuperación total de maná y heridas antes de las rondas de escuadrón.",
        "El distrito comercial de Solaria se había transformado en un carnaval: puestos de comida ambulante, músicos callejeros y miles de ciudadanos debatiendo apasionadamente sobre las probabilidades de los cuatro semifinalistas: la temible Clase A-1, la calculadora Clase A-2 de Elysea de Vane, la resistente Clase B-1 y la invicta cenicienta del torneo, nuestra Clase F-2.",
        "En el comedor de la academia, el ambiente de nuestro grupo era de una camaradería resplandeciente.",
        "Kaoruko servía tazas de té verde con una sonrisa pícara y radiante que no lograba disimular: había notado de inmediato la complicidad secreta, los roces discretos de manos por debajo de la mesa y las miradas iluminadas que Subaru Hoshina y Lighter intercambiaban desde el desayuno. Ver a su mejor amiga tan en paz y feliz llenaba de dicha el corazón de la chica azabache.",
        "—Seier-kun, míralos... —me susurró Kaoruko al oído mientras me servía un plato de fruta fresca—. Subaru-chan parece otra persona hoy. Sus ojos brillan tan lindo.",
        "—Lighter es un buen tipo —respondí con una sonrisa tranquila, dándole una palmada suave en la espalda a mi amiga—. Supo encontrar la llave correcta para abrir esa puerta.",
        "* * *",
        "Por la tarde, la diversión dio paso a la estrategia militar.",
        "Nos reunimos en el aula de táctica con la pizarra llena de diagramas, runas y notas que Natsuki, Erza y Lucy habían preparado meticulosamente.",
        "—Mañana comienzan los combates de tres contra tres —anunció Erza, golpeando la mesa con los nudillos para captar la atención de todos—. La modalidad de trío es la más brutal del torneo: el espacio de la arena se reduce tácticamente, las sinergias mágicas de área son devastadoras y un solo error de posicionamiento puede dejar a dos compañeros aislados frente a una ofensiva combinada.",
        "—Nuestras cuatro plazas oficiales siguen siendo las mismas: Silva, Erza, Ken y Lighter —explicó Natsuki señalando los cuatro nombres en la cabecera—. Pero para mañana solo podemos enviar a tres de ellos a la arena.",
        "Ken levantó la mano de un salto, completamente recuperado tras la siesta y los pasteles de Burnice.",
        "—¡Yo estoy al cien por ciento! —exclamó con entusiasmo—. ¡Mis llamas están listas para volar!",
        "—Lighter acaba de pelear dos días seguidos y sus nudillos necesitan terminar de desinflamarse para la gran final de cuatro —analizó Erza con criterio militar—. Silva es nuestro ariete indiscutible de choque frontal. Ken es el desorganizador aéreo. Y yo aportaré la cobertura de armaduras y armas pesadas.",
        "Lighter asintió con una sonrisa tranquila, apoyando la espalda contra la pared junto a Subaru Hoshina.",
        "—Me parece el trío perfecto —dijo el boxeador—. Tienen contundencia física, magia destructiva y liderazgo táctico. Vayan y aseguren ese pase a la final de cuatro. Yo guardaré mis mejores ganchos para el cierre.",
        "Hoshina lo miró de reojo y le apretó suavemente el brazo en un gesto de complicidad silenciosa.",
        "Erza, Ken y yo nos colocamos en el centro de la sala.",
        "La formación para el asalto de tres estaba sellada: el acero de la comandante, el fuego del fénix y el impacto demoledor de mis puños de tierra.",
        "El día de descanso había renovado nuestros cuerpos y fortalecido los lazos que nos unían como una verdadera familia.",
        "Y a la mañana siguiente, el Gran Coliseo Central temblaría ante el choque más violento que Solaria jamás había presenciado."
    ]
}

def compilar_capitulo_28():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 28              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25, CAPITULO_26, CAPITULO_27, CAPITULO_28
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 28
    txt_cap28 = os.path.join(capitulos_dir, "28_Capitulo_28.txt")
    with open(txt_cap28, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_28['header'].upper()}: {CAPITULO_28['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_28['paragraphs']))
    print(f"[OK] Texto plano Capítulo 28 guardado en:\n  {txt_cap28}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 28
    docx_cap28 = os.path.join(capitulos_dir, "28_Capitulo_28.docx")
    create_single_chapter_document(CAPITULO_28, docx_cap28, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 28 guardado en:\n  {docx_cap28}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 28)
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
    compilar_capitulo_28()
