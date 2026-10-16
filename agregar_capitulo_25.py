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

CAPITULO_25 = {
    "header": "Capítulo 25",
    "title": "El Peso de la Experiencia",
    "paragraphs": [
        "El amanecer del tercer día del Gran Torneo Interclases encontró a la capital de Solaria sumida en un fervor colectivo sin precedentes.",
        "Dos días de duelos individuales habían bastado para dinamitar la jerarquía intocable de la academia. La victoria abrasadora de Ken sobre Julian de Montfort y mi demoledor nocaut a puño limpio contra la 'Muralla de Acero' de la Clase B-1 habían puesto el nombre de la Clase F-2 en boca de todo el imperio. Los nobles ya no se reían en los pasillos; sus miradas estaban cargadas de un rencor ácido y de una urgencia desesperada por frenar nuestro avance.",
        "Este tercer día marcaba el cierre definitivo de la fase de duelos individuales uno contra uno antes de dar paso a los combates por equipos.",
        "Y la Clase A-1, herida en lo más hondo de su orgullo aristocrático, exigía sangre.",
        "* * *",
        "En los vestuarios subterráneos del coliseo, la atmósfera en nuestro campamento era de una calma absoluta.",
        "—Los bastardos de la A-1 están histéricos —comentó Natsuki, asomándose al corredor de piedra con una lista oficial en las manos—. Tras el ridículo que hizo Montfort el primer día, han decidido poner toda la carne en el asador para este último combate. No van a mandar a un cadete cualquiera. Han registrado a su peleador número uno.",
        "—Maximilian de Croix —dijo Erza, cruzando los brazos con una expresión severa—. El espadachín de la Orden del Halcón Carmesí. Jamás ha perdido un combate formal en la capital. Su técnica combina esgrima imperial de alta velocidad con maná de relámpago puro. Dicen que su espada se mueve más rápido que el parpadeo de un halcón.",
        "—Un niño mimado con medallas de oro y un sable reluciente —resopló Caesar King, sonriendo de lado mientras afilaba su hacha de combate—. Nada que nuestros muchachos no puedan quebrar.",
        "La mirada de todos los presentes se posó sobre el rincón más tranquilo de la habitación.",
        "Sentado sobre una caja de madera, Lighter se ajustaba con parsimonia las vendas de tela cruda alrededor de sus nudillos. Sus gafas de sol oscuras colgaban de su cuello y en sus labios descansaba una brizna de hierba seca. No había nervios en su rostro, ni tensión en sus hombros, ni la menor prisa.",
        "—Es tu turno, Lighter —le dije, apoyando una mano sobre su hombro fornido—. Eres el peleador más experimentado que tenemos. Ese noble cree que el mundo es un salón de baile.",
        "—El mundo es un callejón lleno de barro, Silva —respondió Lighter con su voz profunda y rasposa, apretando el último nudo de la venda con los dientes—. Y en el barro no importan los apellidos ni las capas de seda. Solo importa quién sabe encajar un golpe y seguir de pie.",
        "Subaru Hoshina dio dos pasos al frente desde la sombra del umbral.",
        "Llevaba las manos entrelazadas sobre su regazo y sus penetrantes ojos verdes observaban al boxeador con una intensidad silenciosa. Dio un paso más, quedando a escasos centímetros de él, y lo miró fijamente a los ojos sin vacilar.",
        "—No dejes que sus relámpagos te deslumbren —dijo Hoshina en un susurro cargado de una confianza sincera—. Recuerda lo que me enseñaste en el patio: mantén la base del pie, pivota la cadera y haz que se agote contra tu guardia. Sé que vas a ganar.",
        "Lighter la contempló durante dos segundos con una sonrisa templada y sincera. Se quitó la brizna de hierba de los labios, le dedicó un guiño cómplice y asintió con suavidad.",
        "—Cuenta con ello, señorita hielo. Le apagaré las luces en el primer asalto.",
        "* * *",
        "El rugido del coliseo al abrirse las compuertas de hierro fue ensordecedor.",
        "—¡EN LA ESQUINA NORTE! —bramó el anunciador imperial, haciendo temblar los cristales de las gradas—. ¡El prodigio invicto de la espada, heredero de la Orden del Halcón y orgullo máximo de la Clase A-1... Maximilian de Croix!",
        "La multitud aristocrática se puso de pie en masa, agitando pañuelos dorados y coreando su nombre.",
        "Por el túnel norte avanzó un joven alto, rubio y de facciones afiladas, envuelto en una capa carmesí de terciopelo y una coraza de plata bruñida. Su sable largo, forjado con aleación de electrum, crepitaba con chispas de electricidad azul y blanca que silbaban como víboras hambrientas.",
        "—¡Y EN LA ESQUINA SUR! —anunció la voz con evidente expectación—. ¡El contragolpeador de las Tierras Yermas, representando a la Clase F-2... Lighter!",
        "Lighter caminó hacia el centro de la arena con un paso perezoso, las manos ligeramente hundidas en los bolsillos de sus pantalones cortos de cuero y el torso cubierto únicamente por su camiseta gris sin mangas. No llevaba armas, ni armadura, ni artefactos mágicos.",
        "Solo dos puños vendados y la mirada tranquila de un hombre que había visto demasiada sangre en los caminos como para impresionarse por un puñado de fuegos artificiales.",
        "Maximilian lo esperó en el centro de la arena con la espada desenvainada, mirándolo con un asco infinito.",
        "—Ustedes, los salvajes de la F-2, han tenido demasiada suerte estos dos días —dijo el noble con voz gélida y arrogante—. Sorprendieron a Montfort con trucos baratos y vencieron a Goran porque era demasiado lento. Pero conmigo se acabaron los cuentos de hadas. Mi esgrima del Halcón Eléctrico no tiene puntos ciegos. Te voy a cortar en pedazos antes de que puedas pestañear.",
        "Lighter lo miró de arriba abajo con una calma casi paternal, soltando una exhalación corta por la nariz.",
        "—Tienes una boca muy rápida, muchacho —respondió Lighter, adelantando el pie izquierdo y levantando los puños a las mejillas en su clásica guardia de boxeo—. Vamos a ver si tus pies se mueven a la misma velocidad.",
        "—¡COMIENCEN! —rugió el árbitro.",
        "* * *",
        "Maximilian explotó con la velocidad de un rayo caído del cielo.",
        "¡BZ-ZZZT!",
        "Una estela de electricidad cegadora envolvió la arena cuando el noble cruzó diez metros en una milésima de segundo. Su espada trazó un corte diagonal descendente buscando rebanar el hombro de Lighter con filo relampagueante.",
        "La multitud contuvo el aliento esperando ver sangre volar.",
        "Pero Lighter ni siquiera parpadeó.",
        "Con un cabeceo milimétrico hacia la izquierda, dejó que la hoja electrificada pasara a escasos tres milímetros de su oreja. El viento del ataque le despeinó el flequillo, pero su balance corporal permaneció inalterable.",
        "Maximilian frunció el ceño, sorprendido por la esquiva, y encadenó una tormenta frenética de seis estocadas de alta frecuencia:",
        "¡SWOOSH! ¡ZAP! ¡SWOOSH! ¡ZAP!",
        "Lo que siguió en la arena fue una clase magistral de boxeo y esquiva pura.",
        "Lighter no retrocedía en línea recta como un novato. Se movía en círculos con un balanceo de torso fluido y rítmico, balanceándose de lado a lado como un péndulo de hierro. Deslizaba los pies sobre la arena con la suavidad que le había enseñado a Hoshina, utilizando el hombro izquierdo para desviar las trayectorias del sable y agachándose por debajo de los cortes horizontales con una soltura desconcertante.",
        "Para los nobles en las gradas, parecía que la espada de Maximilian atravesaba el aire vacío una y otra vez.",
        "—¡Quédate quieto, maldito perro callejero! —rugió Maximilian, perdiendo los estribos al ver que su técnica más letal no lograba rozar ni la tela de la camiseta de su rival.",
        "—Ese es tu problema, niño rico —habló Lighter en mitad del combate, esquivando un estoque con una flexión de cintura perfecta—. Estás tan acostumbrado a que tus rivales teman tu apellido que nunca aprendiste a leer los ojos de quien tienes enfrente.",
        "Lighter desvió la punta del sable con un golpe seco de su antebrazo vendado.",
        "—Tus ojos me dicen hacia dónde vas a tirar el golpe dos segundos antes de que muevas la muñeca —continuó con voz áspera y serena—. En las Tierras Yermas, un tipo con tu impaciencia y tu orgullo no dura vivo ni para el desayuno.",
        "* * *",
        "Las palabras de Lighter cayeron como una bofetada humillante sobre el noble.",
        "Ciego de rabia, Maximilian retrocedió tres pasos, clavó los talones en el suelo y liberó todo su maná de relámpago en una sobrecarga salvaje. Un aura eléctrica de color azul zafiro envolvió su cuerpo y su espada, haciendo crujir las piedras del suelo bajo una temperatura letal.",
        "—¡VOY A BORRAR ESA SONRISA DE TU ROSTRO! —bramó el noble—. ¡ARTE SUPREMA DEL HALCÓN: JUICIO DEL RAYO CEGADOR!",
        "Maximilian se arrojó hacia adelante a la velocidad máxima que su cuerpo podía soportar, proyectando un empuje frontal concentrado capaz de atravesar murallas de asedio.",
        "Era un ataque devastador. Pero era un ataque predecible, dictado por la ira y el desespero.",
        "Lighter no dio un solo paso atrás.",
        "Esperó hasta el último milisegundo, cuando la punta chispeante del sable estaba a escasos centímetros de su pecho.",
        "En ese instante exacto, ejecutó una rotación de hombro defensiva: giró el torso cuarenta y cinco grados, amortiguando la onda eléctrica con la solapa de su músculo deltoide, y entró con el pie derecho directo en el punto ciego del noble.",
        "El sable de Maximilian cortó el aire vacío una vez más, dejando su flanco derecho completamente desprotegido.",
        "—Fin del asalto —susurró Lighter.",
        "Plantó el talón en la arena con la firmeza de una roca. Giró la cadera con violencia controlada y descargó un gancho de izquierda demoledor directo a la zona hepática de Maximilian.",
        "¡PAAAF!",
        "El impacto al hígado fue un trueno seco y sordo.",
        "La respiración de Maximilian se detuvo en seco. Todo el flujo de maná eléctrico que envolvía su cuerpo se cortocircuitó y se apagó en una fracción de segundo, como una vela soplada por un huracán. Sus ojos se abrieron como platos, la mandíbula se le desencajó por la agonía visceral y sus piernas perdieron toda la fuerza.",
        "Pero Lighter no le dio tiempo de caer.",
        "Con un movimiento de cintura relámpago, cargó el peso sobre la pierna trasera y desató el contragolpe definitivo: un derechazo cruzado impecable, veloz como el rayo y pesado como el plomo, impactando con precisión quirúrgica en la punta de la mandíbula del noble.",
        "¡¡¡BOOOOOOM!!!",
        "El cerebro de Maximilian se desconectó instantáneamente.",
        "El sable de electrum salió disparado por los aires girando sobre sí mismo, mientras el cuerpo del campeón invicto de la Clase A-1 volaba dos metros hacia atrás antes de estrellarse contra la arena con un golpe inerte.",
        "* * *",
        "El silencio que envolvió al Gran Coliseo fue el más denso de los tres días de competencia.",
        "No hubo explosiones gigantescas, ni llamas cósmicas, ni cráteres descomunales.",
        "Había sido una demolición quirúrgica, fría y perfecta. Una lección magistral de cómo la experiencia real, la serenidad y los puños curtidos en la batalla de verdad pulverizaban la vanidad de quienes jamás habían conocido el peligro genuino.",
        "Lighter bajó la guardia con lentitud. Se frotó la barbilla, sacudió el polvo de sus nudillos vendados y soltó una bocanada de aire cansado pero satisfecho.",
        "Maximilian de Croix yacía boca arriba sobre la arena, con los ojos vidriosos fijos en el cielo y la boca entreabierta, completamente inconsciente.",
        "El árbitro tardó tres largos segundos en recuperar el habla antes de levantar el brazo derecho hacia el boxeador:",
        "—¡G-GANADOR POR NOCAUT FULMINANTE... LIGHTER, DE LA CLASE F-2!",
        "La arena explotó en un alarido de proporciones épicas.",
        "Las gradas de los plebeyos y las clases intermedias saltaban de júbilo incontrolable: la invicta Clase A-1 había sido humillada dos veces en duelos individuales por la misma clase de 'desechos'.",
        "En nuestro palco, la celebración fue una fiesta total.",
        "Caesar King soltaba carcajadas sonoras golpeando la barandilla con el puño; Lucy saltaba agitando su gorra; Burnice daba brincos alocados; y Ken, con su brazo vendado, vitoreaba con orgullo fraternal.",
        "Pero entre todo el alboroto, los ojos de Lighter buscaron un punto específico en las gradas.",
        "Allí estaba Subaru Hoshina.",
        "La chica de cabello plateado ya no mantenía los brazos cruzados ni ocultaba sus emociones bajo una máscara de hielo. De pie junto a Kaoruko, Hoshina contemplaba a Lighter con una sonrisa radiante, hermosa y desprovista de cualquier reserva.",
        "Levantó la mano derecha en un saludo claro y orgulloso, asintiendo con la cabeza en un gesto de reconocimiento total.",
        "Lighter sonrió de oreja a oreja en la arena, se llevó los dos dedos a la frente en un saludo militar despreocupado y comenzó a caminar hacia el túnel de vestuarios.",
        "La fase de duelos individuales había concluido.",
        "Con tres victorias arrolladoras en tres días consecutivos, la Clase F-2 se alzaba invicta como la fuerza más temible del torneo.",
        "Y ahora, los combates grupales de dos, tres y cuatro guerreros estaban listos para encender la arena una vez más."
    ]
}

def compilar_capitulo_25():
    print("==================================================")
    print("         COMPILACIÓN DEL CAPÍTULO 25              ")
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
    
    titulo = "El camino del reencarnado"
    subtitulo = "Crónicas de Sangre, Acero y Deseo"
    autor = "Seier Silva"
    capitulos = [
        PROLOGO, CAPITULO_1, CAPITULO_2, CAPITULO_3, CAPITULO_4,
        CAPITULO_5, CAPITULO_6, CAPITULO_7, CAPITULO_8, CAPITULO_9,
        CAPITULO_10, CAPITULO_11, CAPITULO_12, CAPITULO_13, CAPITULO_14,
        CAPITULO_15, CAPITULO_16, CAPITULO_17, CAPITULO_18, CAPITULO_19,
        CAPITULO_20, CAPITULO_21, CAPITULO_22, CAPITULO_23, CAPITULO_24,
        CAPITULO_25
    ]
    
    capitulos_dir = os.path.join(NOVELA_DIR, "capitulos")
    os.makedirs(capitulos_dir, exist_ok=True)
    
    # 1. Guardar texto plano del Capítulo 25
    txt_cap25 = os.path.join(capitulos_dir, "25_Capitulo_25.txt")
    with open(txt_cap25, "w", encoding="utf-8") as f:
        f.write(f"=== {CAPITULO_25['header'].upper()}: {CAPITULO_25['title'].upper()} ===\n\n")
        f.write("\n\n".join(CAPITULO_25['paragraphs']))
    print(f"[OK] Texto plano Capítulo 25 guardado en:\n  {txt_cap25}")
    
    # 2. Guardar Word (.docx) individual del Capítulo 25
    docx_cap25 = os.path.join(capitulos_dir, "25_Capitulo_25.docx")
    create_single_chapter_document(CAPITULO_25, docx_cap25, novel_title=titulo, author=autor)
    print(f"[OK] Word individual Capítulo 25 guardado en:\n  {docx_cap25}")
    
    # 3. Guardar Libro Maestro Completo (Prólogo + Cap 1 a 25)
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
    compilar_capitulo_25()
